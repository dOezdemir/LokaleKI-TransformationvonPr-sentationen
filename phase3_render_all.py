# ==========================================================
# Phase 3
# ==========================================================
import json
import re
import logging
import shutil
import hashlib
import html as html_lib
import time
from pathlib import Path
from typing import Optional, Tuple, Dict, List, Any

from jinja2 import Environment, FileSystemLoader, select_autoescape
from markupsafe import Markup


# ==========================================================
# Pfade: Projektpfad und wichtige Verzeichnisse definieren
# ==========================================================
PROJECT_ROOT = Path(__file__).resolve().parent
DATA_DIR = PROJECT_ROOT / "data"
RAW_DIR = DATA_DIR / "raw"
PROCESSED_DIR = DATA_DIR / "processed"
EXPORT_DIR = PROJECT_ROOT / "export"
SLIDES_DIR = EXPORT_DIR / "slides"
STYLES_DIR = EXPORT_DIR / "style"
EXPORT_ASSETS_DIR = EXPORT_DIR / "assets"
SOURCE_ASSETS_DIR = PROJECT_ROOT / "assets"
TEMPLATES_DIR = PROJECT_ROOT / "templates"
AUDIT_DIR = PROJECT_ROOT / "audit"
LOG_FILE = RAW_DIR / "logs" / "phase3_render.log"
SUMMARIES_PATH = PROCESSED_DIR / "summaries.json"
TOPIC_HIERARCHY_PATH = PROCESSED_DIR / "topic_hierarchy.json"


# ==========================================================
# Template-Namen automatisch auflösen
# ==========================================================
# Prüft, ob ein Template mit oder ohne .j2-Endung existiert
def resolve_template_name(base_name: str) -> str:
    candidates = [base_name]

    if base_name.endswith(".j2"):
        candidates.append(base_name[:-3])
    else:
        candidates.append(base_name + ".j2")

    for cand in candidates:
        if (TEMPLATES_DIR / cand).exists():
            return cand

    raise FileNotFoundError(
        f"Template nicht gefunden: {base_name}. "
        f"Gesucht in: {TEMPLATES_DIR}. "
        f"Kandidaten: {candidates}"
    )

# Namen der verwendeten Templates automatisch auflösen
TEMPLATE_SLIDE = resolve_template_name("slide.html")
TEMPLATE_INDEX = resolve_template_name("index.html")
TEMPLATE_GLOSSARY = resolve_template_name("glossary.html")
TEMPLATE_QUIZ = resolve_template_name("quiz.html")


# ==========================================================
# PDF Bullet Fix: Bullet-Prefix entfernen
# ==========================================================
BULLET_STRIP_RE = re.compile(r"^\s*([•\-\u2022\u25CF\u25AA\u25E6]|(\d+[\.\)]))\s+")

# Namen der verwendeten Templates automatisch auflösen
def _strip_bullet_prefix_text(s: str) -> str:
    if not isinstance(s, str):
        return ""
    return BULLET_STRIP_RE.sub("", s, count=1).lstrip()

# Entfernt Aufzählungszeichen aus einer Liste von Text-Runs
def _strip_bullet_prefix_runs(runs: list) -> list:
    if not isinstance(runs, list) or not runs:
        return runs

    new_runs = []
    removed = False

    for r in runs:
        if not isinstance(r, dict):
            continue

        t = r.get("text") or ""

        if not removed and isinstance(t, str) and t.strip():
            t2 = BULLET_STRIP_RE.sub("", t, count=1)
            if t2 != t:
                r2 = dict(r)
                r2["text"] = t2.lstrip()
                new_runs.append(r2)
                removed = True
                continue

            if t.strip() in {"•", "-", "–", "—"}:
                removed = True
                continue

            removed = True

        new_runs.append(r)

    if new_runs and isinstance(new_runs[0], dict):
        new_runs[0] = dict(new_runs[0])
        new_runs[0]["text"] = (new_runs[0].get("text") or "").lstrip()

    return new_runs


# ==========================================================
# I/O Helpers
# ==========================================================
# Liest eine JSON-Datei ein, falls sie existiert
def read_json(path: Path, default=None):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))

# Schreibt Text in eine Datei und erstellt fehlende Ordner automatisch
def write_text(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")

# Schreibt Text in eine Datei und erstellt fehlende Ordner automatisch
def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


# ==========================================================
# Logging für Phase 3 konfigurieren
# ==========================================================
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)


# ==========================================================
# Jinja: Custom Filter + Globals
# ==========================================================
# Prüft, ob ein Regex-Muster in einem Wert vorkommt
def jinja_regex_search(value: Any, pattern: str, flags: int = 0) -> str:
    try:
        s = "" if value is None else str(value)
        if re.search(pattern, s, flags=flags):
            return "true"
        return ""
    except Exception:
        return ""


SECTION_RE = re.compile(
    r"^\s*(\d{1,2}\s*[\.\)]?\s+|teil\s+\d+|kapitel\s+\d+)\b", re.IGNORECASE
)

# Prüft, ob ein Titel wie eine Abschnittsüberschrift aussieht
def is_section_title(title: Any) -> str:
    try:
        t = "" if title is None else str(title).strip()
        if not t:
            return ""
        return "true" if SECTION_RE.search(t) else ""
    except Exception:
        return ""


# ==========================================================
# Templates: Jinja2-Umgebung für HTML-Templates einrichten
# ==========================================================
env = Environment(
    loader=FileSystemLoader(str(TEMPLATES_DIR)),
    autoescape=select_autoescape(["html", "xml"]),
)

env.filters["regex_search"] = jinja_regex_search
env.globals["is_section_title"] = is_section_title

slide_tpl = env.get_template(TEMPLATE_SLIDE)
index_tpl = env.get_template(TEMPLATE_INDEX)
glossary_tpl = env.get_template(TEMPLATE_GLOSSARY)
quiz_tpl = env.get_template(TEMPLATE_QUIZ)


# ==============================================================
# Tooltips nur für Abkürzungen z.B. "Künstliche Intelligenz (KI)"
# ==============================================================
ABBR_PAIR_RE = re.compile(
    r"(?P<long>[A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß0-9][A-Za-zÄÖÜäöüß0-9 \-–—/]+?)\s*\((?P<abbr>[A-Za-z][A-Za-z0-9\-\/]{1,})\)"
)

# Prüft, ob ein Ausdruck wie eine Abkürzung aussieht
def _looks_like_abbreviation(abbr: str) -> bool:
    if not abbr or len(abbr) < 2:
        return False
    return any(ch.isupper() for ch in abbr)

# Extrahiert Abkürzungen und ihre Langformen aus einem Text
def _extract_abbrev_pairs_from_text(text: str) -> Dict[str, str]:
    out: Dict[str, str] = {}
    if not isinstance(text, str) or not text.strip():
        return out

    for m in ABBR_PAIR_RE.finditer(text):
        longf = re.sub(r"\s+", " ", (m.group("long") or "").strip())
        abbr = (m.group("abbr") or "").strip()

        if not _looks_like_abbreviation(abbr):
            continue
        if len(longf) < 3:
            continue
        if longf.upper() == abbr.upper():
            continue

        out[abbr] = longf

    return out

# Erstellt ein Mapping aller gefundenen Abkürzungen aus den JSON-Folien
def build_abbreviation_map(json_files: List[Path]) -> Dict[str, str]:
    abbr_map: Dict[str, str] = {}

    def add_from_text(t: str):
        for k, v in _extract_abbrev_pairs_from_text(t).items():
            abbr_map.setdefault(k, v)

    for jp in json_files:
        data = read_json(jp, default={}) or {}

        title = data.get("überschrift") or data.get("ueberschrift") or ""
        add_from_text(title)

        for el in data.get("elemente") or []:
            if not isinstance(el, dict):
                continue

            typ = el.get("typ")

            if typ == "text":
                add_from_text(el.get("inhalt") or "")
                for p in el.get("paragraphs") or []:
                    if not isinstance(p, dict):
                        continue
                    add_from_text(p.get("plain_text") or "")

            elif typ == "tabelle":
                rows = el.get("inhalt") or []
                if not isinstance(rows, list):
                    continue
                for row in rows:
                    if not isinstance(row, list):
                        continue
                    for cell in row:
                        if isinstance(cell, str):
                            add_from_text(cell)

    return abbr_map

# Erstellt ein Regex-Muster, das eine Abkürzung auch dann erkennt, wenn HTML-Tags dazwischen liegen
def _abbr_pattern_allowing_tags(abbr: str) -> re.Pattern:
    pieces = []
    for ch in abbr:
        pieces.append(re.escape(ch))
        pieces.append(r"(?:<[^>]+>)*")
    core = "".join(pieces[:-1])
    return re.compile(rf"(?<![\w])({core})(?![\w])")

# Fügt Tooltips für bekannte Abkürzungen in HTML-Inhalte ein
def inject_abbrev_tooltips_into_html(html: str, abbr_map: Dict[str, str]) -> str:
    if not html or not isinstance(html, str) or not abbr_map:
        return html

    keys = sorted(abbr_map.keys(), key=len, reverse=True)

    for abbr in keys:
        longf = abbr_map.get(abbr)
        if not longf:
            continue

        pat = _abbr_pattern_allowing_tags(abbr)
        html = pat.sub(
            lambda m: (
                f'<span class="tt" tabindex="0" data-tip="{html_lib.escape(longf)}">{m.group(1)}</span>'
            ),
            html,
        )

    return html


# ==========================================================
# Normalisierung / Helpers
# ==========================================================
# Gibt den Titel einer Folie zurück, ansonsten einen Fallback-Wert
def get_title(data: dict, fallback: str) -> str:
    t = (data.get("überschrift") or data.get("ueberschrift") or "").strip()
    return t if t else fallback

# Gibt den Titel einer Folie zurück, ansonsten einen Fallback-Wert
def pos_style(pos, cw: float, ch: float) -> str:
    if isinstance(pos, list):
        best = None
        best_area = -1.0
        for p in pos:
            if not isinstance(p, dict):
                continue
            w = float(p.get("breite") or 0)
            h = float(p.get("höhe") or 0)
            area = w * h
            if area > best_area:
                best_area = area
                best = p
        pos = best or {}

    left = (float(pos.get("links") or 0) / cw) * 100
    top = (float(pos.get("oben") or 0) / ch) * 100
    width = (float(pos.get("breite") or 0) / cw) * 100
    height = (float(pos.get("höhe") or 0) / ch) * 100
    return (
        f"left:{left:.4f}%; top:{top:.4f}%; width:{width:.4f}%; height:{height:.4f}%;"
    )

# Vereinheitlicht einen Hex-Farbwert
def _normalize_hex_color(s: str):
    if not isinstance(s, str):
        return None
    c = s.strip()
    if not c:
        return None
    if c.startswith("#"):
        c = c[1:]
    if len(c) == 8:
        c = c[-6:]
    if len(c) != 6:
        return None
    return "#" + c.upper()

# Vereinheitlicht einen Hex-Farbwert
def _color_from_el(el: dict):
    col = el.get("farbe")
    if isinstance(col, str) and col.strip():
        n = _normalize_hex_color(col)
        if n:
            return n
    return None

# Übersetzt verschiedene Align-Werte in CSS-Ausrichtungen
def _align_from_el(el: dict):
    a = el.get("textausrichtung")
    if not isinstance(a, str) or not a.strip():
        return None
    s = a.strip().lower()
    if "center" in s:
        return "center"
    if "right" in s:
        return "right"
    if "justify" in s:
        return "justify"
    if "left" in s:
        return "left"
    mapping = {
        "links": "left",
        "zentriert": "center",
        "rechts": "right",
        "block": "justify",
    }
    return mapping.get(s, s)

# Baut einen CSS-Style-String für Textelemente auf
def text_style_from_el(el: dict) -> str:
    parts = []
    col_css = _color_from_el(el)
    if col_css:
        parts.append(f"color:{col_css};")

    size = el.get("schriftgröße")
    if isinstance(size, (int, float)) and size > 0:
        parts.append(f"font-size:{float(size):.2f}px;")

    font = el.get("schriftart")
    if isinstance(font, str) and font.strip():
        parts.append(f"font-family:{font}, Arial, sans-serif;")

    align = _align_from_el(el)
    if align:
        parts.append(f"text-align:{align};")

    return "".join(parts)


# ==========================================================
# Teaser/Summary Auswahl + Bereinigung
# ==========================================================
SECTION_HEADING_RE = re.compile(
    r"^\s*(\d{1,2}\s*[\.\)]?\s+|teil\s+\d+|kapitel\s+\d+)\b", re.IGNORECASE
)

# Entfernt doppelte Wiederholungen am Satzanfang
def _dedupe_leading_phrase(t: str) -> str:
    if not t:
        return t
    s = t.strip()
    for _ in range(2):
        m = re.match(r"^((?:\S+\s+){0,5}\S+)\s+\1\b", s, flags=re.IGNORECASE)
        if m:
            s = s[m.end(1) :].lstrip()
        else:
            break
    return s

# Bereinigt Zusammenfassungen sprachlich und formal
def _clean_summary_text(text: str, *, title: str = "") -> str:
    if not isinstance(text, str):
        return ""
    t = text.strip()
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"\s+([.,;:!?])", r"\1", t)
    t = re.sub(r"\(\s+", "(", t)
    t = re.sub(r"\s+\)", ")", t)

    if title:
        ti = re.sub(r"\s+", " ", title).strip()
        if ti and t.lower().startswith(ti.lower()):
            t = t[len(ti) :].lstrip(" :-–—.").strip()

    t = _dedupe_leading_phrase(t)
    if t and t[0].isalpha() and t[0].islower():
        t = t[0].upper() + t[1:]
    return t.strip()

# Kürzt einen Vorschautext auf eine bestimmte Länge
def _clean_teaser(text: str, max_chars: int = 180) -> str:
    if not isinstance(text, str):
        return ""
    t = re.sub(r"\s+", " ", text).strip()
    if len(t) <= max_chars:
        return t
    return t[: max_chars - 1].rstrip() + "…"

# Kürzt einen Vorschautext auf eine bestimmte Länge
def teaser_from_summaries(
    summaries: dict,
    slide_json_name: str,
    *,
    fallback_title: str = "",
) -> Tuple[str, str]:
    if not isinstance(summaries, dict):
        return _clean_teaser(fallback_title), ""

    slides = summaries.get("slides") or {}
    entry = slides.get(slide_json_name) or {}
    if not isinstance(entry, dict):
        return _clean_teaser(fallback_title), ""

    modus = (entry.get("modus") or "").strip()
    summary = (entry.get("summary") or "").strip()
    kurz = (entry.get("kurzer_originaltext") or "").strip()
    title = (entry.get("ueberschrift") or fallback_title or "").strip()
    wc = int(entry.get("wortanzahl") or 0)

    is_divider = False
    if wc <= 16:
        is_divider = True
    if (title or "").strip().lower() == "kein titel gefunden":
        is_divider = True
    if title and SECTION_HEADING_RE.search(title):
        is_divider = True

    if summary:
        base = _clean_summary_text(summary, title=title)
        if base:
            return _clean_teaser(base), modus

    if kurz:
        base = _clean_summary_text(kurz, title=title)
        if base:
            return _clean_teaser(base), modus

    if is_divider and title:
        return _clean_teaser(title, max_chars=120), modus

    return _clean_teaser(title or fallback_title), modus


# ==========================================================
# Sortierung: Sortierschlüssel für PDF- und PPTX-Slides
# ==========================================================
def slide_sort_key(name: str):
    m = re.search(r"(pdf|pptx)_slide_(\d+)", name)
    if not m:
        return (9, 999999, name)
    src = m.group(1)
    num = int(m.group(2))
    src_order = 0 if src == "pdf" else 1
    return (src_order, num, name)


# ==========================================================
# Text Rendering: Rendert einen einzelnen Text-Run als HTML-Span
# ==========================================================
def _render_run_span(run: dict) -> str:
    text = html_lib.escape(str(run.get("text", "")))
    if not text:
        return ""

    style_parts = []

    col = run.get("farbe")
    if isinstance(col, str) and col.strip():
        n = _normalize_hex_color(col)
        if n:
            style_parts.append(f"color:{n};")

    sz = run.get("schriftgröße")
    if isinstance(sz, (int, float)) and sz > 0:
        style_parts.append(f"font-size:{float(sz):.2f}px;")

    fn = run.get("schriftart")
    if isinstance(fn, str) and fn.strip():
        style_parts.append(f"font-family:{fn}, Arial, sans-serif;")

    if run.get("bold") is True:
        text = f"<strong>{text}</strong>"
    if run.get("italic") is True:
        text = f"<em>{text}</em>"
    if run.get("underline") is True:
        text = f"<u>{text}</u>"

    style = "".join(style_parts)
    return f'<span style="{style}">{text}</span>' if style else f"<span>{text}</span>"

# Rendert Absätze und Listen eines Textelements als HTML
def render_paragraphs_with_runs(el: dict, *, is_pdf: bool = False) -> str:
    paragraphs = el.get("paragraphs")
    if not isinstance(paragraphs, list) or not paragraphs:
        return ""

    ps = []
    for p in paragraphs:
        if not isinstance(p, dict):
            continue
        plain = (p.get("plain_text") or "").strip()
        runs = p.get("runs") or []
        if plain or any(
            (isinstance(r, dict) and (r.get("text") or "").strip()) for r in runs
        ):
            ps.append(p)
    if not ps:
        return ""

    # Rendert den Inhalt eines einzelnen Absatzes
    def paragraph_inner_html(p: dict, is_bullet: bool) -> str:
        runs = p.get("runs") or []

        if is_bullet and isinstance(runs, list) and runs:
            runs = _strip_bullet_prefix_runs(runs)

        if isinstance(runs, list) and runs:
            out = []
            for r in runs:
                if isinstance(r, dict):
                    out.append(_render_run_span(r))
            html = "".join(out).strip()
            if html:
                html = re.sub(r"^\s*(•|-\s+|\d+[\.\)]\s+)", "", html, count=1).strip()
                return html

        txt = str(p.get("plain_text", ""))
        if is_bullet:
            txt = _strip_bullet_prefix_text(txt)
        return html_lib.escape(txt).strip()

    out = []
    in_list = False
    current_level = 0

    # Öffnet eine HTML-Liste
    def open_ul():
        nonlocal in_list
        out.append("<ul>")
        in_list = True

    # Schließt alle geöffneten Listen
    def close_ul_all():
        nonlocal in_list, current_level
        if not in_list:
            return
        while current_level > 0:
            out.append("</ul>")
            current_level -= 1
        out.append("</ul>")
        in_list = False

    # Passt die Verschachtelungstiefe von Listen an
    def ensure_level(target_level: int):
        nonlocal current_level
        if target_level < 0:
            target_level = 0
        while target_level > current_level:
            out.append("<ul>")
            current_level += 1
        while target_level < current_level:
            out.append("</ul>")
            current_level -= 1

    for p in ps:
        is_bullet = bool(p.get("is_bullet") or False)

        if not is_bullet:
            close_ul_all()
            html_line = paragraph_inner_html(p, is_bullet=False)
            if not html_line:
                continue
            if is_pdf:
                out.append(f"<p>{html_line}</p>")
            else:
                out.append(html_line)
                out.append("<br>")
            continue

        if not in_list:
            open_ul()
            current_level = 0

        lvl = int(p.get("level", 0) or 0)
        ensure_level(lvl)

        li_html = paragraph_inner_html(p, is_bullet=True)
        if not li_html:
            continue
        out.append(f"<li>{li_html}</li>")

    close_ul_all()

    if out and out[-1] == "<br>":
        out.pop()

    return "".join(out)

# Einfacher Fallback, wenn keine Absatz-/Run-Struktur vorhanden ist
def render_text_as_html_fallback(raw: str) -> str:
    raw = raw or ""
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    if not lines:
        return ""

    bullet_prefixes = ("•", "·", "‣", "◦", "▪", "-", "–", "—", "*", "○")

    def is_bullet_line(s: str) -> bool:
        s2 = s.lstrip()
        return (
            s2.startswith(bullet_prefixes) or re.match(r"^\d+[\.\)]\s+", s2) is not None
        )

    bullet_like = sum(1 for ln in lines if is_bullet_line(ln))
    looks_like_list = bullet_like >= 2 and bullet_like / max(len(lines), 1) >= 0.6

    if looks_like_list:
        items = [
            f"<li>{html_lib.escape(_strip_bullet_prefix_text(ln))}</li>" for ln in lines
        ]
        return "<ul>" + "".join(items) + "</ul>"

    return html_lib.escape("\n".join(lines)).replace("\n", "<br>")


# ==========================================================
# Cluster Infos:  Liest Clusterinformationen aus semantic_index.json
# ==========================================================
def build_cluster_tabs():
    semantic = read_json(PROCESSED_DIR / "semantic_index.json", default={}) or {}
    clusters = semantic.get("clusters", {}) or {}

    slide_to_cluster = {}
    cluster_labels = {}

    if isinstance(clusters, dict):
        for cid, cobj in clusters.items():
            cid = str(cid)
            cobj = cobj or {}
            cluster_labels[cid] = cobj.get("label") or f"Cluster {cid}"
            for s in cobj.get("slides") or []:
                slide_to_cluster[str(s)] = cid

    return slide_to_cluster, cluster_labels, clusters


# ==========================================================
# Slide Rendering
# ==========================================================
def render_slide(
    data: dict,
    *,
    title: str,
    slide_file: str,
    prev_href: Optional[str],
    next_href: Optional[str],
    slide_pos: int,
    slide_total: int,
    all_slides: List[str],
    abbr_map: Dict[str, str],
) -> Tuple[str, int]:
    is_pdf = slide_file.startswith("pdf_slide_") or slide_file.startswith("pdf_")

    cw = float(data["canvas"]["width"])
    ch = float(data["canvas"]["height"])
    aspect = cw / ch if ch else 1.0

    raw_elements = sorted(data.get("elemente", []), key=lambda e: e.get("z_index", 0))
    elements = []

    for el in raw_elements:
        typ = el.get("typ")

        pos = el.get("position", {})
        if typ == "diagramm_region":
            pos = el.get("bbox", {}) or {}

        style = "position:absolute;" + pos_style(pos, cw, ch)

        if typ == "diagramm_region":
            chart_img = el.get("chart_image")
            if chart_img:
                elements.append(
                    {
                        "typ": "chart_img",
                        "style": style,
                        "src": "../assets/" + chart_img,
                        "title": "Diagramm",
                    }
                )
            else:
                elements.append(
                    {
                        "typ": "diagramm_region",
                        "style": style,
                        "title": "Diagramm-Region",
                    }
                )

        elif typ == "text":
            html_text = ""
            if isinstance(el.get("paragraphs"), list) and el.get("paragraphs"):
                html_text = render_paragraphs_with_runs(el, is_pdf=is_pdf)
            if not html_text:
                html_text = render_text_as_html_fallback(str(el.get("inhalt", "")))

            html_text = inject_abbrev_tooltips_into_html(html_text, abbr_map)

            elements.append(
                {
                    "typ": "text",
                    "style": style + text_style_from_el(el),
                    "inhalt": Markup(html_text),
                }
            )

        elif typ == "bild":
            elements.append(
                {"typ": "bild", "style": style, "src": "../assets/" + el["datei"]}
            )

        elif typ == "tabelle":
            rows = el.get("inhalt", []) or []
            new_rows = []
            for row in rows:
                if not isinstance(row, list):
                    continue
                new_row = []
                for cell in row:
                    if isinstance(cell, str):
                        safe = html_lib.escape(cell)
                        safe = inject_abbrev_tooltips_into_html(safe, abbr_map)
                        new_row.append(Markup(safe))
                    else:
                        new_row.append(Markup(""))
                new_rows.append(new_row)

            elements.append({"typ": "tabelle", "style": style, "rows": new_rows})

    html = slide_tpl.render(
        title=title,
        aspect=f"{aspect:.6f}",
        elements=elements,
        canvas_w=cw,
        canvas_h=ch,
        slide_file=slide_file,
        is_pdf=is_pdf,
        prev_href=prev_href,
        next_href=next_href,
        slide_pos=slide_pos,
        slide_total=slide_total,
        all_slides_json=json.dumps(all_slides, ensure_ascii=False),
    )
    return html, len(elements)


# ==========================================================
# Index Rendering: Rendert die Übersichtsseite
# ==========================================================
def build_index(*, items, topic_hierarchy, slide_meta_by_json):
    return index_tpl.render(
        items=items,
        topic_hierarchy=topic_hierarchy,
        slide_meta_by_json=slide_meta_by_json,
    )


# ==========================================================
# Quiz Generierung: Erstellt ein einfaches Lückentext-Quiz aus den Folieninhalten
# ==========================================================
def generate_final_quiz():
    segments_path = PROCESSED_DIR / "segments_index.json"
    semantic_path = PROCESSED_DIR / "semantic_index.json"
    if not segments_path.exists() or not semantic_path.exists():
        logging.warning(
            "Quiz nicht erzeugt – segments_index.json oder semantic_index.json fehlt."
        )
        return

    segments = read_json(segments_path, default={}) or {}
    semantic = read_json(semantic_path, default={}) or {}

    slides = segments.get("slides", {}) or {}
    clusters_raw = semantic.get("clusters", {}) or {}

    slide_to_cluster = {}
    clusters_by_id = {}

    if isinstance(clusters_raw, dict):
        for cid, cobj in clusters_raw.items():
            cid = str(cid)
            cobj = cobj or {}
            clusters_by_id[cid] = cobj
            for s in cobj.get("slides", []) or []:
                slide_to_cluster[str(s)] = cid

    def valid(t: str) -> bool:
        t = (t or "").strip()
        if not t:
            return False
        low = t.lower()
        if low.startswith("pdf seite") or low == "kein titel gefunden" or t.isdigit():
            return False
        return True

    def get_sentences(slide_data: dict) -> List[str]:
        s = slide_data.get("saetze")
        if not isinstance(s, list):
            s = slide_data.get("sätze")
        if not isinstance(s, list):
            return []
        return [x.strip() for x in s if isinstance(x, str) and valid(x.strip())]

    def pick_sentence(slide_data: dict):
        sentences = get_sentences(slide_data)
        if not sentences:
            return None
        longish = [s for s in sentences if len(s) >= 40]
        return max(longish or sentences, key=len)

    def normalize_word(w: str) -> str:
        return re.sub(r"\s+", " ", (w or "").strip())

    STOP = {
        "diese",
        "dieser",
        "dieses",
        "werden",
        "wird",
        "können",
        "kann",
        "damit",
        "dabei",
        "weil",
        "aber",
    }

    def find_blank_word(sentence: str, keywords: List[str]):
        s_low = sentence.lower()
        kws = [normalize_word(k) for k in (keywords or []) if isinstance(k, str)]
        kws = [k for k in kws if len(k) >= 3]
        kws.sort(key=len, reverse=True)
        for kw in kws:
            if kw.lower() in s_low:
                return kw
        words = re.findall(r"[A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß\-]{3,}", sentence)
        if not words:
            return None
        words2 = [w for w in words if w.lower() not in STOP]
        return max(words2 or words, key=len)

    def make_cloze(sentence: str, answer: str) -> str:
        return re.compile(re.escape(answer), re.IGNORECASE).sub(
            "____", sentence, count=1
        )

    quiz = []
    for slide_file, slide_data in (slides or {}).items():
        slide_data = slide_data or {}
        sentence = pick_sentence(slide_data)
        if not sentence:
            continue

        cid = slide_to_cluster.get(slide_file)
        cluster_obj = clusters_by_id.get(str(cid), {}) if cid is not None else {}
        answer_word = find_blank_word(
            sentence, cluster_obj.get("top_keywords", []) or []
        )
        if not answer_word:
            continue

        cloze = make_cloze(sentence, answer_word)
        if "____" not in cloze:
            continue

        quiz.append(
            {
                "slide": f"slides/{slide_file.replace('.json', '.html')}",
                "question": cloze,
                "answer": answer_word,
            }
        )

    write_text(AUDIT_DIR / "quiz.json", json.dumps(quiz, indent=2, ensure_ascii=False))
    quiz_data_inline = json.dumps(quiz, ensure_ascii=False)
    write_text(
        EXPORT_DIR / "quiz.html", quiz_tpl.render(quiz_data_inline=quiz_data_inline)
    )
    logging.info(f"Cloze-Quiz generiert mit {len(quiz)} Fragen.")


# ==========================================================
# Audit Snapshot (Run-Datei)
# ==========================================================
# Erstellt eine Audit-Datei mit allen erzeugten Artefakten und deren Hashes
def write_run_audit(run_id: str, params: dict):
    AUDIT_DIR.mkdir(parents=True, exist_ok=True)
    artifacts = []

    for p in sorted(EXPORT_DIR.rglob("*")):
        if p.is_file():
            artifacts.append({"path": p.as_posix(), "sha256": sha256_file(p)})

    payload = {
        "run_id": run_id,
        "created_at_epoch": int(time.time()),
        "params": params,
        "artifacts": artifacts,
    }
    out = AUDIT_DIR / f"run_{run_id}.json"
    write_text(out, json.dumps(payload, indent=2, ensure_ascii=False))
    logging.info(f"Run-Audit geschrieben: {out.as_posix()} | Dateien: {len(artifacts)}")


# ==========================================================
# MAIN
# ==========================================================
def main():
    logging.info("=== Phase 3 gestartet ===")

    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    SLIDES_DIR.mkdir(parents=True, exist_ok=True)
    STYLES_DIR.mkdir(parents=True, exist_ok=True)
    EXPORT_ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    AUDIT_DIR.mkdir(parents=True, exist_ok=True)
    # CSS-Datei in den Export kopieren
    css_src = SOURCE_ASSETS_DIR / "style.css"
    if css_src.exists():
        shutil.copy2(css_src, STYLES_DIR / "style.css")
        logging.info(f"CSS kopiert von {css_src} nach {STYLES_DIR / 'style.css'}")
    else:
        logging.warning(f"style.css nicht gefunden: {css_src}")

    _slide_to_cluster, _cluster_labels, _clusters = build_cluster_tabs()

    # Alle Folien-JSON-Dateien laden
    json_files = sorted(
        PROCESSED_DIR.glob("*_slide_*.json"), key=lambda p: slide_sort_key(p.name)
    )
    if not json_files:
        print("Keine *_slide_*.json Dateien gefunden.")
        return

    slide_total = len(json_files)
    all_slide_html_names = [p.stem + ".html" for p in json_files]

    summaries = read_json(SUMMARIES_PATH, default={}) or {}
    topic_hierarchy = read_json(TOPIC_HIERARCHY_PATH, default={}) or {}

    # Abkürzungen aus allen Folien sammeln
    abbr_map = build_abbreviation_map(json_files)
    logging.info(f"Abkürzungs-Mapping gebaut: {len(abbr_map)} Einträge")

    index_items = []
    slide_meta_by_json = {}

    for i, jp in enumerate(json_files):
        data = read_json(jp, default={}) or {}
        slide_html_name = jp.stem + ".html"
        slide_title = get_title(data, jp.stem)

        prev_html = (json_files[i - 1].stem + ".html") if i > 0 else None
        next_html = (json_files[i + 1].stem + ".html") if i < slide_total - 1 else None

        # Einzelne Folie rendern
        html, n = render_slide(
            data,
            title=slide_title,
            slide_file=slide_html_name,
            prev_href=prev_html,
            next_href=next_html,
            slide_pos=i + 1,
            slide_total=slide_total,
            all_slides=all_slide_html_names,
            abbr_map=abbr_map,
        )

        write_text(SLIDES_DIR / slide_html_name, html)

        # Teaser für die Übersichtsseite erzeugen
        teaser, modus = teaser_from_summaries(
            summaries, jp.name, fallback_title=slide_title
        )

        slide_meta_by_json[jp.name] = {
            "json_file": jp.name,
            "filename": f"slides/{slide_html_name}",
            "title": slide_title,
            "teaser": teaser,
            "modus": modus,
        }

        index_items.append(
            {
                "filename": f"slides/{slide_html_name}",
                "title": slide_title,
                "teaser": teaser,
                "modus": modus,
            }
        )

        logging.info(f"{slide_html_name} erzeugt | Elemente: {n}")
    # Quiz-Seite erzeugen
    generate_final_quiz()

    # Index-Seite schreiben
    write_text(
        EXPORT_DIR / "index.html",
        build_index(
            items=index_items,
            topic_hierarchy=topic_hierarchy,
            slide_meta_by_json=slide_meta_by_json,
        ),
    )
    # Glossar-Seite erzeugen, falls glossary.json vorhanden ist
    glossary_path = PROCESSED_DIR / "glossary.json"
    glossary_data = read_json(glossary_path, default=None)
    if glossary_data:
        glossar_dict = glossary_data.get("glossar", {}) or {}
        term_objects = []

        for _, obj in glossar_dict.items():
            if not isinstance(obj, dict):
                continue

            slides_json = obj.get("vorkommen_slides", []) or []
            if not isinstance(slides_json, list):
                slides_json = []

            display = (obj.get("display") or obj.get("term") or "").strip()
            definition = (obj.get("definition") or "").strip()

            examples = obj.get("beispiele", []) or []
            if not isinstance(examples, list):
                examples = []

            term_objects.append(
                {
                    "name": display,
                    "term": (obj.get("term") or "").strip(),
                    "definition": definition,
                    "slides": [
                        f"slides/{s.replace('.json', '.html')}" for s in slides_json
                    ],
                    "examples": examples,
                    "count": int(obj.get("haeufigkeit") or 0),
                }
            )

        term_objects.sort(key=lambda x: (x.get("name") or "").lower())
        write_text(EXPORT_DIR / "glossar.html", glossary_tpl.render(terms=term_objects))
        logging.info(f"Glossar generiert mit {len(term_objects)} Begriffen.")

    # map.json zusätzlich ins Audit-Verzeichnis kopieren
    map_src = PROCESSED_DIR / "map.json"
    if map_src.exists():
        shutil.copy2(map_src, AUDIT_DIR / "map.json")
        logging.info(f"map.json nach Audit kopiert: {AUDIT_DIR / 'map.json'}")
   
    # Run-Audit schreiben
    run_id = time.strftime("%Y_%m_%d_%H%M%S")
    write_run_audit(
        run_id=run_id,
        params={
            "phase": 3,
            "templates_dir": str(TEMPLATES_DIR),
            "template_slide": TEMPLATE_SLIDE,
            "template_index": TEMPLATE_INDEX,
            "template_glossary": TEMPLATE_GLOSSARY,
            "template_quiz": TEMPLATE_QUIZ,
            "processed_dir": str(PROCESSED_DIR),
            "export_dir": str(EXPORT_DIR),
            "audit_dir": str(AUDIT_DIR),
            "source_assets_dir": str(SOURCE_ASSETS_DIR),
            "styles_dir": str(STYLES_DIR),
        },
    )

    logging.info("=== Phase 3 abgeschlossen ===")
    print(f"OK: {len(index_items)} Seiten gelistet nach: {EXPORT_DIR.resolve()}")
    print("Öffne: export/index.html")


if __name__ == "__main__":
    main()
