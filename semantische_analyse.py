# ==========================================================
# Phase 2
# ==========================================================
import os
import json
import logging
import re
import time
import random
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional
from collections import defaultdict, Counter

import numpy as np
from sentence_transformers import SentenceTransformer
from keybert import KeyBERT
from sklearn.cluster import KMeans

# Versucht optional die HuggingFace-Pipeline zu importieren, falls das nicht klappt, wird später ohne automatisches Summarizing gearbeitet.
try:
    from transformers import pipeline as hf_pipeline
except Exception:
    hf_pipeline = None


# ==========================================================
# Konfiguration
# ==========================================================
# Modellname für semantische Embeddings
EMBED_MODEL_NAME = "paraphrase-multilingual-MiniLM-L12-v2"
# Modellname für automatische Zusammenfassungen
SUM_MODEL_NAME = "t5-small"
# Mindestanzahl an Wörtern, damit eine Zusammenfassung erzeugt wird
MIN_WORDS_FOR_SUMMARY = 40
# Fallback: Wie viele Wörter vom Originaltext gespeichert werden, wenn keine echte Zusammenfassung erstellt wird
SHORT_TEXT_FALLBACK_WORDS = 60
# Maximale Eingabelänge für das Summarizer-Modell
SUMMARY_MAX_CHARS = 800
# Parameter für die Mindest- und Maximallänge der Zusammenfassung
SUMMARY_MIN_LEN = 20
SUMMARY_MAX_LEN = 80
# Fester Seed für reproduzierbare Ergebnisse
RANDOM_SEED = 42


# ==========================================================
# Projektpfade
# ==========================================================
PROJECT_ROOT: Path = Path(__file__).resolve().parent

DATA_DIR: Path = PROJECT_ROOT / "data"
RAW_DIR: Path = DATA_DIR / "raw"
RAW_SLIDES_DIR: Path = RAW_DIR / "slides"
RAW_LOGS_DIR: Path = RAW_DIR / "logs"

PROCESSED_DIR: Path = DATA_DIR / "processed"

EXPORT_DIR: Path = PROJECT_ROOT / "export"
ASSETS_DIR: Path = EXPORT_DIR / "assets"

LOG_PATH: Path = RAW_LOGS_DIR / "pipeline_phase2.log"

# Ausgabe-Dateien der zweiten Pipeline-Phase
SEMANTIC_INDEX_PATH: Path = PROCESSED_DIR / "semantic_index.json"
SEGMENTS_INDEX_PATH: Path = PROCESSED_DIR / "segments_index.json"
GLOSSARY_PATH: Path = PROCESSED_DIR / "glossary.json"
TOPIC_HIERARCHY_PATH: Path = PROCESSED_DIR / "topic_hierarchy.json"
SUMMARIES_PATH: Path = PROCESSED_DIR / "summaries.json"
METRICS_PATH: Path = PROCESSED_DIR / "metrics.json"


# ==========================================================
# Stopwords: Englische Stopwörter für Keyword-Extraktion und kleine deutsche Stopwortliste für gemischte oder deutsche Texte
# ==========================================================
EN_STOPWORDS = {
    "the",
    "a",
    "an",
    "and",
    "or",
    "but",
    "if",
    "then",
    "else",
    "when",
    "while",
    "to",
    "of",
    "in",
    "on",
    "at",
    "for",
    "from",
    "with",
    "without",
    "as",
    "by",
    "is",
    "are",
    "was",
    "were",
    "be",
    "been",
    "being",
    "it",
    "this",
    "that",
    "these",
    "those",
    "we",
    "you",
    "they",
    "i",
    "he",
    "she",
    "them",
    "us",
    "our",
    "your",
    "their",
    "not",
    "no",
    "yes",
    "can",
    "could",
    "should",
    "would",
    "will",
    "may",
    "might",
    "must",
    "more",
    "most",
    "less",
    "least",
    "very",
    "also",
    "just",
    "about",
    "into",
    "over",
    "under",
}

DE_LIGHT = {
    "der",
    "die",
    "das",
    "und",
    "oder",
    "nicht",
    "mit",
    "für",
    "von",
    "im",
    "in",
    "am",
    "auf",
    "zu",
    "ein",
    "eine",
}


# ==========================================================
# Logging: Erstellt die benötigten Verzeichnisse für Logs und verarbeitete Dateien
# ==========================================================
def setup_logging() -> None:
    RAW_LOGS_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    # Root-Logger holen und auf INFO setzen
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # Vorhandene Handler entfernen, damit keine doppelten Logs entstehen
    for h in list(root_logger.handlers):
        root_logger.removeHandler(h)

    # Datei-Handler für das Logfile einrichten
    file_handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
    file_handler.setFormatter(
        logging.Formatter("%(levelname)s | %(asctime)s | %(message)s")
    )
    root_logger.addHandler(file_handler)


# ==========================================================
# Utils
# ==========================================================
# Setzt den Zufalls-Seed für Python, NumPy und Hashing, damit Ergebnisse möglichst reproduzierbar bleiben
def set_deterministic(seed: int = 42) -> None:
    random.seed(seed)
    np.random.seed(seed)
    os.environ["PYTHONHASHSEED"] = str(seed)


# Berechnet die Kosinus-Ähnlichkeit zwischen zwei Vektoren
def cosine_sim(a: np.ndarray, b: np.ndarray) -> float:
    na = np.linalg.norm(a)
    nb = np.linalg.norm(b)
    if na == 0.0 or nb == 0.0:
        return 0.0
    return float(np.dot(a, b) / (na * nb))


# Liest eine JSON-Datei sicher ein
def safe_read_json(path: str | Path) -> Dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# Schreibt Daten als JSON-Datei und erstellt bei Bedarf das Zielverzeichnis
def safe_write_json(path: str | Path, payload: Any) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

# Teilt Text in Abschnitte anhand leerer Zeilen
def split_paragraphs(text: str) -> List[str]:
    blocks = re.split(r"\n\s*\n+", (text or "").strip())
    return [b.strip() for b in blocks if b.strip()]

# Zerlegt Text in einzelne Sätze bzw. sinnvolle Texteinheiten
def split_sentences(text: str) -> List[str]:
    t = (text or "").strip()
    if not t:
        return []

    # Spezialzeichen und Bulletpoints vereinheitlichen
    t = t.replace("", "\n• ").replace("•", "\n• ")
    t = re.sub(r"\r\n?", "\n", t)
    parts = re.split(r"(?<=[.!?])\s+|\n+|(?<=;)\s+|(?<=:)\s+", t)

    cleaned: List[str] = []
    for p in parts:
        p = p.strip()
        # Entfernt führende Listenzeichen
        p = re.sub(r"^[\-\–\—•\*]+\s*", "", p)
        if p:
            cleaned.append(p)
    return cleaned

# Extrahiert Wörter aus einem Text
def tokenize_words(text: str) -> List[str]:
    return re.findall(r"\b\w+\b", (text or ""), flags=re.UNICODE)

# Berechnet einfache Metriken zur Textkomplexität
def text_complexity_metrics(text: str, sentences: List[str]) -> Dict[str, Any]:
    words = tokenize_words(text)
    word_count = len(words)
    sent_count = len(sentences) if sentences else 0

    avg_word_len = float(np.mean([len(w) for w in words])) if word_count else 0.0
    avg_sent_len_words = float(word_count / sent_count) if sent_count else 0.0

    uniq = len(set(w.lower() for w in words)) if word_count else 0
    ttr = float(uniq / word_count) if word_count else 0.0

    return {
        "wortanzahl": word_count,
        "satzanzahl": sent_count,
        "durchschnitt_wortlaenge": round(avg_word_len, 3),
        "durchschnitt_satzlaenge_woerter": round(avg_sent_len_words, 3),
        "type_token_ratio": round(ttr, 3),
    }

# Holt Überschrift und gesamten Textinhalt einer Folie
def extract_title_and_body(slide_data: Dict) -> Tuple[str, str]:
    title = (slide_data.get("überschrift") or "").strip()

    body_texts: List[str] = []
    for el in slide_data.get("elemente", []):
        if el.get("typ") == "text":
            t = (el.get("inhalt") or "").strip()
            if t:
                body_texts.append(t)

    body = "\n".join(body_texts).strip()
    return title, body

# Baut aus den wichtigsten Keywords ein kurzes Label
def top_keywords_label(keywords: List[str], max_terms: int = 4) -> str:
    kw = [k.strip() for k in keywords if k and k.strip()]
    return ", ".join(kw[:max_terms]) if kw else "ohne label"

# Erkennt grob, ob ein Text eher deutsch, englisch oder gemischt ist
def detect_language_mode(text: str) -> str:
    t = (text or "").lower()
    if not t.strip():
        return "mixed"

    de_strong = any(ch in t for ch in ["ä", "ö", "ü", "ß"])
    words = re.findall(r"\b[a-zäöüß]+\b", t)
    if not words:
        return "mixed"

    de_markers = {
        "der",
        "die",
        "das",
        "und",
        "oder",
        "nicht",
        "mit",
        "für",
        "von",
        "im",
        "in",
        "am",
        "auf",
        "zu",
        "ein",
        "eine",
    }
    en_markers = {
        "the",
        "and",
        "or",
        "not",
        "with",
        "for",
        "from",
        "in",
        "on",
        "at",
        "to",
        "a",
        "an",
        "of",
        "is",
        "are",
    }

    de_hits = sum(1 for w in words if w in de_markers)
    en_hits = sum(1 for w in words if w in en_markers)

    if de_hits >= 2 and en_hits == 0:
        return "de"
    if en_hits >= 2 and de_hits == 0 and not de_strong:
        return "en"
    return "mixed"

# Wählt passende Stoppwörter für KeyBERT je nach Sprache des Textes
def get_stopwords_for_keybert(text: str):
    mode = detect_language_mode(text)
    if mode == "de":
        return "german"
    if mode == "en":
        return list(EN_STOPWORDS)
    return list(EN_STOPWORDS.union(DE_LIGHT))


# ==========================================================
# Glossar: Baut ein Glossar aus den Keywords und Textsegmenten der Folien auf
# ==========================================================
def build_glossarx(
    slide_segments: Dict[str, Dict[str, List[str]]],
    slide_keywords: Dict[str, List[str]],
    min_occurrences_slides: int = 2,
    max_examples_per_term: int = 3,
    max_terms_total: int = 300,
    min_term_len: int = 3,
    allow_numbers: bool = False,
    stop_terms: Optional[set] = None,
    summarizer=None,
    definition_mode: str = "rule",
    definition_max_len_chars: int = 220,
) -> Dict[str, Any]:

    if stop_terms is None:
        stop_terms = set()

    stop_terms = set(x.lower().strip() for x in stop_terms) | {
        "etc",
        "z.b",
        "zb",
        "bzw",
        "u.a",
        "ua",
        "d.h",
        "dh",
        "fig",
        "abb",
        "table",
        "image",
        "slide",
        "kapitel",
        "inhalt",
    }

    # Entfernt doppelte Leerzeichen
    def _clean_spaces(t: str) -> str:
        return re.sub(r"\s+", " ", (t or "").strip())

    # Vereinheitlicht einen Glossarbegriff für spätere Vergleiche
    def _normalize_term(t: str) -> str:
        t = (t or "").strip()
        if not t:
            return ""

        t = t.replace("•", " ").replace("", " ").replace("\u00ad", "")
        t = t.strip("()[]{}")
        t = (
            t.replace("“", '"')
            .replace("”", '"')
            .replace("„", '"')
            .replace("’", "'")
            .replace("‘", "'")
        )
        t = _clean_spaces(t)
        t_low = t.lower()

        if not allow_numbers and re.fullmatch(r"[\d\W_]+", t_low or ""):
            return ""
        if len(t_low) < min_term_len:
            return ""
        return t_low

    # Wählt eine passende Darstellungsform für einen Begriff aus
    def _display_form(original_candidates: List[str], normalized: str) -> str:
        cands = []
        for c in original_candidates:
            if _normalize_term(c) == normalized:
                cands.append(c.strip())

        if not cands:
            return normalized.title() if " " in normalized else normalized

        def score(s: str) -> Tuple[int, int, int]:
            upp = sum(1 for ch in s if ch.isupper())
            low = sum(1 for ch in s if ch.islower())
            length = len(s)
            return (upp, -abs(upp - low), length)

        cands_sorted = sorted(cands, key=score, reverse=True)
        return cands_sorted[0]

    # Erzeugt einen Regex-Ausdruck, um einen Begriff im Text zu finden
    def _phrase_regex(term_norm: str) -> re.Pattern:
        tokens = term_norm.split()
        if len(tokens) == 1:
            pat = r"(?<!\w)" + re.escape(tokens[0]) + r"(?!\w)"
        else:
            joiner = r"(?:\s+|[-–—])+"
            pat = r"(?<!\w)" + joiner.join(re.escape(t) for t in tokens) + r"(?!\w)"
        return re.compile(pat, flags=re.IGNORECASE)

    # Speichert verschiedene Schreibweisen eines Begriffs
    term_variants: Dict[str, List[str]] = defaultdict(list)
    for fn, kws in slide_keywords.items():
        for k in kws or []:
            k = (k or "").strip()
            if not k:
                continue
            norm = _normalize_term(k)
            if not norm:
                continue
            if norm in stop_terms:
                continue
            term_variants[norm].append(k)

    if not term_variants:
        return {
            "meta": {
                "min_occurrences_slides": min_occurrences_slides,
                "max_examples_per_term": max_examples_per_term,
                "max_terms_total": max_terms_total,
                "definition_mode": definition_mode,
                "note": "Keine Kandidatenbegriffe aus KeyBERT gefunden.",
            },
            "glossar": {},
        }

    term_patterns: Dict[str, re.Pattern] = {
        t: _phrase_regex(t) for t in term_variants.keys()
    }

    term_to_slides = defaultdict(set)
    term_to_examples = defaultdict(list)
    term_to_count = Counter()

    # Durchsucht alle Sätze aller Folien nach Glossarbegriffen
    for fn in sorted(slide_segments.keys()):
        seg = slide_segments.get(fn, {}) or {}
        sentences = seg.get("saetze", []) or []

        for s in sentences:
            s_clean = _clean_spaces(s)
            if not s_clean:
                continue

            for term_norm, pat in term_patterns.items():
                if pat.search(s_clean):
                    term_to_slides[term_norm].add(fn)
                    term_to_count[term_norm] += 1
                    if len(term_to_examples[term_norm]) < max_examples_per_term:
                        term_to_examples[term_norm].append(s_clean)

    candidates = []
    for term_norm, slides in term_to_slides.items():
        slide_count = len(slides)
        if slide_count < min_occurrences_slides:
            continue
        score = (slide_count, int(term_to_count[term_norm]), -len(term_norm))
        candidates.append((score, term_norm))

    # Sortiert Begriffe nach Relevanz
    candidates.sort(key=lambda x: (x[0][0], x[0][1], x[0][2], x[1]), reverse=True)
    candidates = candidates[:max_terms_total]

    # Erstellt eine einfache Definition aus einem Beispielsatz
    def _rule_definition(term_display: str, examples: List[str]) -> str:
        if not examples:
            return ""
        ex = sorted(examples, key=len, reverse=True)[0].strip()
        cut = ex
        m = re.search(r"[.!?]", ex)
        if m:
            cut = ex[: m.start() + 1].strip()
        cut = cut[:definition_max_len_chars].strip()
        return f"{term_display}: {cut}" if cut else ""

    # Erstellt eine Definition mithilfe des T5-Modells
    def _t5_definition(term_display: str, examples: List[str]) -> str:
        if summarizer is None or not examples:
            return ""
        context = " ".join(examples[:2]).strip()
        if not context:
            return ""
        prompt = f"summarize: Definiere den Begriff '{term_display}' kurz anhand dieses Kontextes: {context}"
        prompt = prompt[:800]
        try:
            out = summarizer(
                prompt, max_length=60, min_length=10, do_sample=False, truncation=True
            )
            if isinstance(out, list) and out:
                txt = (out[0].get("summary_text") or "").strip()
                return txt[:definition_max_len_chars].strip()
        except Exception:
            return ""
        return ""

    # Erstellt alternative Schreibweisen
    def _aliases(term_norm: str) -> List[str]:
        als = set()
        if " " in term_norm:
            als.add(term_norm.replace(" ", "-"))
        return sorted(als)

    glossar: Dict[str, Any] = {}

    for _score, term_norm in candidates:
        slides_sorted = sorted(list(term_to_slides[term_norm]))
        exs = term_to_examples.get(term_norm, [])
        display = _display_form(term_variants.get(term_norm, []), term_norm)

        if definition_mode == "none":
            definition = ""
        elif definition_mode == "t5":
            definition = _t5_definition(display, exs) or _rule_definition(display, exs)
        else:
            definition = _rule_definition(display, exs)

        glossar[term_norm] = {
            "term": term_norm,
            "display": display,
            "definition": definition,
            "vorkommen_slides": slides_sorted,
            "haeufigkeit": int(term_to_count[term_norm]),
            "beispiele": exs,
            "aliases": _aliases(term_norm),
        }

    glossar = dict(sorted(glossar.items(), key=lambda x: x[0]))

    return {
        "meta": {
            "min_occurrences_slides": int(min_occurrences_slides),
            "max_examples_per_term": int(max_examples_per_term),
            "max_terms_total": int(max_terms_total),
            "min_term_len": int(min_term_len),
            "allow_numbers": bool(allow_numbers),
            "definition_mode": definition_mode,
        },
        "glossar": glossar,
    }


# ==========================================================
# Summaries
# ==========================================================
# Kürzt einen Text auf eine maximale Wortanzahl
def _truncate_words(text: str, max_words: int) -> str:
    words = tokenize_words(text)
    if len(words) <= max_words:
        return text.strip()
    parts = (text or "").strip().split()
    return " ".join(parts[:max_words]).strip()

# Erstellt eine Zusammenfassung einer Folie
def build_summary(summarizer, title: str, body: str, full_text: str) -> Dict[str, Any]:
    words = tokenize_words(full_text)
    word_count = len(words)

    # Fallback, wenn der Text zu kurz ist oder kein Summarizer verfügbar ist
    if word_count < MIN_WORDS_FOR_SUMMARY or summarizer is None:
        return {
            "modus": "fallback",
            "wortanzahl": word_count,
            "kurzer_originaltext": _truncate_words(
                full_text, SHORT_TEXT_FALLBACK_WORDS
            ),
            "summary": "",
            "modell": SUM_MODEL_NAME if summarizer is not None else "nicht_verfuegbar",
        }

    inp = f"summarize: {title}\n{body}".strip()
    inp = inp[:SUMMARY_MAX_CHARS]

    try:
        out = summarizer(
            inp,
            max_length=SUMMARY_MAX_LEN,
            min_length=SUMMARY_MIN_LEN,
            do_sample=False,
            truncation=True,
        )

        summary_text = ""
        if isinstance(out, list) and out:
            summary_text = (out[0].get("summary_text") or "").strip()

        if not summary_text:
            return {
                "modus": "fallback",
                "wortanzahl": word_count,
                "kurzer_originaltext": _truncate_words(
                    full_text, SHORT_TEXT_FALLBACK_WORDS
                ),
                "summary": "",
                "modell": SUM_MODEL_NAME,
            }

        return {
            "modus": "t5",
            "wortanzahl": word_count,
            "kurzer_originaltext": "",
            "summary": summary_text,
            "modell": SUM_MODEL_NAME,
        }

    except Exception as e:
        logging.exception(f"T5 Summarization fehlgeschlagen: {e}")
        return {
            "modus": "fallback",
            "wortanzahl": word_count,
            "kurzer_originaltext": _truncate_words(
                full_text, SHORT_TEXT_FALLBACK_WORDS
            ),
            "summary": "",
            "modell": SUM_MODEL_NAME,
            "fehler": str(e),
        }


# ====================================================================================
# Pipeline Klasse: Übernimmt die semantische Analyse der bereits erzeugte JSON-Folien
# ====================================================================================
class SemanticPipelinePhase2:
    def __init__(self):
        set_deterministic(RANDOM_SEED)

        logging.info("Phase 2 startet: Lade lokale Modelle...")
        logging.info(f"Projektpfad: {PROJECT_ROOT}")
        logging.info(f"Input-JSONs aus: {PROCESSED_DIR}")
        logging.info(f"Assets-Verzeichnis: {ASSETS_DIR}")
        logging.info(f"Logs nach: {RAW_LOGS_DIR}")

        self.embedder = SentenceTransformer(EMBED_MODEL_NAME)
        self.kw_model = KeyBERT(model=self.embedder)

        self.summarizer = None
        if hf_pipeline is None:
            logging.warning(
                "transformers nicht installiert -> keine T5-Zusammenfassungen möglich."
            )
        else:
            try:
                self.summarizer = hf_pipeline(
                    "summarization",
                    model=SUM_MODEL_NAME,
                    tokenizer=SUM_MODEL_NAME,
                    device=-1,
                )
                logging.info(f"T5 Summarizer geladen: {SUM_MODEL_NAME}")
            except Exception as e:
                logging.exception(f"T5 Summarizer konnte nicht geladen werden: {e}")
                self.summarizer = None

        logging.info("Modelle geladen")

    # Sucht alle JSON-Dateien der Folien aus Phase 1
    def _get_slide_files(self) -> List[str]:
        files: List[str] = []
        if not PROCESSED_DIR.is_dir():
            logging.warning(f"Processed-Verzeichnis nicht gefunden: {PROCESSED_DIR}")
            return files

        for fn in os.listdir(PROCESSED_DIR):
            if not fn.endswith(".json"):
                continue
            if fn.startswith("pptx_slide_") or fn.startswith("pdf_slide_"):
                files.append(fn)

        return sorted(files)

    # Hauptmethode für die semantische Verarbeitung aller Folien
    def process(self):
        start = time.time()
        slide_files = self._get_slide_files()
        logging.info(f"Analysiere {len(slide_files)} Slide-JSONs aus {PROCESSED_DIR}.")

        slide_vecs: List[np.ndarray] = []
        slide_names: List[str] = []

        segments_index: Dict[str, Any] = {"slides": {}}
        slide_segments: Dict[str, Dict[str, List[str]]] = {}
        slide_keywords: Dict[str, List[str]] = {}
        summaries_out: Dict[str, Any] = {"meta": {}, "slides": {}}

        global_complexity_accu = {"wortanzahl": 0, "satzanzahl": 0}
        global_words: List[str] = []

        for fn in slide_files:
            path = PROCESSED_DIR / fn
            data = safe_read_json(path)

            title, body = extract_title_and_body(data)
            full_text = (title + "\n" + body).strip()

            if not full_text:
                logging.info(f"Überspringe leere Folie: {fn}")
                continue

            # Text in Abschnitte und Sätze aufteilen
            paragraphs = split_paragraphs(full_text)
            sentences = split_sentences(full_text)
            slide_segments[fn] = {"abschnitte": paragraphs, "saetze": sentences}

            # Segmentinformationen speichern
            segments_index["slides"][fn] = {
                "foliennummer": data.get("foliennummer"),
                "ueberschrift": title,
                "quelle_json": str(path.relative_to(PROJECT_ROOT)),
                "abschnitte": paragraphs,
                "saetze": sentences,
                "counts": {
                    "abschnitte": len(paragraphs),
                    "saetze": len(sentences),
                },
            }

            # Textkomplexität berechnen
            complexity = text_complexity_metrics(full_text, sentences)
            global_complexity_accu["wortanzahl"] += complexity["wortanzahl"]
            global_complexity_accu["satzanzahl"] += complexity["satzanzahl"]
            global_words.extend(tokenize_words(full_text))

            # Keywords extrahieren
            stopw = get_stopwords_for_keybert(full_text)
            keywords_scored = self.kw_model.extract_keywords(
                full_text,
                keyphrase_ngram_range=(1, 2),
                stop_words=stopw,
                top_n=10,
            )
            keywords = [k for k, _score in keywords_scored]
            slide_keywords[fn] = keywords

            # Embeddings für Titel und Fließtext berechnen
            title_emb = self.embedder.encode(
                title if title else " ", convert_to_numpy=True
            )
            body_emb = self.embedder.encode(
                body if body else " ", convert_to_numpy=True
            )
            title_body_cosine = cosine_sim(title_emb, body_emb)

            slide_emb = (title_emb + body_emb) / 2.0
            slide_vecs.append(slide_emb)
            slide_names.append(fn)

            # Zusammenfassung erzeugen
            summary_obj = build_summary(
                summarizer=self.summarizer,
                title=title,
                body=body,
                full_text=full_text,
            )

            summaries_out["slides"][fn] = {
                "foliennummer": data.get("foliennummer"),
                "ueberschrift": title,
                "quelle_json": str(path.relative_to(PROJECT_ROOT)),
                **summary_obj,
            }

            # Semantische Informationen in die bestehende Folien-JSON schreiben
            data.setdefault("semantik", {})
            data["semantik"].update(
                {
                    "segmentierung": {
                        "abschnitte": paragraphs,
                        "saetze": sentences,
                    },
                    "textkomplexitaet": complexity,
                    "keywords": keywords,
                    "embeddings": {
                        "ueberschrift": title_emb.astype(float).tolist(),
                        "fliesstext": body_emb.astype(float).tolist(),
                    },
                    "cosine": {
                        "ueberschrift_vs_fliesstext": round(title_body_cosine, 6)
                    },
                    "language_mode": detect_language_mode(full_text),
                    "zusammenfassung": summary_obj,
                    "pfade": {
                        "json_datei": str(path.relative_to(PROJECT_ROOT)),
                        "assets_verzeichnis": (
                            str(ASSETS_DIR.relative_to(PROJECT_ROOT))
                            if ASSETS_DIR.exists()
                            else "export/assets"
                        ),
                    },
                }
            )
            # Segmentindex abspeichern
            safe_write_json(path, data)
            logging.info(f"Phase 2 Semantik geschrieben: {fn}")

        safe_write_json(SEGMENTS_INDEX_PATH, segments_index)

        summaries_out["meta"] = {
            "modell": SUM_MODEL_NAME,
            "min_words_for_summary": MIN_WORDS_FOR_SUMMARY,
            "fallback_short_text_words": SHORT_TEXT_FALLBACK_WORDS,
            "created_at_epoch": int(time.time()),
            "input_dir": str(PROCESSED_DIR.relative_to(PROJECT_ROOT)),
            "output_dir": str(PROCESSED_DIR.relative_to(PROJECT_ROOT)),
        }
        safe_write_json(SUMMARIES_PATH, summaries_out)

        # Globale Type-Token-Ratio berechnen
        global_ttr = 0.0
        if global_words:
            global_ttr = len(set(w.lower() for w in global_words)) / len(global_words)

        semantic_index: Dict[str, Any] = {
            "meta": {
                "input_dir": str(PROCESSED_DIR.relative_to(PROJECT_ROOT)),
                "output_dir": str(PROCESSED_DIR.relative_to(PROJECT_ROOT)),
                "assets_dir": (
                    str(ASSETS_DIR.relative_to(PROJECT_ROOT))
                    if ASSETS_DIR.exists()
                    else "export/assets"
                ),
            },
            "clusters": {},
            "neighbors": {},
            "global_textkomplexitaet": {
                "wortanzahl": int(global_complexity_accu["wortanzahl"]),
                "satzanzahl": int(global_complexity_accu["satzanzahl"]),
                "type_token_ratio": round(float(global_ttr), 3),
            },
        }

        topic_hierarchy: Dict[str, Any] = {"root": {"title": "Themen", "children": []}}

        # Clustering nur, wenn mindestens zwei Folien vorhanden sind
        if len(slide_vecs) >= 2:
            X = np.vstack(slide_vecs)
            k = min(5, len(slide_vecs))
            kmeans = KMeans(n_clusters=k, random_state=RANDOM_SEED, n_init=10).fit(X)

            clusters: Dict[int, List[str]] = {}
            for i, sfn in enumerate(slide_names):
                cid = int(kmeans.labels_[i])
                clusters.setdefault(cid, []).append(sfn)

            cluster_struct: Dict[str, Any] = {}
            for cid, fns in clusters.items():
                kw_counter = Counter()
                for sfn in fns:
                    for kw in slide_keywords.get(sfn, []):
                        kw_counter[kw.lower()] += 1
                # Cluster mit Top-Keywords beschriften
                top_cluster_kws = [t for t, _c in kw_counter.most_common(8)]
                label = top_keywords_label(top_cluster_kws, max_terms=4)
                centroid = kmeans.cluster_centers_[cid].astype(float).tolist()

                cluster_struct[str(cid)] = {
                    "label": label,
                    "top_keywords": top_cluster_kws,
                    "slides": fns,
                    "centroid": centroid,
                }

                topic_hierarchy["root"]["children"].append(
                    {
                        "type": "cluster",
                        "cluster_id": str(cid),
                        "title": label,
                        "children": [{"type": "slide", "file": sfn} for sfn in fns],
                    }
                )

            semantic_index["clusters"] = cluster_struct

            # Ähnlichste Nachbarfolien berechnen
            norms = np.linalg.norm(X, axis=1, keepdims=True)
            norms[norms == 0] = 1.0
            Xn = X / norms
            sim_matrix = Xn @ Xn.T

            for i, sfn in enumerate(slide_names):
                sims = sim_matrix[i].copy()
                sims[i] = -1.0
                top_idx = np.argsort(-sims)[:5]

                neighbors = []
                for j in top_idx:
                    if sims[j] <= 0:
                        continue
                    neighbors.append(
                        {
                            "file": slide_names[int(j)],
                            "cosine": round(float(sims[j]), 6),
                        }
                    )
                semantic_index["neighbors"][sfn] = neighbors

        safe_write_json(SEMANTIC_INDEX_PATH, semantic_index)
        safe_write_json(TOPIC_HIERARCHY_PATH, topic_hierarchy)

        # Glossar aus den gesammelten Begriffen bauen
        glossary_payload = build_glossarx(
            slide_segments=slide_segments,
            slide_keywords=slide_keywords,
            min_occurrences_slides=2,
            max_examples_per_term=3,
            max_terms_total=300,
            min_term_len=3,
            allow_numbers=False,
            stop_terms=set(),
            summarizer=self.summarizer,
            definition_mode="rule",
        )
        safe_write_json(GLOSSARY_PATH, glossary_payload)

        # Metriken-Datei aktualisieren
        metrics_existing = safe_read_json(METRICS_PATH) if METRICS_PATH.exists() else {}
        metrics_existing["phase2"] = {
            "slides_analysiert": len(slide_names),
            "input_dir": str(PROCESSED_DIR.relative_to(PROJECT_ROOT)),
            "output_dir": str(PROCESSED_DIR.relative_to(PROJECT_ROOT)),
            "logs_dir": str(RAW_LOGS_DIR.relative_to(PROJECT_ROOT)),
            "assets_dir": (
                str(ASSETS_DIR.relative_to(PROJECT_ROOT))
                if ASSETS_DIR.exists()
                else "export/assets"
            ),
            "semantic_index": SEMANTIC_INDEX_PATH.name,
            "segments_index": SEGMENTS_INDEX_PATH.name,
            "glossary": GLOSSARY_PATH.name,
            "summaries": SUMMARIES_PATH.name,
            "topic_hierarchy": TOPIC_HIERARCHY_PATH.name,
            "runtime_seconds": round(time.time() - start, 3),
        }
        safe_write_json(METRICS_PATH, metrics_existing)

        logging.info(f"Phase 2 fertig in {time.time() - start:.2f}s.")


# ==========================================================
# Public Runner: Setzt alle wichtigen Pfade neu anhand des Projektverzeichnisses
# ==========================================================
def run_phase2(project_root: Path) -> None:
    global PROJECT_ROOT, DATA_DIR, RAW_DIR, RAW_SLIDES_DIR, RAW_LOGS_DIR
    global PROCESSED_DIR, EXPORT_DIR, ASSETS_DIR
    global LOG_PATH, SEMANTIC_INDEX_PATH, SEGMENTS_INDEX_PATH, GLOSSARY_PATH
    global TOPIC_HIERARCHY_PATH, SUMMARIES_PATH, METRICS_PATH

    PROJECT_ROOT = project_root.resolve()

    DATA_DIR = PROJECT_ROOT / "data"
    RAW_DIR = DATA_DIR / "raw"
    RAW_SLIDES_DIR = RAW_DIR / "slides"
    RAW_LOGS_DIR = RAW_DIR / "logs"
    PROCESSED_DIR = DATA_DIR / "processed"
    EXPORT_DIR = PROJECT_ROOT / "export"
    ASSETS_DIR = EXPORT_DIR / "assets"
    LOG_PATH = RAW_LOGS_DIR / "pipeline_phase2.log"
    SEMANTIC_INDEX_PATH = PROCESSED_DIR / "semantic_index.json"
    SEGMENTS_INDEX_PATH = PROCESSED_DIR / "segments_index.json"
    GLOSSARY_PATH = PROCESSED_DIR / "glossary.json"
    TOPIC_HIERARCHY_PATH = PROCESSED_DIR / "topic_hierarchy.json"
    SUMMARIES_PATH = PROCESSED_DIR / "summaries.json"
    METRICS_PATH = PROCESSED_DIR / "metrics.json"

    setup_logging()

    pipeline = SemanticPipelinePhase2()
    pipeline.process()
    print("Phase 2 abgeschlossen. Siehe Logs:", LOG_PATH)


# ==========================================================
# Start
# ==========================================================
if __name__ == "__main__":
    run_phase2(Path(__file__).resolve().parent)
