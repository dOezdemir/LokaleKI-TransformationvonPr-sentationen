# Lokale KI – Transformation von Präsentationen

Dieses Projekt verarbeitet **PowerPoint- und PDF-Präsentationen lokal** und wandelt sie in eine strukturierte, analysierbare und browserbasierte Darstellung um.  
Dabei werden Inhalte aus Folien extrahiert, semantisch analysiert und als HTML-Ausgabe mit Zusatzfunktionen bereitgestellt.

## Ziel des Projekts

Das Tool hilft dabei, Präsentationen nicht nur anzuzeigen, sondern auch **inhaltlich aufzubereiten**:

- Extraktion von Texten, Bildern, Tabellen und Diagrammen
- semantische Analyse der Folieninhalte
- Erstellung von Zusammenfassungen und Schlagwörtern
- Generierung einer navigierbaren HTML-Version
- Aufbau eines Glossars und eines Quiz
- nachvollziehbare Verarbeitung durch Logs und Audit-Hashes

---

## Funktionen

### Phase 1 – Extraktion der Präsentationsinhalte
Das Tool liest **PPTX- und PDF-Dateien** ein und speichert die Inhalte als strukturierte JSON-Dateien.

#### Für PPTX-Dateien:
- Extraktion von:
  - Titeln
  - Textfeldern
  - Absätzen und Textformatierungen
  - Bildern
  - Tabellen
  - Diagrammen
  - Gruppenobjekten
- Speicherung der Folienstruktur in `data/processed/*.json`
- Export eingebetteter Bilder nach `export/assets/`

#### Für PDF-Dateien:
- Extraktion von:
  - Textblöcken
  - vermuteten Überschriften
  - Bildern
  - erkannten Diagrammregionen
- Speicherung pro Seite als JSON
- Export von Bildern und Diagramm-Ausschnitten

#### Zusätzlich:
- Audit-Logging mit SHA-256-Hashes
- technische Metadaten zur Laufzeit und Umgebung

---

### Phase 2 – Semantische Analyse
In der zweiten Phase werden die erzeugten JSON-Dateien inhaltlich analysiert.

#### Enthaltene Analysen:
- Satz- und Absatzsegmentierung
- Textkomplexitätsmetriken
- Keyword-Extraktion
- semantische Embeddings
- Ähnlichkeitsberechnung zwischen Folien
- Clusterbildung thematisch ähnlicher Folien
- automatische bzw. fallback-basierte Zusammenfassungen
- Glossarerstellung
- Themenhierarchie für die Navigation

#### Wichtige Ausgabedateien:
- `semantic_index.json`
- `segments_index.json`
- `glossary.json`
- `summaries.json`
- `topic_hierarchy.json`
- `metrics.json`

---

### Phase 3 – HTML-Rendering
Aus den strukturierten und analysierten Daten wird eine klickbare HTML-Ausgabe erzeugt.

#### Generierte Seiten:
- `export/index.html` – Startseite / Übersicht
- `export/slides/*.html` – einzelne Folienseiten
- `export/glossar.html` – Glossar
- `export/quiz.html` – Quiz

#### Weitere Features:
- Navigation zwischen Folien
- Teasertexte für die Startseite
- Themencluster
- Tooltip-Unterstützung für Abkürzungen
- Zwischen den Folien springen
- CSS-basierte Darstellung im Browser

---

Für Tester*innen: So benutzt du das Tool
Schnellstart

- Lege eine PPTX-Datei und/oder eine PDF-Datei in den Ordner data/raw/slides_input/.
- Starte zuerst die Extraktion mit main.py.
- Starte danach die semantische Analyse mit semantische_analyse.py.
- Starte zuletzt das Rendering mit phase3_render_all.py.
- Öffne export/index.html im Browser.
  
---

## Projektstruktur

```text
.
├── main.py
├── semantische_analyse.py
├── phase3_render_all.py
├── templates/
│   ├── slide.html.j2
│   ├── index.html.j2
│   ├── glossary.html.j2
│   └── quiz.html.j2
├── assets/
│   └── style.css
├── data/
│   ├── raw/
│   │   ├── slides_input/
│   │   └── logs/
│   └── processed/
├── export/
│   ├── index.html
│   ├── glossar.html
│   ├── quiz.html
│   ├── slides/
│   └── assets/
└── audit/

