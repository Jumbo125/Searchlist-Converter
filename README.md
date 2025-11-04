# CSV → A4 Tabelle (PNG/JPG/PDF/XLSX) — README, Kurzanleitung & Code-Analyse


Dieses Dokument enthält:
1) 🇩🇪 README (Deutsch)
2) 🇬🇧 README (English)
3) Kurzanleitung (DE)
4) Code-Analyse & Verbesserungsvorschläge (kurz & konkret)
5) requirements.txt (Vorschlag) & Packaging-Hinweise

-------------------------------------------------------------------------------
1) 🇩🇪 README (Deutsch)
-------------------------------------------------------------------------------

# CSV → A4 Tabelle (PNG/JPG/PDF/XLSX)

Konvertiert CSV-Dateien **oder** Cochrane Search Manager-TXT in sauber gesetzte A4-Tabellen
als **PNG**, **JPG**, **PDF** oder **Excel (XLSX)**. Mit Zebra-Zeilen, anpassbaren Header-
farben, robuster Spaltenbreitenverteilung, zuverlässigem Textumbruch (inkl. optionaler
Silbentrennung in Headern) und Export auf mehrere Seiten.

## Highlights
- Eingaben: **CSV** und **Cochrane Search Manager .txt** (Spalten „ID / Search / Hits“)
- Ausgaben: **PNG**, **JPG**, **PDF** (mehrseitig) oder **XLSX**
- A4 mit 300 DPI, wahlweise **Hoch-/Querformat**
- **Zebra-Zeilen**, **Header-Farbe** aus Presets oder frei wählbar
- **Robuster Umbruch**: passt auch sehr lange Tokens an; Header optional mit **Silbentrennung**
- **Automatische Spaltenbreiten** mit Mindestbreite, natürlicher Breite und „Puffer“
- **UTF‑8-Umschaltung** sowie Presets für EBSCO/PubMed-CSV (Trennzeichen)
- **Option „Leere Spalten entfernen“** (Body-only) über temporäre bereinigte CSV
- **Excel-Export** mit Drucktitelzeile, Umbruch, Ränder, A4, Freeze Panes

## Systemvoraussetzungen
- Python **3.9+** (Windows, macOS, Linux)
- Abhängigkeiten: `Pillow`, `openpyxl`, `pyphen` (optional), `tkinter` (Standard bei CPython)
- Systemschriftarten (z. B. Segoe UI / Arial / DejaVu Sans / Helvetica)

## Installation
```bash
python -m venv .venv
# Windows:
.venv\Scripts\pip install -U pip
.venv\Scripts\pip install -r requirements.txt
# macOS/Linux:
source .venv/bin/activate
pip install -U pip
pip install -r requirements.txt
```
Falls du (noch) keine `requirements.txt` hast: siehe Abschnitt 5 unten.

## Start
```bash
# innerhalb des aktivierten venv
python your_script.py
```
Ein GUI-Fenster startet: **„CSV → A4 Tabelle (PNG/JPG/PDF/XLSX)“**.

## Bedienung (GUI)
1. **Datei wählen**: CSV oder Cochrane-TXT.
2. **Zieldatei**: Speicherort & Name festlegen.
3. **Ausgabeformat**: PNG, JPG, PDF oder XLSX.
4. **Ausrichtung**: Hochformat oder Querformat.
5. **Farben**: Header- & Zebra-Farbe aus Presets wählen oder „Benutzerdefiniert…“.
6. **Trennzeichen**: `,` `;` `Tab` `|` „Benutzerdefiniert“ oder Presets **EBSCO (,)** / **PubMed (,)**.
7. **UTF‑8 korrekt darstellen**: aktivieren für Umlaute etc. (oder deaktivieren für cp1252).
8. **Silbentrennung (Header)**: Auto (DE/EN), de_DE, en_US oder Aus.
9. **Leere Spalten aus CSV entfernen**: entfernt Body-Only-leere Spalten via temporärer CSV.
10. **Erstellen**: Export als Bild(er)/PDF/XLSX. Mehrseitige PNG/JPG werden _base_01, _base_02 … benannt.

## Eingabedetails
- **CSV**: wird per `csv.reader` mit gewähltem Trennzeichen eingelesen.
- **Cochrane TXT**: robustes Parsen von „ID / Search / Hits“, mehrzeilige Queries inkl.
  *Meta-Feld* `Date Run` wird als einzeiliger Hinweis oberhalb der Tabelle ausgegeben.

## Ausgabedetails
- **PDF**: 1..n Seiten, 300 DPI, A4.
- **PNG/JPG**: bei mehreren Seiten nummerierte Dateien (`_01`, `_02`, …).
- **XLSX**: Auto-Spaltenbreiten, Umbruch, dünne Rahmen, A4, Quer-/Hochformat, Drucktitelzeile.
  `Freeze Panes` ab erster Datenzeile.

## Textumbruch & Spaltenbreiten (Kurz erklärt)
- Mindestbreite je Spalte = max(„3‑Zeichen‑Floor“, längstes **Header-Teilstück** mit harter Chunk‑Größe).
- **Header-Hard-Wrap** (Standard 5) verhindert zu breite Header ohne Leerzeichen.
- **Silbentrennung** (nur Header): via `pyphen` (optional), Auto-Erkennung DE/EN aus Umlauten.
- **Pufferlogik**: Extra-Breite erhält vorrangig, was am meisten Umbrüche spart.

## Bekannte Grenzen
- Sehr breite Tabellen: Schrift wird im Header graduell reduziert (bis Min-Headergröße).
- Schriftarten: Fallback auf `ImageFont.load_default()` wenn Systemfont fehlt.
- CSV-Sonderfälle (eingebettete Trennzeichen/Zeilenumbrüche) hängen von korrekter CSV-Form ab.
- Großes CSV ⇒ rechenintensiver Zeilenhöhen‑Scan; Geduld bei sehr langen Tabellen.

## Fehlerbilder & Hinweise
- „Bitte eine CSV-Datei auswählen.“ / „…Zieldatei…“: Pfade setzen.
- „Die CSV-Datei enthält keine verwertbaren Daten.“: Datei prüfen / Trennzeichen korrekt?
- Spaltenbereinigung schlägt fehl → Dialog zeigt den Fehlertext (z. B. leere gesamte Tabelle).

## Lizenz
Füge hier deine Lizenz ein (z. B. MIT).

-------------------------------------------------------------------------------
2) 🇬🇧 README (English)
-------------------------------------------------------------------------------

# CSV → A4 Table (PNG/JPG/PDF/XLSX)

Convert CSV **or** Cochrane Search Manager TXT into cleanly typeset A4 tables exported as
**PNG**, **JPG**, **PDF**, or **Excel (XLSX)**. Features zebra rows, customizable header
color, robust column width allocation, reliable wrapping (including optional **hyphenation
for headers**), and multi-page export.

## Highlights
- Inputs: **CSV** and **Cochrane Search Manager .txt** (“ID / Search / Hits”)
- Outputs: **PNG**, **JPG**, **PDF** (multi-page) or **XLSX**
- A4 at 300 DPI, **portrait/landscape**
- Zebra rows, header color presets or custom
- Robust wrapping incl. hard-chunk header wrap; optional header **hyphenation** (`pyphen`)
- Automatic column width distribution with minimum floor & natural width + buffer
- UTF‑8 toggle and presets for EBSCO/PubMed CSVs
- Option to **remove empty columns** (body-only) via a temporary cleaned CSV
- XLSX export with print title row, wrap, margins, A4, freeze panes

## Requirements
- Python **3.9+** (Windows, macOS, Linux)
- Deps: `Pillow`, `openpyxl`, optional `pyphen`; `tkinter` ships with CPython
- System fonts (Segoe UI / Arial / DejaVu Sans / Helvetica)

## Install
```bash
python -m venv .venv
# Windows:
.venv\Scripts\pip install -U pip
.venv\Scripts\pip install -r requirements.txt
# macOS/Linux:
source .venv/bin/activate
pip install -U pip
pip install -r requirements.txt
```

## Run
```bash
python your_script.py
```

## Usage (GUI)
1) Pick CSV or Cochrane TXT → 2) Choose output path → 3) Format (PNG/JPG/PDF/XLSX)
→ 4) Orientation → 5) Colors → 6) Separator (or presets) → 7) UTF‑8 toggle
→ 8) Header hyphenation → 9) Remove empty columns (optional) → 10) **Create**.

## Input / Output specifics
- CSV via `csv.reader` with chosen delimiter.
- Cochrane TXT: robust multi-line parsing; `Date Run` printed as a one-line note above table.
- PDF multi-page, PNG/JPG numbered when multiple pages; XLSX with wrapped cells and borders.

## License
Insert your license (e.g., MIT).

-------------------------------------------------------------------------------
3) Kurzanleitung (DE)
-------------------------------------------------------------------------------

**Schnellstart**
1. Programm starten: `python your_script.py`
2. CSV **oder** Cochrane-TXT wählen
3. Ziel + Format (PNG/JPG/PDF/XLSX) festlegen
4. Optional: Farben, Silbentrennung, UTF‑8, Trennzeichen, „Leere Spalten entfernen“
5. **Erstellen** klicken → Datei(en) werden gespeichert

**Tipps**
- Mehrseitige PNG/JPG werden als `name_01.png`, `name_02.png`, … geschrieben
- Für Umlaute immer **UTF‑8** aktivieren (sofern CSV in UTF‑8 vorliegt)
- Bei sehr schmalen Spalten die Chunk-Größe für Header (CODE: `HEADER_HARD_WRAP_CHARS`) ggf. erhöhen

-------------------------------------------------------------------------------
4) Code-Analyse & Empfehlungen
-------------------------------------------------------------------------------

**Stärken**
- Sehr robuster Textumbruch inkl. Header-Hard-Wrap & optionaler Silbentrennung (nur Header)
- Smarte Spaltenbreiten: Mindestbreite (3-Zeichen-Floor vs. längstes Header-Teilstück) +
  Pufferverteilung nach „Umbruch-Einsparpotential“
- Cochrane-TXT-Parser mit mehrzeiligen Queries & Meta-Feld `Date Run`
- Optionale Entfernung leerer Body-Spalten via temporärer CSV
- Excel-Export mit sinnvollen Druckeinstellungen (A4, Titelzeile, Umbruch, Freeze Panes)

**Verbesserungsvorschläge (kurz)**
1) **Logging**: Zusätzlich zum `messagebox` ein Logfile (z. B. `tempfile.gettempdir()/csv2a4.log`)
   mit `traceback` schreiben → bessere Fehlersuche.
2) **Internationalisierung**: UI-Strings in ein kleines Dict auslagern (DE/EN Toggle).
3) **Tests**: Parser/Normalisierung/Spaltenbreiten als modulare Funktionen testen.
4) **UX**: Fenster optional resizable, DPI-Awareness, größere Standard-Schrift auf High‑DPI.
5) **Performance**: Für extrem große CSVs Zeilenhöhenberechnung inkrementell oder mit Cache.
6) **CLI** (optional): Headless-Modus für Batch-Verarbeitung.
7) **Packaging**: PyInstaller-Hinweise siehe unten.

-------------------------------------------------------------------------------
5) requirements.txt & Packaging
-------------------------------------------------------------------------------

**requirements.txt (Vorschlag)**
Pillow>=10.0.0
openpyxl>=3.1.0
pyphen>=0.14.0     # optional; wird im Code abgefangen, falls nicht vorhanden

**PyInstaller (Beispiel)**
pyinstaller --noconfirm --onefile --windowed \
  --name "CSV_to_A4" \
  --add-data "csvConverter.ico;." \
  your_script.py

- `resource_path()` im Code unterstützt `--onefile`
- Systemfonts werden nicht gebündelt; auf Zielsystem vorhanden sein lassen

===============================================================================
Lizenzhinweise & Python-Bibliotheken
===============================================================================

Dieses Projekt wird von DIR als **Open Source** veröffentlicht.
Empfehlung: Verwende die MIT-Lizenz (siehe LICENSE.txt) oder passe sie nach Bedarf an.

Project is released as **Open Source** by YOU.
Recommendation: Use the MIT License (see LICENSE.txt) or adapt as needed.

-------------------------------------------------------------------------------
Projektlizenz / Project License
-------------------------------------------------------------------------------

- Lizenz: MIT (Standard; anpassbar)
- Copyright:
  Deutsch: (c) 2025 <DEIN NAME/ORGANISATION>
  English: (c) 2025 <YOUR NAME/ORGANIZATION>

Bitte ersetze die Platzhalter mit deinem Namen/Organisation.

-------------------------------------------------------------------------------
Drittanbieter-Bibliotheken (Python) / Third-Party Libraries (Python)
-------------------------------------------------------------------------------

(Hinweis: Lizenzangaben stammen – wenn verfügbar – aus Paket-Metadaten der lokalen
Installation. Prüfe die jeweiligen Projekte für verbindliche Lizenztexte.)

Python-Interpreter: 3.11.8 (PSF License 2.0)
Tkinter / Tcl/Tk: Bestandteil von Python; Tcl/Tk unter einer freizügigen Lizenz.

Folgende Bibliotheken werden vom Script verwendet:
- Pillow (Version: 9.1.0) — Bildverarbeitung / Rendering; Lizenz/Angabe: HPND
- openpyxl (Version: 3.0.10) — Excel-Export (XLSX); Lizenz/Angabe: MIT
- pyphen (Version: 0.17.2) — optionale Silbentrennung (Header); Lizenz/Angabe: Development Status :: 4 - Beta

-------------------------------------------------------------------------------
Nutzungshinweise / Usage Notes
-------------------------------------------------------------------------------
- Für die vollständigen, rechtlich maßgeblichen Lizenztexte der Drittbibliotheken
  siehe die jeweiligen Projektseiten/REPOSITORIES.
- Bei Distribution (z. B. PyInstaller-EXE) füge LICENSE.txt und diese Hinweise bei.
- Wenn du zusätzliche Bibliotheken nutzt, ergänze sie hier.

-------------------------------------------------------------------------------
Attribution / Danksagung
-------------------------------------------------------------------------------
- Icons/Assets (falls genutzt): bitte Quelle & Lizenz ergänzen.
- Cochrane TXT Parser: basiert auf eigener Implementierung; keine TXT-Inhalte enthalten.

