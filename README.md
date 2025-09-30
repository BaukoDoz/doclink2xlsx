# Inhalt Repository:

Das Repository enthält zwei Skripte, die Links aus Word- bzw. PDF-Dateien auslesen, die verlinkten Webseiten abrufen und die Ergebnisse in einer Excel-Datei sammeln.

## Funktionsumfang
- Extrahiert alle externen Links aus DOCX- oder PDF-Dateien
- Sendet HTTP-Anfragen an die gefundenen URLs und liest den Seitentitel aus
- Speichert Titel, URL und Ursprungsseite in einer Excel-Arbeitsmappe (XLSX)

## Voraussetzungen installieren
Alle Schritte funktionieren im Windows Terminal bzw. in PowerShell. Sie benoetigen eine Internetverbindung.

1. **Python installieren**:

    winget install Python.Python.3.12

   Nach der Installation das Terminal schliessen und erneut öffnen, dann pruefen:

    python --version

2. **Git installieren**:

    winget install Git.Git

   Anschliessend kontrollieren:

    git --version

3. **(Optional) Ausfuehrungsrichtlinie für Skripte anpassen**, falls Aktivierung der virtuellen Umgebung fehlschlägt:

    Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

## Projekt herunterladen
Wähle einen Ordner aus, in dem die Projektdateien liegen sollen, und klone das Repository.

Wenn das Projekt als ZIP vorliegt, entpacke den Ordner und öffnen ihn im Terminal.

## Virtuelle Umgebung einrichten
1. **Virtuelle Umgebung erstellen** (legt einen isolierten Python-Ordner im Projekt an):

    python -m venv .venv

2. **Aktivieren**:

    .\\.venv\Scripts\Activate.ps1

   Das Terminal zeigt nun links ein (.venv) an. Bei Bedarf fuer die klassische Eingabeaufforderung (CMD):

    .\\.venv\Scripts\activate.bat

   Unter macOS/Linux lautet der Befehl:

    source .venv/bin/activate

3. **Abhaengigkeiten installieren**:

    python -m pip install --upgrade pip
    pip install -r requirements.txt

## Anwendung
Die Skripte laufen innerhalb der aktivierten virtuellen Umgebung.

### Links aus einer Word-Datei (DOCX) ziehen

    python word_links_to_excel.py PFAD\zur\Datei.docx

Ohne weiteres Argument wird eine Excel-Datei mit gleichem Namen im selben Ordner erstellt (Dateiendung .xlsx). Optional kann ein Ausgabepfad vergeben werden:

    python word_links_to_excel.py PFAD\zur\Datei.docx Ausgabe.xlsx

### Links aus einer PDF-Datei ziehen

    python pdf_links_to_excel.py PFAD\zur\Datei.pdf

Auch hier entsteht standardmaessig eine gleichnamige .xlsx-Datei. Mit zweitem Argument bestimmen Sie den Ausgabepfad.

### Gemeinsame Optionen
Beide Skripte akzeptieren --timeout, um das Warten auf Webseiten anzupassen (Standard 10 Sekunden):

    python word_links_to_excel.py Datei.docx --timeout 5

## Ergebnisdatei
Die erzeugte Excel-Datei enthaelt ein Arbeitsblatt Links mit drei Spalten:
- Seitentitel: Titelleiste der aufgerufenen Webseite (leer, falls kein Titel gefunden wurde)
- Link: Urspruengliche URL aus dem Dokument
- Seite: Seitenzahl im Ursprungsdokument

## Fehlersuche
- Falls keine Titel geladen werden, prüfe Internetverbindung oder höherer Timeout-Wert.
- Bei sehr grossen Dokumenten kann das Skript lange laufen, weil jede URL einzeln aufgerufen wird.
- Wenn PowerShell das Aktivieren der virtuellen Umgebung blockiert, siehe das oben genannte Set-ExecutionPolicy.

- Sollten man die Umgebung verlassen wollen, nutze deactivate.
