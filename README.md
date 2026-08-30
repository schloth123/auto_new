"# OneDrive File Downloader

Ein Python-Skript zum Herunterladen von Dateien von OneDrive-Links, die in einer Excel-Datei gespeichert sind, mit automatischer Aktualisierung der lokalen Pfade in der Tabelle.

## Features

✅ **OneDrive-Links verarbeiten** - Konvertiert OneDrive-Freigabelinks zu direkten Download-URLs  
✅ **Batch-Download** - Lädt mehrere Dateien automatisch herunter  
✅ **Excel-Integration** - Liest Links aus Excel und schreibt lokale Pfade zurück  
✅ **Fehlerbehandlung** - Robustes Error-Handling mit Fortführung bei Fehlern  
✅ **Umfassende Reports** - Detailliertes Logging und Fehlerreporting  
✅ **Flexible Spalten** - Unterstützt verschiedene Spaltenbezeichnungen  

## Installation

### Voraussetzungen
- Python 3.8+
- pip

### Abhängigkeiten installieren

```bash
pip install -r requirements.txt
```

oder mit pip install direkt:

```bash
pip install requests openpyxl
```

## Verwendung

### Grundlegendes Format

```bash
python download_onedrive_files.py -e EXCEL_DATEI -s QUELLSPALTE -t ZIELSPALTE -p ZIELVERZEICHNIS
```

### Parameter

| Parameter | Kurzform | Beschreibung | Beispiel |
|-----------|----------|-------------|---------|
| `--excel` | `-e` | Pfad zur Excel-Datei (erforderlich) | `data.xlsx` |
| `--source-col` | `-s` | Spalte mit OneDrive-Links (Buchstabe oder Index, erforderlich) | `A` oder `1` |
| `--target-col` | `-t` | Spalte für lokale Pfade (Buchstabe oder Index, erforderlich) | `B` oder `2` |
| `--target-path` | `-p` | Lokales Zielverzeichnis (erforderlich) | `./downloads` |
| `--log-file` | `-l` | Optionaler Custom Log-Datei-Pfad | `./logs/report.log` |

### Beispiele

**Basis-Beispiel:**
```bash
python download_onedrive_files.py -e data.xlsx -s A -t B -p ./downloads
```

**Mit benutzerdefinierten Spaltenindizes:**
```bash
python download_onedrive_files.py -e data.xlsx -s 1 -t 2 -p ./downloads
```

**Mit Custom Log-Datei:**
```bash
python download_onedrive_files.py -e data.xlsx -s A -t B -p ./downloads -l ./reports/download.log
```

## Excel-Format

Die Excel-Datei sollte folgendes Format haben:

| OneDrive-Link | Lokaler Pfad |
|---------------|-------------|
| https://1drv.ms/b/c/e7cee1d47fca9dd4/IQDiMORSPfCCQ6IzNV_HwfGhAc4-rb87dHMKwk-9-_G2jI0?e=bXz2c9 | |
| https://1drv.ms/x/c/abc123.../... | |

Nach der Ausführung werden die lokalen Pfade automatisch in der Zielspalte eingetragen:

| OneDrive-Link | Lokaler Pfad |
|---------------|-------------|
| https://1drv.ms/b/c/e7cee1d47fca9dd4/... | /path/to/downloads/document.pdf |
| https://1drv.ms/x/c/abc123.../... | /path/to/downloads/spreadsheet.xlsx |

## Reports und Logging

Das Skript generiert automatisch mehrere Reports:

1. **Log-Datei** (`download_report_TIMESTAMP.log`)
   - Detailliertes Protokoll aller Aktionen
   - Fehlerberichte für fehlgeschlagene Downloads
   - Speicherort: Zielverzeichnis

2. **JSON-Report** (`report_TIMESTAMP.json`)
   - Strukturierter Report mit Statistiken
   - Fehler-Details
   - Speicherort: Zielverzeichnis

3. **Console-Output**
   - Echtzeitinformationen während der Ausführung
   - Zusammenfassung am Ende

### Report-Inhalt

- Gesamtzahl verarbeiteter Dateien
- Erfolgreiche Downloads
- Fehlgeschlagene Downloads mit Fehlermeldungen
- Übersprungene Zeilen (leere Zellen)
- Erfolgsquote in Prozent

## Fehlerbehandlung

Das Skript implementiert robustes Error-Handling:

- **Timeout-Fehler**: Automatischer Timeout nach 30 Sekunden
- **Verbindungsfehler**: Detaillierte Fehlermeldungen
- **HTTP-Fehler**: Status-Codes und Ursachen werden geloggt
- **Unerwartete Fehler**: Werden abgefangen und geloggt, Verarbeitung läuft weiter
- **Fehlende Dateien**: Excel-Datei wird auf Existenz geprüft

Bei Fehlern:
- Die Verarbeitung wird **nicht unterbrochen**
- Fehler werden in den Reports dokumentiert
- Teilweise erfolgreiche Verarbeitungen werden gespeichert
- Exit-Code zeigt das Ergebnis an

## Exit-Codes

| Code | Bedeutung |
|------|-----------|
| 0 | Erfolgreich abgeschlossen, alle Dateien heruntergeladen |
| 1 | Abgeschlossen mit Fehlern (Einige Dateien konnten nicht heruntergeladen werden) |
| 2 | Kritischer Fehler (Prozess abgebrochen) |

## Erweiterte Verwendung

### Skript-Integration in Python

```python
from download_onedrive_files import OneDriveDownloader

# Downloader erstellen
downloader = OneDriveDownloader(
    excel_path='data.xlsx',
    excel_column_source='A',
    excel_column_target='B',
    target_path='./downloads'
)

# Dateien verarbeiten
success = downloader.process_excel()

# Report generieren
downloader.generate_report()
```

## Bekannte Limitierungen

- Große Dateien (>4GB) benötigen längere Timeouts
- OneDrive-Freigabelinks werden automatisch in Download-URLs konvertiert
- Dateinamen werden aus der URL extrahiert
- Windows- und Unix-Pfade werden automatisch verarbeitet

## Troubleshooting

### "Excel file not found"
- Überprüfen Sie den Pfad zur Excel-Datei
- Nutzen Sie absolute Pfade oder relative Pfade vom aktuellen Verzeichnis

### "HTTP error 404"
- Der OneDrive-Link ist möglicherweise abgelaufen oder ungültig
- Überprüfen Sie, ob die Datei noch freigegeben ist

### "Timeout downloading file"
- Erhöhen Sie den Timeout-Wert im Code (derzeit 30 Sekunden)
- Überprüfen Sie Ihre Internetverbindung

### "Invalid column identifier"
- Verwenden Sie gültige Spaltenbuchstaben (A, B, C...) oder Indizes (1, 2, 3...)
- Spalten sind 1-indiziert (erste Spalte = A oder 1)

## Development

### Abhängigkeiten für Entwicklung installieren

```bash
pip install -e ".[dev]"
```

### Code formatieren

```bash
black download_onedrive_files.py
```

### Linting

```bash
flake8 download_onedrive_files.py
```

### Type-Checking

```bash
mypy download_onedrive_files.py
```

## Lizenz

MIT

## Autor

Markus

## Support

Bei Fragen oder Problemen erstellen Sie bitte ein Issue im Repository." 
