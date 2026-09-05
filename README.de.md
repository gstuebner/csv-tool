# CsvTool (CLI)

[English](README.md) · **Deutsch**

Ein performanter, leichtgewichtiger Kommandozeilen-Viewer (CLI) für CSV-, Excel- (`.xls`, `.xlsx`) und LibreOffice-Dateien (`.ods`). Entwickelt mit C# und .NET 8.

![License](https://img.shields.io/badge/license-MIT-blue.svg)

## Funktionen

*   **TUI (Textbasierte Benutzeroberfläche):** Scrollbare Tabellenansicht direkt im Terminal.
*   **Formatunterstützung:**
    *   CSV / Text (Automatische Erkennung des Trennzeichens: `,`, `;`, `\t`).
    *   Excel (`.xlsx`, `.xls`).
    *   OpenDocument-Tabellendokumente (`.ods`).
*   **RFC 4180 konform:** In Anführungszeichen gesetzte Felder mit Zeilenumbrüchen (wie von Excel oder LibreOffice erzeugt) werden korrekt gelesen und geschrieben — ein Datensatz bleibt eine Zeile.
*   **Intelligente Zeichenkodierung:** Erkennt automatisch UTF-8 vs. Windows-1252 (ANSI), um Sonderzeichen und deutsche Umlaute korrekt darzustellen.
*   **Suche:** Integrierte Suchfunktion (Taste `F` oder Argument `-f`).
*   **Spaltenauswahl:** Bestimmte Spalten mit `-c` auswählen (z. B. `-c 2-5,8` oder `-c 5-`).
*   **Zeilenauswahl:** Bestimmte Zeilen mit `-l` filtern (z. B. `-l 10-20`, `-l 50-`, `-l 2,5,8`). Zeile 1 ist die Kopfzeile und bleibt immer erhalten.
*   **Zeilen- und Spaltennummern:** `-n` zeigt Zeilennummern links an und hängt die Spaltennummer an jeden Spaltenkopf an (z. B. `Kunde (2)`), um Nummern für `-l` und `-c` leicht abzulesen.
*   **Export/Konvertierung:** Tabellenblätter nach CSV, Excel oder ODS konvertieren (`-o` Argument). CSV-Ausgabe erfolgt in UTF-8 mit BOM und CRLF-Zeilenumbrüchen und behält das Quelltrennzeichen bei.
*   **Externe Integration:** Die aktuelle Datei schnell in Excel (`E`) oder LibreOffice (`L`) öffnen.
*   **Plattformübergreifend:** Läuft unter Windows, Linux und macOS. Externe Programmaufrufe (`E`, `L`) unterstützen `xdg-open` (Linux) und `open` (macOS).

## Installation

### Voraussetzungen
*   .NET 8.0 SDK (zum Bauen)

### Bauen
Repository klonen und ausführen:

```bash
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```
*(Für andere Plattformen `-r` zu `linux-x64` oder `osx-x64` ändern)*

Die fertige Datei liegt unter `bin/Release/net8.0/win-x64/publish/`.

## Benutzung

```bash
# Datei interaktiv öffnen
csv meinedatei.csv

# Excel-Datei an einem bestimmten Tabellenblatt öffnen (1-basierter Index)
csv daten.xlsx -t 2

# Direkt mit Suchbegriff öffnen
csv daten.csv -f "Suchbegriff"

# Datei konvertieren/exportieren (z. B. Excel nach CSV, ODS oder XLSX)
csv daten.xlsx -o ausgabe.csv

# Blatt 2 einer ODS-Datei nach Excel konvertieren
csv daten.ods -t 2 -o ausgabe.xlsx

# Zeilen- und Spaltennummern anzeigen (für -l und -c)
csv daten.csv -n

# Zeilen 10 bis 20 filtern (Kopfzeile 1 bleibt immer erhalten)
csv daten.csv -l 10-20

# Zeilen ab Zeile 50 bis zum Ende filtern
csv daten.csv -l 50-

# Auswahl prüfen: Kopfzeilen behalten Spaltennummern, Rand behält Original-Zeilennummern
csv daten.csv -c 2,5 -l 10-20 -n

# Spalten 2-5 und Zeilen 1-100 auswählen und als ODS exportieren
csv daten.csv -c 2-5 -l 1-100 -o ausgabe.ods

# Nur Datei-Informationen anzeigen (ohne interaktiven Modus)
csv *.csv
```

## Tastatursteuerung

| Taste | Aktion |
| :--- | :--- |
| `Pfeiltasten`, `BildAuf/Ab` | Navigation |
| `Pos1`, `Ende` | Zum Anfang / Ende springen |
| `1`-`9` | Tabellenblatt wechseln (Excel/ODS) |
| `F` | Suchen |
| `F3` / `Umschalt+F3` | Weitersuchen vorwärts / rückwärts |
| `L` | In LibreOffice öffnen |
| `E` | In Excel öffnen |
| `Q` / `ESC` | Beenden |

## Lizenz

Dieses Projekt ist unter der MIT-Lizenz lizenziert — siehe [LICENSE](LICENSE).

## Drittanbieter-Bibliotheken

Dieses Projekt nutzt folgende Open-Source-Bibliotheken:

*   **ExcelDataReader** (MIT-Lizenz) - Lesen von Excel-Binär- und XML-Formaten.
*   **ClosedXML** (MIT-Lizenz) - Erstellen/Exportieren von Excel-Dateien.
*   **System.Text.Encoding.CodePages** (MIT-Lizenz) - Unterstützung für historische Zeichensätze.

## Autoren

*   Gregor Stübner
*   Claude (Anthropic), Gemini, Deepseek & Kimi (KI-Assistenten)
