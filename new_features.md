# Neue Funktionen & Refactoring (Work in Progress)

Dieses Dokument beschreibt die umfassenden Änderungen, die an der **CsvTool** Architektur vorgenommen wurden. 

> [!WARNING]
> **Status:** Diese Version befindet sich aktuell auf dem Branch `feature/refactoring-and-improvements` und wurde noch **nicht** in den `main` Branch gepusht oder auf ein Remote-Repository übertragen.
> 
> Es funktioniert noch nicht alles perfekt – weitere Tests und Fehlerbehebungen sind für die nächste Session geplant.

## 1. Architektur-Refactoring (Modularisierung)
Die bisherige monolithische Struktur der `Program.cs` wurde aufgelöst und in logische Komponenten unterteilt, um die Wartbarkeit zu verbessern:

*   **`Core/`**: Logik für Datenverarbeitung und Parsing.
    *   `FileData.cs`: Zentrales Datenmodell.
    *   `EncodingDetector.cs`: Automatische Erkennung von UTF-8 und Windows-1252.
    *   `CsvParser.cs`: Trennzeichen-Erkennung und CSV-Parsing.
    *   `ExcelHandler.cs`: Lesen/Schreiben von Excel-Dateien (XLS, XLSX).
    *   `OdsParser.cs`: Parser für OpenDocument Spreadsheets.
*   **`Tui/`**: Logik für die Benutzeroberfläche.
    *   `Renderer.cs`: Zuständig für das Zeichnen von Header, Grid und Footer.
*   **`Platform/`**: Betriebssystemspezifische Funktionen.
    *   `ExternalLauncher.cs`: Zentrale Stelle für den Aufruf externer Programme.

## 2. Verbesserter Cross-Platform Support
Die Shortcuts zum Öffnen von Dateien in externen Programmen wurden für Linux und macOS optimiert:
*   **Linux**: Nutzt nun `xdg-open` als Fallback, falls `scalc` (LibreOffice) oder `excel` nicht direkt im Pfad gefunden werden.
*   **macOS**: Nutzt den `open` Befehl.
*   **Windows**: Behält die bewährte Suche in Standardpfaden bei.

## 3. Build-System
*   Die Skripte `compile-for-linux.sh` und `compile-for-windows.sh` wurden korrigiert und vereinheitlicht.
*   Unterstützung für Single-File Binaries mit Trimming für minimale Dateigröße.

## 4. Bekannte Themen / Nächste Schritte
*   Überprüfung der TUI-Interaktion nach dem Refactoring.
*   Fehlerbehandlung bei beschädigten oder passwortgeschützten Dateien verfeinern.
*   Merge in den `main` Branch nach Abschluss der Tests.
