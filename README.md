# CsvTool (CLI)

A lightweight, high-performance Command Line Interface (CLI) viewer for CSV, Excel (`.xls`, `.xlsx`), and LibreOffice (`.ods`) files. Built with C# and .NET 8.

![License](https://img.shields.io/badge/license-MIT-blue.svg)

## Features

*   **TUI (Text-based User Interface):** Scrollable grid view directly in your terminal.
*   **Format Support:**
    *   CSV / Text (Auto-detects delimiter: `,`, `;`, `\t`).
    *   Excel (`.xlsx`, `.xls`).
    *   OpenDocument Spreadsheet (`.ods`).
*   **RFC 4180 compliant:** Quoted fields containing line breaks (as produced by Excel or LibreOffice) are read and written correctly — one record stays one row.
*   **Smart Encoding:** Automatically detects UTF-8 vs Windows-1252 (ANSI) to handle special characters (like German Umlaute) correctly.
*   **Search:** built-in search functionality (`F` or `-f` argument).
*   **Column Selection:** Select specific columns with `-c` (e.g. `-c 2-5,8`).
*   **Column Numbers:** `-n` appends the column number to each header name (e.g. `Customer (2)`), so the numbers for `-c` can be read off instead of counted by hand. Together with `-c` the original numbers of the source file are shown.
*   **Export/Convert:** Convert sheets to CSV, Excel, or ODS (`-o` argument). CSV output is UTF-8 with BOM, uses CRLF line endings and keeps the delimiter of the source file, so Excel opens it correctly.
*   **External Integration:** Quickly open the current file in Excel (`E`) or LibreOffice (`L`).
*   **Cross-Platform Logic:** Runs on Windows, Linux, and macOS. External tool shortcuts (`E`, `L`) now support `xdg-open` (Linux) and `open` (macOS).

## Installation

### Requirements
*   .NET 8.0 SDK (to build)

### Build
Clone the repository and run:

```bash
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```
*(Change `-r` to `linux-x64` or `osx-x64` for other platforms)*

The resulting binary will be in `bin/Release/net8.0/win-x64/publish/`.

## Usage

```bash
# Open a file interactively
csv myfile.csv

# Open an Excel file at a specific sheet (1-based index)
csv data.xlsx -t 2

# Search immediately upon opening
csv data.csv -f "search term"

# Convert/Export file (e.g., Excel to CSV, ODS, or XLSX)
csv data.xlsx -o output.csv

# Convert Sheet 2 of an ODS file to Excel
csv data.ods -t 2 -o output.xlsx

# Show column numbers in the header to find the numbers for -c
csv data.csv -n

# Verify a selection: the headers keep the original numbers (2) and (5)
csv data.csv -c 2,5 -n

# Select only columns 2-5 and 8, then export to ODS
csv data.csv -c 2-5,8 -o output.ods

# Show file info only (no interactive mode)
csv *.csv
```

## Controls

| Key | Action |
| :--- | :--- |
| `Arrows`, `PgUp/Dn` | Navigation |
| `Home`, `End` | Jump to Start / End |
| `1`-`9` | Switch Sheets (Excel/ODS) |
| `F` | Search |
| `F3` / `Shift+F3` | Find Next / Previous |
| `L` | Open in LibreOffice |
| `E` | Open in Excel |
| `Q` / `ESC` | Quit |

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Third-Party Dependencies

This project uses the following open-source libraries:

*   **ExcelDataReader** (MIT License) - Reading Excel binary and XML formats.
*   **ClosedXML** (MIT License) - Creating/Exporting Excel files.
*   **System.Text.Encoding.CodePages** (MIT License) - Support for legacy encodings.

## Authors

*   Gregor Stübner
*   Claude (Anthropic), Gemini, Deepseek & Kimi (AI Assistants)

