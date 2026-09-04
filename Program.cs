using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CsvTool.Core;

namespace CsvTool
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                PrintUsage();
                return;
            }

            bool infoMode = false;
            bool showColumnNumbers = false;
            bool showLineNumbers = false;
            var filePatterns = new List<string>();
            string? initialSearch = null;
            string? outputFile = null;
            string? columnSelection = null;
            string? lineSelection = null;
            int? initialTab = null;

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];
                if (arg == "-i" || arg == "--info") infoMode = true;
                else if (arg == "-n" || arg == "--numbers") { showColumnNumbers = true; showLineNumbers = true; }
                else if (arg == "--col-numbers") showColumnNumbers = true;
                else if (arg == "--line-numbers") showLineNumbers = true;
                else if (arg == "-f" || arg == "--find") { if (i + 1 < args.Length) initialSearch = args[++i]; }
                else if (arg == "-t" || arg == "--tab") { if (i + 1 < args.Length && int.TryParse(args[i + 1], out int t)) { initialTab = t; i++; } }
                else if (arg == "-o" || arg == "--output") { if (i + 1 < args.Length) outputFile = args[++i]; }
                else if (arg == "-c" || arg == "--columns") { if (i + 1 < args.Length) columnSelection = args[++i]; }
                else if (arg.StartsWith("-c") && arg.Length > 2) { columnSelection = arg.Substring(2); }
                else if (arg == "-l" || arg == "--lines" || arg == "--rows") { if (i + 1 < args.Length) lineSelection = args[++i]; }
                else if (arg.StartsWith("-l") && arg.Length > 2) { lineSelection = arg.Substring(2); }
                else if (arg == "-h" || arg == "--help" || arg == "-?") { PrintUsage(); return; }
                else filePatterns.Add(arg);
            }

            var resolvedFiles = ResolveFiles(filePatterns);
            if (resolvedFiles.Count == 0) { Console.WriteLine("Error: No files found."); return; }

            if (!string.IsNullOrEmpty(outputFile))
            {
                if (resolvedFiles.Count != 1) { Console.WriteLine("Error: When using '-o', exactly one input file must be specified."); return; }
                string filePath = resolvedFiles[0];
                if (!ValidateFile(filePath)) return;
                if (showColumnNumbers || showLineNumbers) Console.WriteLine("Note: '-n' only affects the interactive view and is ignored with '-o'.");

                try
                {
                    var viewer = new CsvViewer();
                    viewer.LoadFile(filePath);
                    if (initialTab.HasValue) viewer.SwitchSheet(initialTab.Value - 1);

                    if (!string.IsNullOrEmpty(lineSelection))
                    {
                        var selectedLines = LineFilter.Parse(lineSelection, viewer.Data.TotalRows);
                        if (selectedLines.Count > 0) viewer.ApplyLineSelection(selectedLines);
                        else Console.WriteLine("Warning: No valid lines selected.");
                    }

                    if (!string.IsNullOrEmpty(columnSelection))
                    {
                        var selectedCols = ColumnFilter.Parse(columnSelection, viewer.Data.TotalCols);
                        if (selectedCols.Count > 0) viewer.ApplyColumnSelection(selectedCols);
                        else Console.WriteLine("Warning: No valid columns selected.");
                    }

                    string extension = Path.GetExtension(outputFile).ToLowerInvariant();
                    if (extension == ".xlsx") viewer.SaveAsExcel(outputFile);
                    else if (extension == ".ods") viewer.SaveAsOds(outputFile);
                    else viewer.SaveAsCsv(outputFile);
                    Console.WriteLine($"Successfully saved to '{outputFile}'.");
                }
                catch (Exception ex) { Console.WriteLine($"Error saving file: {ex.Message}"); }
                return;
            }

            bool hasWildcards = filePatterns.Any(p => p.Contains('*') || p.Contains('?'));
            if (hasWildcards || resolvedFiles.Count > 1) infoMode = true;

            if (infoMode) PrintFileInfoTable(resolvedFiles);
            else
            {
                string filePath = resolvedFiles[0];
                if (!ValidateFile(filePath)) return;
                try
                {
                    var viewer = new CsvViewer();
                    viewer.ShowColumnNumbers = showColumnNumbers;
                    viewer.ShowLineNumbers = showLineNumbers;
                    List<int>? selectedCols = null;
                    List<int>? selectedLines = null;

                    if (!string.IsNullOrEmpty(lineSelection) || !string.IsNullOrEmpty(columnSelection))
                    {
                        viewer.LoadFile(filePath);
                        if (initialTab.HasValue) viewer.SwitchSheet(initialTab.Value - 1);

                        if (!string.IsNullOrEmpty(lineSelection))
                        {
                            selectedLines = LineFilter.Parse(lineSelection, viewer.Data.TotalRows);
                            if (selectedLines.Count == 0) Console.WriteLine("Warning: No valid lines selected.");
                        }

                        if (!string.IsNullOrEmpty(columnSelection))
                        {
                            selectedCols = ColumnFilter.Parse(columnSelection, viewer.Data.TotalCols);
                            if (selectedCols.Count == 0) Console.WriteLine("Warning: No valid columns selected.");
                        }
                    }
                    viewer.Run(filePath, initialSearch, initialTab, selectedCols, selectedLines);
                }
                catch (Exception ex)
                {
                    Console.Clear();
                    Console.WriteLine("An error occurred:");
                    Console.WriteLine(ex.Message);
                }
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine("NAME");
            Console.WriteLine($"    csv {GetVersion()} - A lightweight CLI viewer for CSV Libre Office Calc / Excel files.");
            Console.WriteLine();
            Console.WriteLine("SYNOPSIS");
            Console.WriteLine("    csv [OPTIONS] [FILE | PATTERN]");
            Console.WriteLine();
            Console.WriteLine("DESCRIPTION");
            Console.WriteLine("    Opens and displays CSV, text, and Excel files (.xls, .xlsx) as well as LibreOffice ODS files in a scrollable");
            Console.WriteLine("    terminal interface. Supports searching and launching external editors.");
            Console.WriteLine();
            Console.WriteLine("OPTIONS");
            Console.WriteLine("    -f, --find <TERM>");
            Console.WriteLine("        Immediately search for TERM upon opening the file.");
            Console.WriteLine();
            Console.WriteLine("    -t, --tab <INDEX>");
            Console.WriteLine("        Open the specific Excel/ODS sheet index (1-based).");
            Console.WriteLine();
            Console.WriteLine("    -c, --columns <SPEC>");
            Console.WriteLine("        Select columns to display or export. SPEC is a comma-separated list");
            Console.WriteLine("        of 1-based column numbers or ranges (e.g. -c 2-5,8).");
            Console.WriteLine();
            Console.WriteLine("    -l, --lines, --rows <SPEC>");
            Console.WriteLine("        Select lines/rows to display or export. SPEC is a comma-separated list");
            Console.WriteLine("        of 1-based line numbers or ranges (e.g. -l 10-20, -l 50-, -l 2,5).");
            Console.WriteLine("        Line 1 is the header row, which is always preserved.");
            Console.WriteLine();
            Console.WriteLine("    -n, --numbers");
            Console.WriteLine("        Show line numbers in a fixed left gutter and column numbers in");
            Console.WriteLine("        parentheses in the header (e.g. 'Customer (2)'), so numbers for");
            Console.WriteLine("        '-l' and '-c' can be easily read off. Combined with '-l' and/or '-c',");
            Console.WriteLine("        the original numbers of the source file are shown. Interactive view only.");
            Console.WriteLine();
            Console.WriteLine("    --line-numbers");
            Console.WriteLine("        Show only line numbers in the fixed left gutter.");
            Console.WriteLine();
            Console.WriteLine("    --col-numbers");
            Console.WriteLine("        Show only column numbers in the header.");
            Console.WriteLine();
            Console.WriteLine("    -o, --output <FILE>");
            Console.WriteLine("        Convert the input file (or selected sheet) to a UTF-8 encoded");
            Console.WriteLine("        CSV file, XLSX Excel workbook, or ODS document and save it to the");
            Console.WriteLine("        specified path. The format is determined by the file extension");
            Console.WriteLine("        (.csv, .xlsx, or .ods).");
            Console.WriteLine();
            Console.WriteLine("    -i, --info");
            Console.WriteLine("        Display file metadata (Size, Date, Dimensions, Encoding) in a table");
            Console.WriteLine("        format instead of opening the interactive viewer.");
            Console.WriteLine("        Automatically enabled if a wildcard pattern is provided or multiple");
            Console.WriteLine("        files match.");
            Console.WriteLine();
            Console.WriteLine("    -?, -h, --help");
            Console.WriteLine("        Show this help message.");
            Console.WriteLine();
            Console.WriteLine("CONTROLS");
            Console.WriteLine("    Arrows, PgUp/Dn    Navigation");
            Console.WriteLine("    1-9                Switch Excel/ODS Sheet (if available)");
            Console.WriteLine("    F                  Find/Search");
            Console.WriteLine("    F3 / Shift+F3      Find Next / Previous");
            Console.WriteLine("    L                  Open in LibreOffice");
            Console.WriteLine("    E                  Open in Excel");
            Console.WriteLine("    Q / ESC            Quit");
            Console.WriteLine();
            Console.WriteLine("AUTHORS");
            Console.WriteLine("    Gregor Stübner, Claude (Anthropic), Gemini, Deepseek, Kimi");
        }

        static string GetVersion()
        {
            var v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            return v == null ? "" : $"{v.Major}.{v.Minor}.{v.Build}";
        }

        static List<string> ResolveFiles(List<string> patterns)
        {
            var results = new List<string>();
            foreach (var pattern in patterns)
            {
                if (pattern.Contains('*') || pattern.Contains('?'))
                {
                    string? dir = Path.GetDirectoryName(pattern);
                    if (string.IsNullOrEmpty(dir)) dir = Directory.GetCurrentDirectory();
                    string filePattern = Path.GetFileName(pattern);
                    if (Directory.Exists(dir))
                    {
                        var matches = Directory.GetFiles(dir, filePattern).Where(f => {
                            string ext = Path.GetExtension(f).ToLowerInvariant();
                            return ext == ".csv" || ext == ".txt" || ext == ".xls" || ext == ".xlsx" || ext == ".ods";
                        });
                        results.AddRange(matches);
                    }
                }
                else results.Add(pattern);
            }
            return results.Distinct().OrderBy(f => f).ToList();
        }

        static bool ValidateFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".csv" && ext != ".txt" && ext != ".xls" && ext != ".xlsx" && ext != ".ods") { Console.WriteLine($"Error: {ext} not supported."); return false; }
            if (!File.Exists(filePath)) { Console.WriteLine($"Error: {filePath} not found."); return false; }
            return true;
        }

        static void PrintFileInfoTable(List<string> files)
        {
            string fmt = "{0,-30} | {1,10} | {2,-19} | {3,-12} | {4,-15} | {5,-9}";
            Console.WriteLine(fmt, "Filename", "Size", "Date", "Dimension", "Encoding", "Separator");
            Console.WriteLine(new string('-', 110));
            foreach (var file in files)
            {
                if (!File.Exists(file)) continue;
                try
                {
                    var viewer = new CsvViewer();
                    viewer.LoadFile(file);
                    var data = viewer.Data;
                    string name = Path.GetFileName(file);
                    if (name.Length > 30) name = name.Substring(0, 27) + "...";
                    string sizeStr = FormatBytes(data.FileSize);
                    string dateStr = data.LastWriteTime.ToString("g");
                    string dimStr = $"{data.TotalRows}x{data.TotalCols}";
                    string encName = data.Encoding?.EncodingName ?? "N/A";
                    string sepStr = data.Delimiter != '\0' ? $"'{data.Delimiter}'" : "N/A";
                    Console.WriteLine(fmt, name, sizeStr, dateStr, dimStr, encName, sepStr);
                }
                catch (Exception ex) { Console.WriteLine($"Error reading {Path.GetFileName(file)}: {ex.Message}"); }
            }
        }

        static string FormatBytes(long bytes)
        {
            string[] suffix = { "B", "KB", "MB", "GB" };
            int i; double dblSByte = bytes;
            for (i = 0; i < suffix.Length && bytes >= 1024; i++, bytes /= 1024) dblSByte = bytes / 1024.0;
            return $"{dblSByte:0.##} {suffix[i]}";
        }
    }
}
