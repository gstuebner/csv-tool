using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using ExcelDataReader;
using ClosedXML.Excel;

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
            var filePatterns = new List<string>();
            string initialSearch = null;
            string outputFile = null;
            int? initialTab = null;

            // 1. Parse Arguments
            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];
                if (arg == "-i" || arg == "--info")
                {
                    infoMode = true;
                }
                else if (arg == "-f" || arg == "--find")
                {
                    if (i + 1 < args.Length)
                    {
                        initialSearch = args[++i];
                    }
                    else
                    {
                        Console.WriteLine("Error: Argument '-f' / '--find' requires a search term.");
                        return;
                    }
                }
                else if (arg == "-t" || arg == "--tab")
                {
                    if (i + 1 < args.Length && int.TryParse(args[i + 1], out int t))
                    {
                        initialTab = t;
                        i++;
                    }
                    else
                    {
                        Console.WriteLine("Error: Argument '-t' / '--tab' requires a valid integer sheet number.");
                        return;
                    }
                }
                else if (arg == "-o" || arg == "--output")
                {
                    if (i + 1 < args.Length)
                    {
                        outputFile = args[++i];
                    }
                    else
                    {
                        Console.WriteLine("Error: Argument '-o' / '--output' requires a file path.");
                        return;
                    }
                }
                else if (arg == "-h" || arg == "--help" || arg == "-?")
                {
                    PrintUsage();
                    return;
                }
                else
                {
                    filePatterns.Add(arg);
                }
            }

            // 2. Expand Wildcards & Resolve Files
            var resolvedFiles = ResolveFiles(filePatterns);

            if (resolvedFiles.Count == 0)
            {
                Console.WriteLine("Error: No files found.");
                return;
            }

            // Special Mode: Output/Convert
            if (!string.IsNullOrEmpty(outputFile))
            {
                if (resolvedFiles.Count != 1)
                {
                    Console.WriteLine("Error: When using '-o', exactly one input file must be specified.");
                    return;
                }

                string filePath = resolvedFiles[0];
                if (!ValidateFile(filePath)) return;

                try
                {
                    var viewer = new CsvViewer();
                    viewer.LoadFile(filePath);

                    if (initialTab.HasValue)
                    {
                        // Switch sheet if applicable (1-based index from args, internal is 0-based)
                        viewer.SwitchSheet(initialTab.Value - 1);
                    }

                    string extension = Path.GetExtension(outputFile).ToLowerInvariant();
                    if (extension == ".xlsx")
                    {
                        viewer.SaveAsExcel(outputFile);
                    }
                    else
                    {
                        viewer.SaveAsCsv(outputFile);
                    }
                    Console.WriteLine($"Successfully saved to '{outputFile}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving file: {ex.Message}");
                }
                return;
            }

            // 3. Determine Mode (Implicit Info Mode if wildcards used or multiple files)
            bool hasWildcards = filePatterns.Any(p => p.Contains('*') || p.Contains('?'));
            if (hasWildcards || resolvedFiles.Count > 1)
            {
                infoMode = true;
            }

            // 4. Execute
            if (infoMode)
            {
                PrintFileInfoTable(resolvedFiles);
            }
            else
            {
                // Single file view mode
                string filePath = resolvedFiles[0];
                if (!ValidateFile(filePath)) return;
                
                try
                {
                    var viewer = new CsvViewer();
                    viewer.Run(filePath, initialSearch, initialTab);
                }
                catch (Exception ex)
                {
                    Console.Clear();
                    Console.WriteLine("An error occurred:");
                    Console.WriteLine(ex.Message);
                    
                    // Only show stack trace for unexpected errors
                    if (!ex.Message.Contains("encrypted") && !ex.Message.Contains("supported"))
                    {
                        Console.WriteLine(ex.StackTrace);
                    }
                }
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine("NAME");
            Console.WriteLine("    csv - A lightweight CLI viewer for CSV and Excel / Libre Office Calc files.");
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
            Console.WriteLine("    -o, --output <FILE>");
            Console.WriteLine("        Convert the input file (or selected sheet) to a UTF-8 encoded");
            Console.WriteLine("        CSV file (or XLSX Excel workbook) and save it to the specified path.");
            Console.WriteLine("        The format is determined by the file extension (.csv or .xlsx).");
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
            Console.WriteLine("    Gregor Stübner, Gemini & Deepseek");
        }

        static List<string> ResolveFiles(List<string> patterns)
        {
            var results = new List<string>();
            foreach (var pattern in patterns)
            {
                // Check if it contains wildcards
                if (pattern.Contains('*') || pattern.Contains('?'))
                {
                    string dir = Path.GetDirectoryName(pattern);
                    if (string.IsNullOrEmpty(dir)) dir = Directory.GetCurrentDirectory();
                    
                    string filePattern = Path.GetFileName(pattern);
                    
                    if (Directory.Exists(dir))
                    {
                        try
                        {
                            var matches = Directory.GetFiles(dir, filePattern)
                                .Where(f => {
                                    string ext = Path.GetExtension(f).ToLowerInvariant();
                                    return ext == ".csv" || ext == ".txt" || ext == ".xls" || ext == ".xlsx" || ext == ".ods";
                                });
                            results.AddRange(matches);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Could not process pattern '{pattern}': {ex.Message}");
                        }
                    }
                }
                else
                {
                    // Literal path
                    results.Add(pattern);
                }
            }
            // Remove duplicates and sort
            return results.Distinct().OrderBy(f => f).ToList();
        }

        static bool ValidateFile(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".csv" && extension != ".txt" && extension != ".xls" && extension != ".xlsx" && extension != ".ods")
            {
                Console.WriteLine($"Error: File type '{extension}' is not supported.");
                return false;
            }

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File '{filePath}' not found.");
                return false;
            }
            return true;
        }

        static void PrintFileInfoTable(List<string> files)
        {
            // Table Header
            // Filename | Size (KB) | Date | Dimension | Encoding | Seperator
            
            string fmt = "{0,-30} | {1,10} | {2,-19} | {3,-12} | {4,-15} | {5,-9}";
            Console.WriteLine(fmt, "Filename", "Size (KB)", "Date", "Dimension", "Encoding", "Separator");
            Console.WriteLine(new string('-', 110));

            foreach (var file in files)
            {
                if (!File.Exists(file)) continue;

                try
                {
                    var viewer = new CsvViewer();
                    // Load metadata without running the UI loop
                    viewer.LoadFile(file);

                    string name = Path.GetFileName(file);
                    if (name.Length > 30) name = name.Substring(0, 27) + "...";

                    // Size in KB
                    double kb = viewer.FileSize / 1024.0;
                    string sizeStr = $"{kb:N2} KB";

                    string dateStr = viewer.LastWriteTime.ToString("g");
                    string dimStr = $"{viewer.TotalRows}x{viewer.TotalCols}";
                    
                    string encName = viewer.CurrentEncoding?.EncodingName ?? "N/A";
                    if (encName.Length > 15 && viewer.CurrentEncoding != null) encName = viewer.CurrentEncoding.HeaderName;

                    string sepStr = viewer.Delimiter != '\0' ? $"'{viewer.Delimiter}'" : "N/A";

                    Console.WriteLine(fmt, name, sizeStr, dateStr, dimStr, encName, sepStr);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading {Path.GetFileName(file)}: {ex.Message}");
                }
            }
        }
    }

    class CsvViewer
    {
        private List<string[]> _data = new List<string[]>();
        private int[] _colWidths;
        private int _scrollRow = 0;
        private int _scrollCol = 0; // Index of the first visible column
        
        // Excel State
        private DataSet _excelDataSet;
        private int _currentSheetIndex = 0;

        // Metadata
        private string _fileName;
        private string _fullFilePath;
        private long _fileSize;
        private DateTime _lastWriteTime;
        private Encoding _encoding;
        private char _delimiter;
        private int _totalRows;
        private int _totalCols;
        
        // Public Accessors for Metadata
        public string FileName => _fileName;
        public long FileSize => _fileSize;
        public DateTime LastWriteTime => _lastWriteTime;
        public Encoding CurrentEncoding => _encoding;
        public char Delimiter => _delimiter;
        public int TotalRows => _totalRows;
        public int TotalCols => _totalCols;

        private string _statusMessage = "";
        private string _lastSearchTerm = "";
        private int _highlightRow = -1;

        public void Run(string path, string initialSearch = null, int? initialTab = null)
        {
            Console.Clear();
            LoadFile(path);
            
            Console.CursorVisible = false;

            int lastWidth = Console.WindowWidth;
            int lastHeight = Console.WindowHeight;

            // Handle Initial Tab
            if (initialTab.HasValue && _excelDataSet != null)
            {
                int targetIndex = initialTab.Value - 1; // 1-based to 0-based
                if (targetIndex >= 0 && targetIndex < _excelDataSet.Tables.Count)
                {
                    LoadExcelSheet(targetIndex);
                    CalculateColumnWidths();
                    _statusMessage = $"Switched to Sheet {initialTab.Value}: {_excelDataSet.Tables[targetIndex].TableName}";
                }
                else
                {
                     _statusMessage = $"Sheet {initialTab.Value} not found. Showing Sheet 1.";
                }
            }

            // Handle Initial Search
            if (!string.IsNullOrEmpty(initialSearch))
            {
                _lastSearchTerm = initialSearch;
                FindText(initialSearch, true, 0);
            }

            // Initial Draw
            DrawUI();

            bool running = true;
            while (running)
            {
                // 1. Check for Resize
                if (Console.WindowWidth != lastWidth || Console.WindowHeight != lastHeight)
                {
                    lastWidth = Console.WindowWidth;
                    lastHeight = Console.WindowHeight;
                    Console.Clear();
                    DrawUI();
                }

                // 2. Check for Input
                if (Console.KeyAvailable)
                {
                    var key = Console.ReadKey(true);
                    running = HandleInput(key);
                    
                    // Only redraw if we are still running
                    if (running)
                    {
                        DrawUI();
                    }
                }

                // 3. Small delay to prevent high CPU usage
                Thread.Sleep(30);
            }
            
            Console.ResetColor();
            Console.Clear();
            Console.CursorVisible = true;
        }

        public void LoadFile(string path)
        {
            _fullFilePath = Path.GetFullPath(path);
            _fileName = Path.GetFileName(path);
            var fileInfo = new FileInfo(path);
            _fileSize = fileInfo.Length;
            _lastWriteTime = fileInfo.LastWriteTime;

            string ext = Path.GetExtension(path).ToLowerInvariant();

            if (ext == ".xls" || ext == ".xlsx")
            {
                // Register provider for old Excel (xls) support
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                ParseExcelFile(path);
                _encoding = null; // Not relevant/known for Excel
                _delimiter = '\0'; // Not relevant
            }
            else if (ext == ".ods")
            {
                ParseOdsFile(path);
                _encoding = null;
                _delimiter = '\0';
            }
            else
            {
                // 1. Detect Encoding
                _encoding = DetectEncoding(path);

                // 2. Read first few lines for delimiter sniffing
                var sampleLines = ReadSampleLines(path, _encoding, 5);
                _delimiter = DetectDelimiter(sampleLines);

                // 3. Parse File
                ParseFile(path, _encoding, _delimiter);
            }

            // 4. Calc Widths
            CalculateColumnWidths();
        }

        public void SwitchSheet(int index)
        {
            LoadExcelSheet(index);
        }

        public void SaveAsCsv(string path)
        {
            // Use semicolon ';' as separator for broader compatibility in many regions (e.g. DE)
            // or stick to standard comma. Given the tool context, we use semicolon.
            char separator = ';';
            
            using (var writer = new StreamWriter(path, false, new UTF8Encoding(false)))
            {
                foreach (var row in _data)
                {
                    var sb = new StringBuilder();
                    for (int i = 0; i < row.Length; i++)
                    {
                        string cell = row[i];
                        bool needsQuotes = false;
                        if (cell.Contains(separator) || cell.Contains('"') || cell.Contains('\n') || cell.Contains('\r'))
                        {
                            needsQuotes = true;
                        }
                        
                        if (needsQuotes)
                        {
                            sb.Append('"');
                            sb.Append(cell.Replace("\"", "\"\""));
                            sb.Append('"');
                        }
                        else
                        {
                            sb.Append(cell);
                        }
                        
                        if (i < row.Length - 1) sb.Append(separator);
                    }
                    writer.WriteLine(sb.ToString());
                }
            }
        }

        public void SaveAsExcel(string path)
        {
            ExcelExporter.ExportToFile(_data, path);
        }

        private void ParseOdsFile(string path)
        {
            _data.Clear();
            _excelDataSet = new DataSet();

            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                var contentEntry = archive.GetEntry("content.xml");
                if (contentEntry == null) throw new Exception("Invalid ODS file: content.xml not found.");

                using (var contentStream = contentEntry.Open())
                {
                    var doc = new XmlDocument();
                    doc.Load(contentStream);

                    var nsmgr = new XmlNamespaceManager(doc.NameTable);
                    nsmgr.AddNamespace("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
                    nsmgr.AddNamespace("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");

                    var tables = doc.SelectNodes("//table:table", nsmgr);
                    if (tables != null)
                    {
                        foreach (XmlNode tableNode in tables)
                        {
                            var tableName = tableNode.Attributes["table:name"]?.Value ?? "Sheet " + (_excelDataSet.Tables.Count + 1);
                            var dataTable = new DataTable(tableName);

                            var rows = tableNode.SelectNodes("table:table-row", nsmgr);
                            if (rows != null)
                            {
                                foreach (XmlNode rowNode in rows)
                                {
                                    var cells = new List<string>();
                                    var cellNodes = rowNode.SelectNodes("table:table-cell", nsmgr);

                                    if (cellNodes != null)
                                    {
                                        foreach (XmlNode cellNode in cellNodes)
                                        {
                                            var cellValue = cellNode.InnerText; 

                                            int repeat = 1;
                                            var repeatAttr = cellNode.Attributes["table:number-columns-repeated"];
                                            if (repeatAttr != null && int.TryParse(repeatAttr.Value, out int r))
                                            {
                                                repeat = r;
                                            }
                                            
                                            // Cap repeat to avoid OOM on empty trailing cells
                                            if (repeat > 1000) repeat = 1000;

                                            for (int i = 0; i < repeat; i++)
                                            {
                                                cells.Add(cellValue);
                                            }
                                        }
                                    }

                                    // Expand columns if needed
                                    while (dataTable.Columns.Count < cells.Count)
                                    {
                                        dataTable.Columns.Add();
                                    }

                                    var rowItemArray = new object[dataTable.Columns.Count];
                                    for (int i = 0; i < cells.Count; i++) rowItemArray[i] = cells[i];

                                    dataTable.Rows.Add(rowItemArray);
                                }
                            }
                            _excelDataSet.Tables.Add(dataTable);
                        }
                    }
                }
            }

            if (_excelDataSet.Tables.Count > 0)
            {
                LoadExcelSheet(0);
            }
            else
            {
                _totalRows = 0;
                _totalCols = 0;
            }
        }

        private void ParseExcelFile(string path)
        {
            _data.Clear();
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                try
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Read all sheets
                        _excelDataSet = reader.AsDataSet();
                    }
                }
                catch (Exception ex)
                {
                    // Catch encryption errors specifically or generic reader errors
                    if (ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase))
                    {
                        throw new Exception("File is encrypted (password protected). Opening not supported.");
                    }
                    throw; // Re-throw other errors
                }
            }

            if (_excelDataSet != null && _excelDataSet.Tables.Count > 0)
            {
                LoadExcelSheet(0);
            }
            else
            {
                _totalRows = 0;
                _totalCols = 0;
            }
        }

        private void LoadExcelSheet(int index)
        {
            if (_excelDataSet == null || index < 0 || index >= _excelDataSet.Tables.Count) return;

            _currentSheetIndex = index;
            _data.Clear();
            var table = _excelDataSet.Tables[index];
            
            foreach (System.Data.DataRow row in table.Rows)
            {
                var stringRow = row.ItemArray.Select(x => x?.ToString() ?? "").ToArray();
                _data.Add(stringRow);
            }

            _totalRows = _data.Count;
            _totalCols = _data.Count > 0 ? _data.Max(r => r.Length) : 0;
            
            // Normalize column counts
            for (int i = 0; i < _data.Count; i++)
            {
                if (_data[i].Length < _totalCols)
                {
                    var newRow = new string[_totalCols];
                    Array.Copy(_data[i], newRow, _data[i].Length);
                    for (int j = _data[i].Length; j < _totalCols; j++) newRow[j] = "";
                    _data[i] = newRow;
                }
            }
        }

        private Encoding DetectEncoding(string path)
        {
            using (var stream = File.OpenRead(path))
            {
                if (stream.Length >= 3)
                {
                    byte[] bom = new byte[3];
                    stream.Read(bom, 0, 3);
                    if (bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) return Encoding.UTF8;
                }
            }

            byte[] buffer = new byte[4096];
            using (var stream = File.OpenRead(path))
            {
                int read = stream.Read(buffer, 0, buffer.Length);
                if (IsUtf8(buffer, read)) return Encoding.UTF8;
            }

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding(1252);
        }

        private bool IsUtf8(byte[] buffer, int length)
        {
            int i = 0;
            while (i < length)
            {
                byte c = buffer[i];
                if (c < 0x80) i++;
                else if ((c & 0xE0) == 0xC0) { if (i + 1 >= length || (buffer[i + 1] & 0xC0) != 0x80) return false; i += 2; }
                else if ((c & 0xF0) == 0xE0) { if (i + 2 >= length || (buffer[i + 1] & 0xC0) != 0x80 || (buffer[i + 2] & 0xC0) != 0x80) return false; i += 3; }
                else if ((c & 0xF8) == 0xF0) { if (i + 3 >= length || (buffer[i + 1] & 0xC0) != 0x80 || (buffer[i + 2] & 0xC0) != 0x80 || (buffer[i + 3] & 0xC0) != 0x80) return false; i += 4; }
                else return false;
            }
            return true;
        }

        private List<string> ReadSampleLines(string path, Encoding enc, int count)
        {
            var lines = new List<string>();
            using (var reader = new StreamReader(path, enc))
            {
                string line;
                while ((line = reader.ReadLine()) != null && lines.Count < count) lines.Add(line);
            }
            return lines;
        }

        private char DetectDelimiter(List<string> lines)
        {
            if (lines.Count == 0) return ',';
            var candidates = new[] { ';', ',', '\t' };
            var counts = new Dictionary<char, int>();
            foreach (var c in candidates) counts[c] = 0;
            foreach (var line in lines) foreach (var c in candidates) counts[c] += line.Count(ch => ch == c);
            return counts.OrderByDescending(x => x.Value).First().Key;
        }

        private void ParseFile(string path, Encoding enc, char delimiter)
        {
            _data.Clear();
            using (var reader = new StreamReader(path, enc))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    if (line == null) break;
                    _data.Add(ParseLine(line, delimiter));
                }
            }
            _totalRows = _data.Count;
            _totalCols = _data.Count > 0 ? _data.Max(r => r.Length) : 0;
            
            for (int i = 0; i < _data.Count; i++)
            {
                if (_data[i].Length < _totalCols)
                {
                    var newRow = new string[_totalCols];
                    Array.Copy(_data[i], newRow, _data[i].Length);
                    for (int j = _data[i].Length; j < _totalCols; j++) newRow[j] = "";
                    _data[i] = newRow;
                }
            }
        }

        private string[] ParseLine(string line, char delimiter)
        {
            var result = new List<string>();
            var currentField = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { currentField.Append('"'); i++; }
                        else inQuotes = false;
                    }
                    else currentField.Append(c);
                }
                else
                {
                    if (c == '"') inQuotes = true;
                    else if (c == delimiter) { result.Add(currentField.ToString()); currentField.Clear(); }
                    else currentField.Append(c);
                }
            }
            result.Add(currentField.ToString());
            return result.ToArray();
        }

        private void CalculateColumnWidths()
        {
            if (_totalCols == 0) { _colWidths = new int[0]; return; }
            _colWidths = new int[_totalCols];
            int maxAllowedWidth = 50;
            int limit = Math.Min(_totalRows, 1000); 
            for (int col = 0; col < _totalCols; col++)
            {
                int maxLen = 0;
                for (int row = 0; row < limit; row++)
                {
                    if (_data[row].Length > col) maxLen = Math.Max(maxLen, _data[row][col].Length);
                }
                _colWidths[col] = Math.Clamp(maxLen, 5, maxAllowedWidth);
            }
        }

        private void DrawUI()
        {
            Console.SetCursorPosition(0, 0);
            int width = Console.WindowWidth;
            int height = Console.WindowHeight;

            DrawHeader(width);
            
            int dataRows = height - 2; 
            if (dataRows < 1) dataRows = 1;

            DrawGrid(dataRows, width);

            try { Console.SetCursorPosition(0, height - 1); } catch { }
            DrawFooter(width);
        }

        private void DrawHeader(int width)
        {
            string sizeStr = FormatBytes(_fileSize);
            string encStr = _encoding?.EncodingName ?? "N/A";
            if (encStr.Length > 15 && _encoding != null) encStr = _encoding.HeaderName;
            string dateStr = _lastWriteTime.ToString("g");

            string sepChar = _delimiter == '\0' ? "N/A" : $"'{_delimiter}'";

            string headerText = $" FILE: {_fileName} | SIZE: {sizeStr} | DATE: {dateStr} | DIM: {_totalRows}x{_totalCols} | ENC: {encStr} | SEP: {sepChar}";
            
            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            if (headerText.Length > width) headerText = headerText.Substring(0, width);
            Console.Write(headerText.PadRight(width));
            Console.ResetColor();
        }

        private void DrawFooter(int width)
        {
            string helpText = " Arrows/Pg/Home/End: Move | 'f': Find | F3: Next | Shift+F3: Prev | 'l': LibreOffice | 'e': Excel";
            
            if (_excelDataSet != null && _excelDataSet.Tables.Count > 1)
            {
                helpText += " | 1-9: Sheets";
            }

            helpText += " | ESC/q: Quit";

            if (!string.IsNullOrEmpty(_statusMessage))
            {
                helpText = " " + _statusMessage;
            }

            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            
            // Critical fix for cmd.exe: Don't write to the very last character of the last line
            // to prevent scrolling.
            int safeWidth = width - 1;
            if (helpText.Length > safeWidth) helpText = helpText.Substring(0, safeWidth);
            Console.Write(helpText.PadRight(safeWidth));
            
            Console.ResetColor();
        }

        private void DrawGrid(int maxRows, int consoleWidth)
        {
            if (_totalRows == 0) return;

            var visibleCols = new List<int>();
            int currentWidth = 0;
            
            for (int c = _scrollCol; c < _totalCols; c++)
            {
                int colW = _colWidths[c] + 1; // +1 for separator
                if (currentWidth + colW > consoleWidth)
                {
                    if (visibleCols.Count == 0) visibleCols.Add(c);
                    break;
                }
                currentWidth += colW;
                visibleCols.Add(c);
            }

            Console.SetCursorPosition(0, 1);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write(GetRowString(_data[0], visibleCols, consoleWidth));
            Console.ResetColor();
            
            int dataAreaRows = maxRows - 1; 
            if (dataAreaRows < 1) return;

            for (int r = 0; r < dataAreaRows; r++)
            {
                int dataIndex = 1 + _scrollRow + r;
                int targetY = 2 + r;
                
                if (targetY >= Console.WindowHeight - 1) break; // Avoid footer overlap

                Console.SetCursorPosition(0, targetY);

                if (dataIndex < _totalRows)
                {
                    if (dataIndex == _highlightRow)
                    {
                        Console.BackgroundColor = ConsoleColor.Yellow;
                        Console.ForegroundColor = ConsoleColor.Black;
                    }

                    Console.Write(GetRowString(_data[dataIndex], visibleCols, consoleWidth));

                    if (dataIndex == _highlightRow)
                    {
                        Console.ResetColor();
                    }
                }
                else
                {
                    Console.Write(new string(' ', consoleWidth)); 
                }
            }
        }

        private string GetRowString(string[] rowData, List<int> visibleCols, int consoleWidth)
        {
            var lineBuilder = new StringBuilder();

            foreach (int colIndex in visibleCols)
            {
                string cell = colIndex < rowData.Length ? rowData[colIndex] : "";
                int w = _colWidths[colIndex];
                
                if (cell.Length > w) cell = cell.Substring(0, w - 3) + "...";
                
                lineBuilder.Append(cell.PadRight(w));
                lineBuilder.Append("|"); // Separator
            }
            
            string lineStr = lineBuilder.ToString();
            // Pad or truncate
            if (lineStr.Length > consoleWidth) lineStr = lineStr.Substring(0, consoleWidth);
            else lineStr = lineStr.PadRight(consoleWidth);
            
            return lineStr;
        }

        private bool HandleInput(ConsoleKeyInfo key)
        {
            _statusMessage = ""; 
            int dataRowsCount = _totalRows - 1; 
            if (dataRowsCount < 0) dataRowsCount = 0;
            
            int viewportHeight = Console.WindowHeight - 3; 
            if (viewportHeight < 1) viewportHeight = 1;

            switch (key.Key)
            {
                case ConsoleKey.Escape:
                case ConsoleKey.Q:
                    return false;

                case ConsoleKey.UpArrow:
                    if (_scrollRow > 0) _scrollRow--;
                    break;
                case ConsoleKey.DownArrow:
                    if (_scrollRow < dataRowsCount - 1) _scrollRow++;
                    break;
                case ConsoleKey.PageUp:
                    _scrollRow = Math.Max(0, _scrollRow - viewportHeight);
                    break;
                case ConsoleKey.PageDown:
                    _scrollRow = Math.Min(Math.Max(0, dataRowsCount - 1), _scrollRow + viewportHeight);
                    break;
                case ConsoleKey.Home:
                    _scrollRow = 0;
                    break;
                case ConsoleKey.End:
                    _scrollRow = Math.Max(0, dataRowsCount - 1);
                    break;

                case ConsoleKey.LeftArrow:
                    if (_scrollCol > 0) _scrollCol--;
                    break;
                case ConsoleKey.RightArrow:
                    if (_scrollCol < _totalCols - 1) _scrollCol++;
                    break;

                case ConsoleKey.L:
                    LaunchLibreOffice();
                    break;
                case ConsoleKey.E:
                    LaunchExcel();
                    break;

                case ConsoleKey.D1:
                case ConsoleKey.D2:
                case ConsoleKey.D3:
                case ConsoleKey.D4:
                case ConsoleKey.D5:
                case ConsoleKey.D6:
                case ConsoleKey.D7:
                case ConsoleKey.D8:
                case ConsoleKey.D9:
                    if (_excelDataSet != null)
                    {
                        int sheetIndex = key.Key - ConsoleKey.D1;
                        if (sheetIndex < _excelDataSet.Tables.Count)
                        {
                            if (sheetIndex != _currentSheetIndex)
                            {
                                LoadExcelSheet(sheetIndex);
                                CalculateColumnWidths();
                                _scrollRow = 0;
                                _scrollCol = 0;
                                _statusMessage = $"Switched to Sheet {sheetIndex + 1}: {_excelDataSet.Tables[sheetIndex].TableName}";
                            }
                            else
                            {
                                _statusMessage = $"Already on Sheet {sheetIndex + 1}: {_excelDataSet.Tables[sheetIndex].TableName}";
                            }
                        }
                        else
                        {
                             _statusMessage = $"Sheet {sheetIndex + 1} does not exist.";
                        }
                    }
                    break;
                
                case ConsoleKey.F:
                    ShowSearchDialog();
                    break;
                case ConsoleKey.F3:
                    if ((key.Modifiers & ConsoleModifiers.Shift) != 0)
                        FindText(_lastSearchTerm, false);
                    else
                        FindText(_lastSearchTerm, true);
                    break;
            }
            return true;
        }

        private void ShowSearchDialog()
        {
            int h = Console.WindowHeight;
            int w = Console.WindowWidth;
            Console.SetCursorPosition(0, h - 1);
            Console.BackgroundColor = ConsoleColor.Blue;
            Console.ForegroundColor = ConsoleColor.White;
            // Prevent scroll by using w - 1
            Console.Write(" Search: ".PadRight(w - 1));
            Console.SetCursorPosition(9, h - 1);
            
            Console.CursorVisible = true;
            string term = Console.ReadLine();
            Console.CursorVisible = false;
            Console.ResetColor();

            if (!string.IsNullOrWhiteSpace(term))
            {
                _lastSearchTerm = term;
                FindText(term, true);
            }
        }

        private void FindText(string term, bool forward, int? startRowOverride = null)
        {
            if (string.IsNullOrEmpty(term))
            {
                _statusMessage = "No search term.";
                return;
            }

            int maxScrollRow = _totalRows - 2;
            if (maxScrollRow < 0) return; // No data rows

            int startRow;
            if (startRowOverride.HasValue)
            {
                startRow = startRowOverride.Value;
            }
            else
            {
                // Determine start based on current highlight if available, else scroll pos
                if (_highlightRow != -1)
                {
                    // _highlightRow corresponds to r (index) + 1.
                    // If highlight is at r=5 (row 6), we want to start at r=6 for forward.
                    // _highlightRow is 6. So startRow = 6.
                    startRow = forward ? _highlightRow : _highlightRow - 2;
                }
                else
                {
                    startRow = forward ? _scrollRow + 1 : _scrollRow - 1;
                }
            }

            // Loop range must be within valid scroll rows [0 .. maxScrollRow]
            int foundRow = -1;

            if (forward)
            {
                for (int r = startRow; r <= maxScrollRow; r++)
                {
                    if (r < 0) continue; 
                    // _scrollRow r corresponds to _data[r + 1]
                    if (RowContains(_data[r + 1], term)) 
                    {
                        foundRow = r;
                        break;
                    }
                }
            }
            else
            {
                for (int r = startRow; r >= 0; r--)
                {
                    if (r > maxScrollRow) continue;
                    if (RowContains(_data[r + 1], term))
                    {
                        foundRow = r;
                        break;
                    }
                }
            }

            if (foundRow != -1)
            {
                // If it's a "Find Next", we generally just want to ensure it's visible.
                // If it's an initial find (startRowOverride is 0), we want context.
                // But even for Find Next, context is nice if it jumps far.
                
                // Simple logic: If foundRow is outside current view, center it or show context.
                int viewportHeight = Console.WindowHeight - 3;
                bool isVisible = foundRow >= _scrollRow && foundRow < _scrollRow + viewportHeight;

                if (!isVisible || startRowOverride.HasValue) // Force context on initial search or jump
                {
                    int contextOffset = 5;
                    _scrollRow = Math.Max(0, foundRow - contextOffset);
                }
                // If is visible, we don't strictly *need* to scroll.

                _highlightRow = foundRow + 1;
                _statusMessage = $"Found '{term}' at row {foundRow + 1}";
            }
            else
            {
                _highlightRow = -1;
                _statusMessage = $"'{term}' not found.";
            }
        }

        private bool RowContains(string[] row, string term)
        {
            foreach (var cell in row)
            {
                if (cell.Contains(term, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private void LaunchExcel()
        {
            _statusMessage = "starting Excel...";
            DrawUI(); 

            // 1. Try via PATH "excel"
            if (StartProcess("excel", $"\"{_fullFilePath}\" ")) return;

            // 2. Search Common Paths
            string[] paths = {
                @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
            };

            foreach(var p in paths)
            {
                if (File.Exists(p))
                {
                    if (StartProcess(p, $"\"{_fullFilePath}\" ")) return;
                }
            }

            _statusMessage = "Excel not found.";
        }

        private void LaunchLibreOffice()
        {
            _statusMessage = "starting LibreOffice...";
            DrawUI();

            // 1. Try via PATH "scalc" or "soffice"
            if (StartProcess("scalc", $"\"{_fullFilePath}\" ")) return;

            // 2. Common Paths
            string[] paths = {
                @"C:\Program Files\LibreOffice\program\scalc.exe",
                @"C:\Program Files (x86)\LibreOffice\program\scalc.exe"
            };

             foreach(var p in paths)
            {
                if (File.Exists(p))
                {
                    if (StartProcess(p, $"\"{_fullFilePath}\" ")) return;
                }
            }

            _statusMessage = "LibreOffice not found.";
        }

        private bool StartProcess(string exe, string args)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = args,
                    UseShellExecute = true 
                });
                return true;
            }
            catch
            {
                return false;
            }
        }

        private string FormatBytes(long bytes)
        {
            string[] suffix = { "B", "KB", "MB", "GB" };
            int i;
            double dblSByte = bytes;
            for (i = 0; i < suffix.Length && bytes >= 1024; i++, bytes /= 1024)
            {
                dblSByte = bytes / 1024.0;
            }
            return String.Format("{0:0.##} {1}", dblSByte, suffix[i]);
        }
    }
}
