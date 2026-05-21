using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using CsvTool.Core;
using CsvTool.Tui;
using CsvTool.Platform;

namespace CsvTool
{
    public class CsvViewer
    {
        private FileData _data = new FileData();
        private int _scrollRow = 0;
        private int _scrollCol = 0;
        private string _statusMessage = "";
        private string _lastSearchTerm = "";
        private int _highlightRow = -1;

        public void Run(string path, string? initialSearch = null, int? initialTab = null, List<int>? selectedCols = null)
        {
            Console.Clear();
            LoadFile(path);
            
            Console.CursorVisible = false;

            int lastWidth = Console.WindowWidth;
            int lastHeight = Console.WindowHeight;

            // Handle Initial Tab
            if (initialTab.HasValue && _data.FullDataSet != null)
            {
                int targetIndex = initialTab.Value - 1;
                if (targetIndex >= 0 && targetIndex < _data.FullDataSet.Tables.Count)
                {
                    LoadExcelSheet(targetIndex);
                    CalculateColumnWidths();
                    _statusMessage = $"Switched to Sheet {initialTab.Value}: {_data.FullDataSet.Tables[targetIndex].TableName}";
                }
                else
                {
                     _statusMessage = $"Sheet {initialTab.Value} not found. Showing Sheet 1.";
                }
            }

            if (selectedCols != null && selectedCols.Count > 0)
            {
                ApplyColumnSelection(selectedCols);
            }

            // Handle Initial Search
            if (!string.IsNullOrEmpty(initialSearch))
            {
                _lastSearchTerm = initialSearch;
                FindText(initialSearch, true, 0);
            }

            Renderer.DrawUI(_data, _statusMessage, _scrollRow, _scrollCol, _highlightRow);

            bool running = true;
            while (running)
            {
                if (Console.WindowWidth != lastWidth || Console.WindowHeight != lastHeight)
                {
                    lastWidth = Console.WindowWidth;
                    lastHeight = Console.WindowHeight;
                    Console.Clear();
                    Renderer.DrawUI(_data, _statusMessage, _scrollRow, _scrollCol, _highlightRow);
                }

                if (Console.KeyAvailable)
                {
                    var key = Console.ReadKey(true);
                    running = HandleInput(key);
                    
                    if (running)
                    {
                        Renderer.DrawUI(_data, _statusMessage, _scrollRow, _scrollCol, _highlightRow);
                    }
                }

                Thread.Sleep(30);
            }
            
            Console.ResetColor();
            Console.Clear();
            Console.CursorVisible = true;
        }

        public void LoadFile(string path)
        {
            _data.FullPath = Path.GetFullPath(path);
            _data.FileName = Path.GetFileName(path);
            var fileInfo = new FileInfo(path);
            _data.FileSize = fileInfo.Length;
            _data.LastWriteTime = fileInfo.LastWriteTime;

            string ext = Path.GetExtension(path).ToLowerInvariant();

            if (ext == ".xls" || ext == ".xlsx")
            {
                _data.FullDataSet = ExcelHandler.ReadExcel(path);
                _data.Encoding = null;
                _data.Delimiter = '\0';
                LoadExcelSheet(0);
            }
            else if (ext == ".ods")
            {
                _data.FullDataSet = OdsParser.Parse(path);
                _data.Encoding = null;
                _data.Delimiter = '\0';
                LoadExcelSheet(0);
            }
            else
            {
                _data.Encoding = EncodingDetector.Detect(path);
                _data.Delimiter = CsvParser.DetectDelimiter(path, _data.Encoding);
                _data.Rows = CsvParser.Parse(path, _data.Encoding, _data.Delimiter);
                
                _data.TotalCols = _data.Rows.Count > 0 ? _data.Rows.Max(r => r.Length) : 0;
                NormalizeRows();
            }

            CalculateColumnWidths();
        }

        private void NormalizeRows()
        {
            for (int i = 0; i < _data.Rows.Count; i++)
            {
                if (_data.Rows[i].Length < _data.TotalCols)
                {
                    var newRow = new string[_data.TotalCols];
                    Array.Copy(_data.Rows[i], newRow, _data.Rows[i].Length);
                    for (int j = _data.Rows[i].Length; j < _data.TotalCols; j++) newRow[j] = "";
                    _data.Rows[i] = newRow;
                }
            }
        }

        private void LoadExcelSheet(int index)
        {
            if (_data.FullDataSet == null || index < 0 || index >= _data.FullDataSet.Tables.Count) return;

            _data.CurrentSheetIndex = index;
            _data.Rows.Clear();
            var table = _data.FullDataSet.Tables[index];
            
            foreach (System.Data.DataRow row in table.Rows)
            {
                var stringRow = row.ItemArray.Select(x => x?.ToString() ?? "").ToArray();
                _data.Rows.Add(stringRow);
            }

            _data.TotalCols = _data.Rows.Count > 0 ? _data.Rows.Max(r => r.Length) : 0;
            NormalizeRows();
        }

        private void CalculateColumnWidths()
        {
            if (_data.TotalCols == 0) { _data.ColWidths = Array.Empty<int>(); return; }
            _data.ColWidths = new int[_data.TotalCols];
            int maxAllowedWidth = 50;
            int limit = Math.Min(_data.TotalRows, 1000); 
            for (int col = 0; col < _data.TotalCols; col++)
            {
                int maxLen = 0;
                for (int row = 0; row < limit; row++)
                {
                    if (_data.Rows[row].Length > col) maxLen = Math.Max(maxLen, _data.Rows[row][col].Length);
                }
                _data.ColWidths[col] = Math.Clamp(maxLen, 5, maxAllowedWidth);
            }
        }

        public void SwitchSheet(int index)
        {
            LoadExcelSheet(index);
            CalculateColumnWidths();
        }

        public void SaveAsCsv(string path)
        {
            char separator = ';';
            using (var writer = new StreamWriter(path, false, new System.Text.UTF8Encoding(false)))
            {
                foreach (var row in _data.Rows)
                {
                    var sb = new System.Text.StringBuilder();
                    for (int i = 0; i < row.Length; i++)
                    {
                        string cell = row[i];
                        bool needsQuotes = cell.Contains(separator) || cell.Contains('"') || cell.Contains('\n') || cell.Contains('\r');
                        if (needsQuotes)
                        {
                            sb.Append('"').Append(cell.Replace("\"", "\"\"")).Append('"');
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
            ExcelHandler.SaveAsExcel(_data.Rows, path);
        }

        public void SaveAsOds(string path)
        {
            OdsWriter.SaveAsOds(_data.Rows, path);
        }

        public void ApplyColumnSelection(List<int> selectedCols)
        {
            if (selectedCols == null || selectedCols.Count == 0) return;
            _data.Rows = ColumnFilter.Apply(_data.Rows, selectedCols);
            _data.TotalCols = selectedCols.Count;
            if (_data.ColWidths != null)
            {
                var newWidths = new int[_data.TotalCols];
                for (int i = 0; i < selectedCols.Count; i++)
                {
                    int oldIdx = selectedCols[i];
                    newWidths[i] = oldIdx < _data.ColWidths.Length ? _data.ColWidths[oldIdx] : 10;
                }
                _data.ColWidths = newWidths;
            }
            else
            {
                CalculateColumnWidths();
            }
        }

        private bool HandleInput(ConsoleKeyInfo key)
        {
            _statusMessage = ""; 
            int dataRowsCount = _data.TotalRows - 1; 
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
                    if (_scrollCol < _data.TotalCols - 1) _scrollCol++;
                    break;

                case ConsoleKey.L:
                    _statusMessage = "starting LibreOffice...";
                    Renderer.DrawUI(_data, _statusMessage, _scrollRow, _scrollCol, _highlightRow);
                    if (!ExternalLauncher.LaunchLibreOffice(_data.FullPath)) _statusMessage = "LibreOffice not found.";
                    break;
                case ConsoleKey.E:
                    _statusMessage = "starting Excel...";
                    Renderer.DrawUI(_data, _statusMessage, _scrollRow, _scrollCol, _highlightRow);
                    if (!ExternalLauncher.LaunchExcel(_data.FullPath)) _statusMessage = "Excel not found.";
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
                    if (_data.FullDataSet != null)
                    {
                        int sheetIndex = key.Key - ConsoleKey.D1;
                        if (sheetIndex < _data.FullDataSet.Tables.Count)
                        {
                            if (sheetIndex != _data.CurrentSheetIndex)
                            {
                                LoadExcelSheet(sheetIndex);
                                CalculateColumnWidths();
                                _scrollRow = 0;
                                _scrollCol = 0;
                                _statusMessage = $"Switched to Sheet {sheetIndex + 1}: {_data.FullDataSet.Tables[sheetIndex].TableName}";
                            }
                            else
                            {
                                _statusMessage = $"Already on Sheet {sheetIndex + 1}: {_data.FullDataSet.Tables[sheetIndex].TableName}";
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
            Console.Write(" Search: ".PadRight(w - 1));
            Console.SetCursorPosition(9, h - 1);
            
            Console.CursorVisible = true;
            string? term = Console.ReadLine();
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

            int maxScrollRow = _data.TotalRows - 2;
            if (maxScrollRow < 0) return;

            int startRow;
            if (startRowOverride.HasValue)
            {
                startRow = startRowOverride.Value;
            }
            else
            {
                if (_highlightRow != -1)
                {
                    startRow = forward ? _highlightRow : _highlightRow - 2;
                }
                else
                {
                    startRow = forward ? _scrollRow + 1 : _scrollRow - 1;
                }
            }

            int foundRow = -1;

            if (forward)
            {
                for (int r = startRow; r <= maxScrollRow; r++)
                {
                    if (r < 0) continue; 
                    if (RowContains(_data.Rows[r + 1], term)) 
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
                    if (RowContains(_data.Rows[r + 1], term))
                    {
                        foundRow = r;
                        break;
                    }
                }
            }

            if (foundRow != -1)
            {
                int viewportHeight = Console.WindowHeight - 3;
                bool isVisible = foundRow >= _scrollRow && foundRow < _scrollRow + viewportHeight;

                if (!isVisible || startRowOverride.HasValue)
                {
                    int contextOffset = 5;
                    _scrollRow = Math.Max(0, foundRow - contextOffset);
                }

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
            return row.Any(cell => cell.Contains(term, StringComparison.OrdinalIgnoreCase));
        }

        // Expose metadata for info mode
        public FileData Data => _data;
    }
}
