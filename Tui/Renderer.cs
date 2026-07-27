using System;
using System.Collections.Generic;
using System.Text;
using CsvTool.Core;

namespace CsvTool.Tui
{
    public static class Renderer
    {
        public static void DrawUI(FileData data, string statusMessage, int scrollRow, int scrollCol, int highlightRow)
        {
            Console.SetCursorPosition(0, 0);
            int width = Console.WindowWidth;
            int height = Console.WindowHeight;

            DrawHeader(data, width);
            
            int dataRows = height - 2; 
            if (dataRows < 1) dataRows = 1;

            DrawGrid(data, dataRows, width, scrollRow, scrollCol, highlightRow);

            try { Console.SetCursorPosition(0, height - 1); } catch { }
            DrawFooter(data, statusMessage, width);
        }

        private static void DrawHeader(FileData data, int width)
        {
            string sizeStr = FormatBytes(data.FileSize);
            string encStr = data.Encoding?.EncodingName ?? "N/A";
            if (encStr.Length > 15 && data.Encoding != null) encStr = data.Encoding.HeaderName;
            string dateStr = data.LastWriteTime.ToString("g");

            string sepChar = data.Delimiter == '\0' ? "N/A" : $"'{data.Delimiter}'";

            string headerText = $" FILE: {data.FileName} | SIZE: {sizeStr} | DATE: {dateStr} | DIM: {data.TotalRows}x{data.TotalCols} | ENC: {encStr} | SEP: {sepChar}";
            
            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            if (headerText.Length > width) headerText = headerText.Substring(0, width);
            Console.Write(headerText.PadRight(width));
            Console.ResetColor();
        }

        private static void DrawFooter(FileData data, string statusMessage, int width)
        {
            string helpText = " Arrows/Pg/Home/End: Move | 'f': Find | F3: Next | Shift+F3: Prev | 'l': LibreOffice | 'e': Excel";
            
            if (data.FullDataSet != null && data.FullDataSet.Tables.Count > 1)
            {
                helpText += " | 1-9: Sheets";
            }

            helpText += " | ESC/q: Quit";

            if (!string.IsNullOrEmpty(statusMessage))
            {
                helpText = " " + statusMessage;
            }

            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            
            int safeWidth = width - 1;
            if (helpText.Length > safeWidth) helpText = helpText.Substring(0, safeWidth);
            Console.Write(helpText.PadRight(safeWidth));
            
            Console.ResetColor();
        }

        private static void DrawGrid(FileData data, int maxRows, int consoleWidth, int scrollRow, int scrollCol, int highlightRow)
        {
            if (data.TotalRows == 0 || data.ColWidths == null) return;

            var visibleCols = new List<int>();
            int currentWidth = 0;
            
            for (int c = scrollCol; c < data.TotalCols; c++)
            {
                int colW = data.ColWidths[c] + 1; 
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
            Console.Write(GetRowString(data, data.Rows[0], visibleCols, consoleWidth));
            Console.ResetColor();
            
            int dataAreaRows = maxRows - 1; 
            if (dataAreaRows < 1) return;

            for (int r = 0; r < dataAreaRows; r++)
            {
                int dataIndex = 1 + scrollRow + r;
                int targetY = 2 + r;
                
                if (targetY >= Console.WindowHeight - 1) break; 

                Console.SetCursorPosition(0, targetY);

                if (dataIndex < data.TotalRows)
                {
                    if (dataIndex == highlightRow)
                    {
                        Console.BackgroundColor = ConsoleColor.Yellow;
                        Console.ForegroundColor = ConsoleColor.Black;
                    }

                    Console.Write(GetRowString(data, data.Rows[dataIndex], visibleCols, consoleWidth));

                    if (dataIndex == highlightRow)
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

        private static string GetRowString(FileData data, string[] rowData, List<int> visibleCols, int consoleWidth)
        {
            var lineBuilder = new StringBuilder();

            foreach (int colIndex in visibleCols)
            {
                string cell = Flatten(colIndex < rowData.Length ? rowData[colIndex] : "");
                int w = data.ColWidths![colIndex];

                if (cell.Length > w) cell = cell.Substring(0, w - 3) + "...";
                
                lineBuilder.Append(cell.PadRight(w));
                lineBuilder.Append("|"); 
            }
            
            string lineStr = lineBuilder.ToString();
            if (lineStr.Length > consoleWidth) lineStr = lineStr.Substring(0, consoleWidth);
            else lineStr = lineStr.PadRight(consoleWidth);
            
            return lineStr;
        }

        /// <summary>Line breaks inside a cell would break the grid layout, so show them as spaces.</summary>
        private static string Flatten(string cell)
        {
            if (cell.IndexOfAny(new[] { '\n', '\r', '\t' }) < 0) return cell;
            return cell.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
        }

        private static string FormatBytes(long bytes)
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
