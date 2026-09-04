using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace CsvTool.Core
{
    public class FileData
    {
        public string FullPath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public DateTime LastWriteTime { get; set; }
        public Encoding? Encoding { get; set; }
        public char Delimiter { get; set; }
        
        // The current active data
        public List<string[]> Rows { get; set; } = new List<string[]>();
        public int TotalRows => Rows.Count;
        public int TotalCols { get; set; }
        public int[]? ColWidths { get; set; }

        // Show the column number behind each header name (option -n)
        public bool ShowColumnNumbers { get; set; }
        /// <summary>1-based column numbers of the source file; survives a column selection via -c.</summary>
        public int[]? SourceColumnNumbers { get; set; }

        // Show the line number in a fixed gutter on the left (option -n)
        public bool ShowLineNumbers { get; set; }
        /// <summary>1-based row/line numbers of the source file; survives a line selection via -l.</summary>
        public int[]? SourceRowNumbers { get; set; }

        // Excel/ODS specific
        public DataSet? FullDataSet { get; set; }
        public int CurrentSheetIndex { get; set; } = 0;
    }
}
