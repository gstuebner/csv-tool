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

        // Excel/ODS specific
        public DataSet? FullDataSet { get; set; }
        public int CurrentSheetIndex { get; set; } = 0;
    }
}
