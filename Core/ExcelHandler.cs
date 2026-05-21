using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;
using ClosedXML.Excel;

namespace CsvTool.Core
{
    public static class ExcelHandler
    {
        public static DataSet ReadExcel(string path)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                try
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        return reader.AsDataSet();
                    }
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("password", StringComparison.OrdinalIgnoreCase))
                    {
                        throw new Exception("File is encrypted (password protected). Opening not supported.");
                    }
                    throw;
                }
            }
        }

        public static void SaveAsExcel(IList<string[]> data, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");
                for (int row = 0; row < data.Count; row++)
                {
                    var rowData = data[row];
                    for (int col = 0; col < rowData.Length; col++)
                    {
                        worksheet.Cell(row + 1, col + 1).Value = rowData[col];
                    }
                }
                workbook.SaveAs(filePath);
            }
        }
    }
}
