using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;

namespace CsvTool
{
    public static class ExcelExporter
    {
        public static void ExportToFile(IList<string[]> data, string filePath)
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
