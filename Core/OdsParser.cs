using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace CsvTool.Core
{
    public static class OdsParser
    {
        public static DataSet Parse(string path)
        {
            var dataSet = new DataSet();

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
                            var tableName = tableNode.Attributes?["table:name"]?.Value ?? "Sheet " + (dataSet.Tables.Count + 1);
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
                                            var repeatAttr = cellNode.Attributes?["table:number-columns-repeated"];
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
                            dataSet.Tables.Add(dataTable);
                        }
                    }
                }
            }

            return dataSet;
        }
    }
}
