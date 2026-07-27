using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace CsvTool.Core
{
    public static class OdsWriter
    {
        public static void SaveAsOds(IList<string[]> data, string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false))
            {
                AddEntry(zip, "mimetype", "application/vnd.oasis.opendocument.spreadsheet", CompressionLevel.NoCompression);

                string manifest = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<manifest:manifest xmlns:manifest=""urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"" manifest:version=""1.2"">
  <manifest:file-entry manifest:full-path=""/"" manifest:media-type=""application/vnd.oasis.opendocument.spreadsheet""/>
  <manifest:file-entry manifest:full-path=""content.xml"" manifest:media-type=""text/xml""/>
  <manifest:file-entry manifest:full-path=""styles.xml"" manifest:media-type=""text/xml""/>
  <manifest:file-entry manifest:full-path=""meta.xml"" manifest:media-type=""text/xml""/>
</manifest:manifest>";
                AddEntry(zip, "META-INF/manifest.xml", manifest);

                string styles = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<office:document-styles xmlns:office=""urn:oasis:names:tc:opendocument:xmlns:office:1.0"" office:version=""1.2"">
  <office:styles/>
</office:document-styles>";
                AddEntry(zip, "styles.xml", styles);

                string meta = $@"<?xml version=""1.0"" encoding=""UTF-8""?>
<office:document-meta xmlns:office=""urn:oasis:names:tc:opendocument:xmlns:office:1.0"" xmlns:meta=""urn:oasis:names:tc:opendocument:xmlns:meta:1.0"" office:version=""1.2"">
  <office:meta>
    <meta:creation-date>{DateTime.Now:yyyy-MM-ddTHH:mm:ss}</meta:creation-date>
    <meta:generator>csv-tool</meta:generator>
  </office:meta>
</office:document-meta>";
                AddEntry(zip, "meta.xml", meta);

                var sb = new StringBuilder();
                sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
                sb.AppendLine(@"<office:document-content xmlns:office=""urn:oasis:names:tc:opendocument:xmlns:office:1.0"" xmlns:table=""urn:oasis:names:tc:opendocument:xmlns:table:1.0"" xmlns:text=""urn:oasis:names:tc:opendocument:xmlns:text:1.0"" office:version=""1.2"">");
                sb.AppendLine(@"  <office:body>");
                sb.AppendLine(@"    <office:spreadsheet>");
                sb.AppendLine(@"      <table:table table:name=""Sheet1"">");

                foreach (var row in data)
                {
                    sb.AppendLine(@"        <table:table-row>");
                    foreach (var cell in row)
                    {
                        sb.AppendLine($@"          <table:table-cell office:value-type=""string"">{CellParagraphs(cell)}</table:table-cell>");
                    }
                    sb.AppendLine(@"        </table:table-row>");
                }

                sb.AppendLine(@"      </table:table>");
                sb.AppendLine(@"    </office:spreadsheet>");
                sb.AppendLine(@"  </office:body>");
                sb.AppendLine(@"</office:document-content>");

                AddEntry(zip, "content.xml", sb.ToString());
            }
        }

        private static void AddEntry(ZipArchive zip, string entryName, string content, CompressionLevel level = CompressionLevel.Optimal)
        {
            var entry = zip.CreateEntry(entryName, level);
            using (var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false)))
            {
                writer.Write(content);
            }
        }

        /// <summary>
        /// A line break inside a cell becomes its own text:p, which is how ODF represents
        /// multi-line cell content.
        /// </summary>
        private static string CellParagraphs(string cell)
        {
            var lines = cell.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var sb = new StringBuilder();
            foreach (var line in lines) sb.Append("<text:p>").Append(EscapeXml(line)).Append("</text:p>");
            return sb.ToString();
        }

        private static string EscapeXml(string text)
        {
            return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }
    }
}
