using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CsvTool.Core
{
    public static class CsvParser
    {
        public static char DetectDelimiter(string path, Encoding enc)
        {
            var candidates = new[] { ';', ',', '\t' };
            var counts = new Dictionary<char, int>();
            foreach (var c in candidates) counts[c] = 0;

            bool any = false;
            using (var reader = new StreamReader(path, enc))
            {
                bool inQuotes = false;
                int records = 0;
                int ch;
                while ((ch = reader.Read()) != -1 && records < 5)
                {
                    any = true;
                    char c = (char)ch;
                    if (inQuotes)
                    {
                        if (c == '"')
                        {
                            if (reader.Peek() == '"') reader.Read();
                            else inQuotes = false;
                        }
                    }
                    else
                    {
                        if (c == '"') inQuotes = true;
                        else if (c == '\r' || c == '\n')
                        {
                            if (c == '\r' && reader.Peek() == '\n') reader.Read();
                            records++;
                        }
                        else if (counts.ContainsKey(c)) counts[c]++;
                    }
                }
            }

            if (!any) return ',';
            return counts.OrderByDescending(x => x.Value).First().Key;
        }

        public static List<string[]> Parse(string path, Encoding enc, char delimiter)
        {
            using (var reader = new StreamReader(path, enc))
            {
                return ParseRecords(reader, delimiter);
            }
        }

        /// <summary>
        /// Reads CSV records character by character (RFC 4180), so quoted fields may
        /// span multiple physical lines. Line breaks inside a field are normalized to '\n'.
        /// </summary>
        private static List<string[]> ParseRecords(TextReader reader, char delimiter)
        {
            var records = new List<string[]>();
            var fields = new List<string>();
            var field = new StringBuilder();
            bool inQuotes = false;
            bool pending = false; // something was read since the last record was emitted

            int ch;
            while ((ch = reader.Read()) != -1)
            {
                char c = (char)ch;
                pending = true;

                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (reader.Peek() == '"') { reader.Read(); field.Append('"'); }
                        else inQuotes = false;
                    }
                    else if (c == '\r')
                    {
                        if (reader.Peek() == '\n') reader.Read();
                        field.Append('\n');
                    }
                    else field.Append(c);
                }
                else
                {
                    if (c == '"') inQuotes = true;
                    else if (c == delimiter) { fields.Add(field.ToString()); field.Clear(); }
                    else if (c == '\r' || c == '\n')
                    {
                        if (c == '\r' && reader.Peek() == '\n') reader.Read();
                        fields.Add(field.ToString());
                        field.Clear();
                        records.Add(fields.ToArray());
                        fields.Clear();
                        pending = false;
                    }
                    else field.Append(c);
                }
            }

            // Trailing record without a final line break (an unterminated quote ends here too).
            if (pending)
            {
                fields.Add(field.ToString());
                records.Add(fields.ToArray());
            }

            return records;
        }
    }
}
