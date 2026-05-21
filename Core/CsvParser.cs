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
            var lines = new List<string>();
            using (var reader = new StreamReader(path, enc))
            {
                string? line;
                while ((line = reader.ReadLine()) != null && lines.Count < 5) lines.Add(line);
            }

            if (lines.Count == 0) return ',';
            var candidates = new[] { ';', ',', '\t' };
            var counts = new Dictionary<char, int>();
            foreach (var c in candidates) counts[c] = 0;
            foreach (var line in lines) foreach (var c in candidates) counts[c] += line.Count(ch => ch == c);
            return counts.OrderByDescending(x => x.Value).First().Key;
        }

        public static List<string[]> Parse(string path, Encoding enc, char delimiter)
        {
            var data = new List<string[]>();
            using (var reader = new StreamReader(path, enc))
            {
                while (!reader.EndOfStream)
                {
                    string? line = reader.ReadLine();
                    if (line == null) break;
                    data.Add(ParseLine(line, delimiter));
                }
            }
            return data;
        }

        private static string[] ParseLine(string line, char delimiter)
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
    }
}
