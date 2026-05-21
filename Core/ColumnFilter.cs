using System;
using System.Collections.Generic;
using System.Linq;

namespace CsvTool.Core
{
    public static class ColumnFilter
    {
        public static List<int> Parse(string spec, int maxCols)
        {
            var result = new List<int>();
            if (string.IsNullOrWhiteSpace(spec) || maxCols <= 0) return result;

            var parts = spec.Split(',', StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (trimmed.Contains('-'))
                {
                    var rangeParts = trimmed.Split('-');
                    if (rangeParts.Length == 2 &&
                        int.TryParse(rangeParts[0].Trim(), out int start) &&
                        int.TryParse(rangeParts[1].Trim(), out int end))
                    {
                        int s = start - 1;
                        int e = end - 1;
                        if (s < 0) s = 0;
                        if (e >= maxCols) e = maxCols - 1;
                        for (int i = s; i <= e; i++) result.Add(i);
                    }
                }
                else if (int.TryParse(trimmed, out int col))
                {
                    int idx = col - 1;
                    if (idx >= 0 && idx < maxCols) result.Add(idx);
                }
            }

            return result.Distinct().OrderBy(x => x).ToList();
        }

        public static List<string[]> Apply(List<string[]> rows, List<int> indices)
        {
            if (indices == null || indices.Count == 0) return rows;
            var filtered = new List<string[]>(rows.Count);
            foreach (var row in rows)
            {
                var newRow = new string[indices.Count];
                for (int i = 0; i < indices.Count; i++)
                {
                    int srcIdx = indices[i];
                    newRow[i] = srcIdx < row.Length ? row[srcIdx] : "";
                }
                filtered.Add(newRow);
            }
            return filtered;
        }
    }
}
