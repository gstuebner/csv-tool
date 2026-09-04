using System;
using System.Collections.Generic;
using System.Linq;

namespace CsvTool.Core
{
    public static class LineFilter
    {
        /// <summary>
        /// Parses a line filter specification (e.g. "1-10,20,50-") into a sorted list of 0-based row indices.
        /// Line numbers in the specification are 1-based.
        /// </summary>
        public static List<int> Parse(string spec, int totalRows)
        {
            var result = new List<int>();
            if (string.IsNullOrWhiteSpace(spec) || totalRows <= 0) return result;

            var parts = spec.Split(',', StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (trimmed.Contains('-'))
                {
                    // Check for open range like "50-"
                    if (trimmed.EndsWith('-'))
                    {
                        string startStr = trimmed.Substring(0, trimmed.Length - 1).Trim();
                        if (int.TryParse(startStr, out int start))
                        {
                            int s = Math.Clamp(start - 1, 0, totalRows - 1);
                            for (int i = s; i < totalRows; i++) result.Add(i);
                        }
                    }
                    // Check for open range like "-50"
                    else if (trimmed.StartsWith('-'))
                    {
                        string endStr = trimmed.Substring(1).Trim();
                        if (int.TryParse(endStr, out int end))
                        {
                            int e = Math.Clamp(end - 1, 0, totalRows - 1);
                            for (int i = 0; i <= e; i++) result.Add(i);
                        }
                    }
                    else
                    {
                        var rangeParts = trimmed.Split('-');
                        if (rangeParts.Length == 2 &&
                            int.TryParse(rangeParts[0].Trim(), out int start) &&
                            int.TryParse(rangeParts[1].Trim(), out int end))
                        {
                            int s = start - 1;
                            int e = end - 1;
                            if (s > e) (s, e) = (e, s);
                            if (s < 0) s = 0;
                            if (e >= totalRows) e = totalRows - 1;
                            for (int i = s; i <= e; i++) result.Add(i);
                        }
                    }
                }
                else if (int.TryParse(trimmed, out int line))
                {
                    int idx = line - 1;
                    if (idx >= 0 && idx < totalRows) result.Add(idx);
                }
            }

            return result.Distinct().OrderBy(x => x).ToList();
        }

        /// <summary>
        /// Filters rows according to the selected indices (Option A: Header row is always preserved at index 0).
        /// Returns the filtered rows.
        /// </summary>
        public static List<string[]> Apply(List<string[]> rows, List<int> indices)
        {
            if (rows == null || rows.Count == 0) return rows ?? new List<string[]>();
            if (indices == null || indices.Count == 0) return rows;

            var filtered = new List<string[]>();
            // Option A: Always keep header (row 0)
            filtered.Add(rows[0]);

            foreach (int idx in indices)
            {
                // If index 0 was requested, it's already added as header; do not duplicate as data row
                if (idx > 0 && idx < rows.Count)
                {
                    filtered.Add(rows[idx]);
                }
            }

            return filtered;
        }

        /// <summary>
        /// Computes the new 1-based source row numbers mapping for the filtered rows.
        /// Index 0 is the header (source line 1), followed by the source line numbers of each data row.
        /// </summary>
        public static int[] RemapSourceRowNumbers(int[]? oldSourceNumbers, List<int> indices, int totalRows)
        {
            var sourceNumbers = new List<int>();
            int headerNum = oldSourceNumbers != null && oldSourceNumbers.Length > 0 ? oldSourceNumbers[0] : 1;
            sourceNumbers.Add(headerNum);

            foreach (int idx in indices)
            {
                if (idx > 0 && idx < totalRows)
                {
                    int origNum = oldSourceNumbers != null && idx < oldSourceNumbers.Length
                        ? oldSourceNumbers[idx]
                        : idx + 1;
                    sourceNumbers.Add(origNum);
                }
            }

            return sourceNumbers.ToArray();
        }
    }
}
