# Implementation Plan - Column Filtering on Export (`-c` / `--columns`)

This plan outlines the implementation of a column selection filter when writing outputs via the `-o` / `--output` parameter. It will allow users to specify which columns to export using ranges and indices (e.g., `-c 1-3,6`).

## Proposed CLI Design

- **New Option**: `-c` or `--columns <RANGE>`
- **Example Usage**: `csv -c 1-3,6 -o output.csv input.xlsx`
- **Behavior**:
  - Only works in combination with `-o` / `--output`. Using `-c` without `-o` will print a helpful error message.
  - Column indices are **1-based** (e.g. `1` means the first column).
  - Supports comma-separated single columns (e.g., `1,3,5`) and dash-separated ranges (e.g., `1-3`).
  - Preserves the order of columns specified by the user (e.g., `-c 3,1` outputs Column 3 first, then Column 1).
  - If a row in the input has fewer columns than the requested index, an empty cell (`""`) will be written in its place.

## Open Questions

> [!NOTE]
> **Column Reordering**: The proposed design allows reordering columns (e.g., `-c 3,1` writes Column 3 and then Column 1). Do you prefer this behavior, or should the columns always be written in their original relative order (e.g., Column 1 then Column 3) regardless of how they are listed in the `-c` argument?

## Proposed Changes

---

### CLI Parsing and Validation

#### [MODIFY] [Program.cs](file:///home/gregor/Dokumente/cs/csv-tool/Program.cs)
- Add `-c` / `--columns` argument parsing to the main argument loop.
- Validate that if `-c` is passed, `-o` must also be specified.
- Implement the parsing logic in a helper method `ParseColumns(string columnsStr)`:
  - Splits by `,`.
  - Parses single indices and ranges (e.g., `start-end`).
  - Converts 1-based user input to 0-based indices.
  - Throws an `ArgumentException` with a clear explanation if the input format is invalid.
- Pass the parsed list of 0-based column indices into the `SaveAsCsv` and `SaveAsExcel` calls.

---

### CSV & Excel Export Filtration

#### [MODIFY] [CsvViewer.cs](file:///home/gregor/Dokumente/cs/csv-tool/CsvViewer.cs)
- Update `SaveAsCsv(string path, List<int>? columnFilter = null)` to write only the filtered columns.
- Update `SaveAsExcel(string path, List<int>? columnFilter = null)` to construct a filtered dataset before calling `ExcelHandler.SaveAsExcel`.

---

## Verification Plan

### Automated Tests
- We will write a set of unit tests or verify manually via console commands.
- Command to compile:
  ```bash
  bash compile-for-linux.sh
  ```

### Manual Verification
- Verify invalid format error:
  ```bash
  ./bin/Release/net8.0/linux-x64/publish/csv -c abc -o test_out.csv input.csv
  ```
- Verify usage without `-o` error:
  ```bash
  ./bin/Release/net8.0/linux-x64/publish/csv -c 1-3 input.csv
  ```
- Verify correct column filtering (e.g., columns 1-3 and 6):
  ```bash
  ./bin/Release/net8.0/linux-x64/publish/csv -c 1-3,6 -o test_out.csv input.csv
  ```
- Verify reordering or out-of-bounds handling (e.g., column index larger than actual columns).
