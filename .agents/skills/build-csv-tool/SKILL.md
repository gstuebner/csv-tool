---
name: build-csv-tool
description: Instructions and guidelines for building and packaging csv-tool using the project compile scripts.
---

# Build CSV Tool

This skill documents how to build and package `csv-tool` properly across platforms.

## Important Rule
Do **not** use `dotnet build` or raw `dotnet publish` directly for distribution or release artifacts. Always use the provided build scripts in the repository root:

* **Linux x64 (self-contained, trimmed):**
  `./compile-for-linux.sh`
  * Cleans `bin/` and `obj/`.
  * Generates self-contained, trimmed Linux executable in `bin/Release/net8.0/linux-x64/publish/csv`.

* **Windows x64 (framework-dependent):**
  `./compile-for-windows.sh` (or `compile-for-windows.cmd` on Windows)
  * Requires .NET 8 Runtime on the target machine.
  * Output: `bin/Release/net8.0/win-x64/publish/csv.exe`.

* **Windows x64 (standalone, self-contained):**
  `./compile-for-windows-standalone.sh`
  * Runs on Windows without a .NET installation.
  * Output: `bin/Release/net8.0/win-x64/standalone/csv.exe`.

* **All platforms (Distribution bundle):**
  `./compile-all.sh`
  * Runs all the above scripts sequentially and copies the final executables to `dist/`:
    * `dist/csv-linux-x64`
    * `dist/csv.exe`
    * `dist/csv-standalone.exe`

## Quick verification during development
For fast verification while testing code changes on Linux, run:
```bash
./compile-for-linux.sh
```
and test the resulting binary at:
```bash
bin/Release/net8.0/linux-x64/publish/csv [args...]
```
