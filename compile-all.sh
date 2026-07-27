#!/bin/bash
# Baut alle Varianten nacheinander und sammelt die fertigen Binaries in dist/.
# Jedes Teilskript löscht bin/ und obj/, deshalb wird nach jedem Schritt kopiert.
set -e
cd "$(dirname "$0")"

rm -rf dist
mkdir -p dist

echo "=== Linux x64 (self-contained) ==="
./compile-for-linux.sh
cp bin/Release/net8.0/linux-x64/publish/csv dist/csv-linux-x64

echo "=== Windows x64 (framework-dependent, benötigt .NET 8 Runtime) ==="
./compile-for-windows.sh
cp bin/Release/net8.0/win-x64/publish/csv.exe dist/csv.exe

echo "=== Windows x64 (self-contained, ohne .NET-Installation lauffähig) ==="
./compile-for-windows-standalone.sh
cp bin/Release/net8.0/win-x64/standalone/csv.exe dist/csv-standalone.exe

echo
echo "Fertig. Ergebnisse in dist/:"
ls -lh dist
