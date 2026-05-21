#!/bin/bash
# Löscht die Verzeichnisse bin und obj
rm -rf bin obj

# Erstellt das Windows-Binary (framework-dependent, schlank)
dotnet publish -c Release -r win-x64
