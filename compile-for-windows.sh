#!/bin/bash
# Löscht die Verzeichnisse bin und obj (falls vorhanden)
rm -rf bin obj

# Erstellt das Windows-Binary
dotnet publish -c Release -r win-x64 --self-contained=true -p:PublishSingleFile=true -p:PublishTrimmed=true
