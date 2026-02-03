#!/bin/bash
# Löscht die Verzeichnisse bin und obj
rm -rf bin obj

# Erstellt das Linux-Binary
dotnet publish -c Release -r linux-x64 --self-contained=true -p:PublishSingleFile=true -p:PublishTrimmed=true
