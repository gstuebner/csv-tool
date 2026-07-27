#!/bin/bash
# Löscht die Verzeichnisse bin und obj
rm -rf bin obj

# Erstellt das Linux-Binary (self-contained, mit Trimming)
dotnet publish -c Release -r linux-x64
