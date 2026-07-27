#!/bin/bash
# Löscht die Verzeichnisse bin und obj
rm -rf bin obj

# Erstellt das Windows-Binary (self-contained, mit Trimming) - läuft ohne installiertes .NET,
# ist dafür aber deutlich größer. Landet in bin/Release/net8.0/win-x64/standalone.
dotnet publish -c Release -r win-x64 \
  --self-contained=true \
  -p:PublishTrimmed=true \
  -p:TrimMode=partial \
  -o bin/Release/net8.0/win-x64/standalone
