rd /s /q bin obj
REM Erstellt das Windows-Binary (framework-dependent, schlank) - wie compile-for-windows.sh
dotnet publish -c Release -r win-x64
