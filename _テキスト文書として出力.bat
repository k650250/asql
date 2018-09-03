@rem -*- mode: bat; coding: cp932-dos -*-
@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -Command "./asql -FilePath '%~1' -ExportToDirectoryPath '%~dp1' -ExportPlainText -ShowExports -SkipErrors"
exit /b 0
