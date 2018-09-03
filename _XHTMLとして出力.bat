@rem -*- mode: bat; coding: cp932-dos -*-
@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -Command "./asql -FilePath '%~1' -ExportToDirectoryPath '%~dp1' -ExportXHTML -XHTML_AnchorTarget '_blank' -XHTML_CSSFilePath '.\style.css' -XHTML_JavaScriptFilePath '.\jquery-1.4.2.min.js','.\numeric.js' -Encoding UTF8 -ShowExports -SkipErrors"
exit /b 0
