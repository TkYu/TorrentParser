@echo off
cd /d %~dp0
rem srm uninstall TorrentParserShell.dll
taskkill /f /im explorer.exe
start explorer.exe
start explorer.exe %~dp0