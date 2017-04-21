@echo off
taskkill /f /im explorer.exe
start explorer.exe
start explorer.exe %~dp0