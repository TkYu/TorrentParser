@echo off
cd /d %~dp0
rem srm uninstall TorrentParserShell.dll
FOR /F "tokens=3 delims= " %%G in ('reg query "hklm\system\controlset001\control\nls\language" /v Installlanguage') DO (
IF [%%G] EQU [0409] (
  goto enUS
)
IF [%%G] EQU [0404] (
  goto zhTW
)
goto zhCN
)

:zhCN
echo 按任意键重启explorer以使其生效，如不需要请直接关闭
goto restart

:zhTW
echo 按任意鍵重啟explorer以使其生效，如不需要請直接關閉
goto restart

:enUS
echo Press any key to restart explorer.
goto restart

:restart
pause>nul
taskkill /f /im explorer.exe
start explorer.exe
start explorer.exe %~dp0