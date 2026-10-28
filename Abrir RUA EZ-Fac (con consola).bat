@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONUTF8=1
title RUA EZ-Fac (con consola)
echo Abriendo RUA EZ-Fac ... Esta ventana muestra lo que hace la app; si la cerras, la app se cierra.
".venv\Scripts\python.exe" -m appfacturas ui
pause
