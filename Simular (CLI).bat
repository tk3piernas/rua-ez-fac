@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONUTF8=1
if not exist datos mkdir datos
".venv\Scripts\python.exe" -m appfacturas simular --entrada --salida "datos\ultima_simulacion.txt"
echo.
echo El informe quedo en datos\ultima_simulacion.txt. No se movio ni se escribio nada.
pause
