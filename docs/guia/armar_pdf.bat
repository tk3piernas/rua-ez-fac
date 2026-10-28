@echo off
rem Regenera "docs\Guia de uso RUA EZ-Fac.pdf" desde el HTML de esta carpeta, con Edge (viene con Windows).
cd /d "%~dp0"
set "EDGE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE%" set "EDGE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
"%EDGE%" --headless --disable-gpu --no-pdf-header-footer --print-to-pdf="%~dp0..\Guia de uso RUA EZ-Fac.pdf" "%~dp0Guia de uso RUA EZ-Fac.html"
echo Listo: docs\Guia de uso RUA EZ-Fac.pdf
