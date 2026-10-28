@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Instalar RUA EZ-Fac
set "PY=.venv\Scripts\python.exe"
set "PYW=.venv\Scripts\pythonw.exe"
set "PIPQ=--disable-pip-version-check -q"
if not exist datos mkdir datos
echo ==================================================
echo   RUA EZ-Fac - instalar o actualizar
echo ==================================================
echo   Carpeta: %~dp0
echo.
echo %~dp0 | findstr /i /c:"OneDrive" >nul
if not errorlevel 1 goto :enonedrive
echo %~dp0 | findstr /i /c:"Para subir a GitHub" >nul
if not errorlevel 1 goto :espejo

echo [1/5] Python ...
if exist "%PY%" goto :tengo_venv
set "BASEPY="
py -3.15 -c "pass" >nul 2>&1 && set "BASEPY=py -3.15"
if defined BASEPY goto :crear_venv
rem Without the "py" launcher there is no Python Install Manager: open the download page.
where py >nul 2>&1
if errorlevel 1 goto :sinpython
echo       No tengo Python 3.15; lo instalo con el Python Install Manager ...
py install 3.15
py -3.15 -c "pass" >nul 2>&1 && set "BASEPY=py -3.15"
if not defined BASEPY goto :sinpython
:crear_venv
echo       Creando el entorno de la app ...
%BASEPY% -m venv .venv
if errorlevel 1 goto :sinpython
:tengo_venv
"%PY%" -c "import sys;print('      Python', sys.version.split()[0])"

echo [2/5] Librerias (la primera vez tarda unos minutos) ...
"%PY%" -m pip install %PIPQ% --upgrade pip
"%PY%" -m pip install %PIPQ% -r requirements.txt -r requirements-ui.txt
if errorlevel 1 goto :offline
"%PY%" -m pip install %PIPQ% --no-deps -r requirements-nodeps.txt
if not errorlevel 1 goto :ruta
:offline
if not exist wheels goto :falla
echo       Sin internet: instalo desde la carpeta wheels ...
"%PY%" -m pip install %PIPQ% --no-index --find-links wheels -r requirements.txt -r requirements-ui.txt
if errorlevel 1 goto :falla
"%PY%" -m pip install %PIPQ% --no-index --find-links wheels --no-deps -r requirements-nodeps.txt
if errorlevel 1 goto :falla

:ruta
echo [3/5] Conectando la app ...
"%PY%" -c "import sysconfig,pathlib;pathlib.Path(sysconfig.get_paths()['purelib'],'appfacturas.pth').write_text(r'%~dp0app')"

echo [4/5] Iconos en el Escritorio y en el menu Inicio ...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$s = New-Object -ComObject WScript.Shell; foreach ($c in @([Environment]::GetFolderPath('Programs'), [Environment]::GetFolderPath('Desktop'))) { try { $v = Join-Path $c 'App Facturas.lnk'; if (Test-Path $v) { Remove-Item $v }; $l = $s.CreateShortcut((Join-Path $c 'RUA EZ-Fac.lnk')); $l.TargetPath = '%~dp0.venv\Scripts\pythonw.exe'; $l.Arguments = '-m appfacturas ui'; $l.WorkingDirectory = '%~dp0'; $l.Description = 'RUA EZ-Fac: carga de facturas de proveedores'; $l.IconLocation = '%~dp0app\appfacturas\static\ezfac.ico'; $l.Save(); Write-Host ('      ' + $c + '\RUA EZ-Fac') } catch { Write-Host ('      No pude crear el icono en ' + $c) } }"

echo [5/5] Revisando que este todo ...
if exist app\tests (
  "%PY%" -m pytest -q app\tests > datos\ultimo_test.txt 2>&1
  if errorlevel 1 echo       Algunas pruebas fallaron: detalle en datos\ultimo_test.txt
)
if exist app\tests_privados (
  "%PY%" -m pytest -q app\tests_privados >> datos\ultimo_test.txt 2>&1
  if errorlevel 1 echo       Algunas pruebas privadas fallaron: detalle en datos\ultimo_test.txt
)
echo.
"%PY%" -W ignore -m appfacturas verificar
set "RC=%errorlevel%"
echo.
echo (Esto quedo guardado en datos\verificacion.txt)
echo.
if not "%RC%"=="0" goto :fin_con_faltas
choice /c SN /n /m "Listo. Abro RUA EZ-Fac ahora? [S/N] "
if errorlevel 2 exit /b 0
start "" "%PYW%" -m appfacturas ui
exit /b 0

:fin_con_faltas
echo La app quedo instalada, pero hay cosas que faltan (FALTA arriba).
pause
exit /b 1

:espejo
echo Esta carpeta es solo para subir los archivos a GitHub: no se instala desde aca.
echo Usa el instalador de la carpeta de la app (Documentos\App Facturas).
pause
exit /b 1

:enonedrive
echo Esta carpeta esta dentro de OneDrive y el OneDrive de RUA no deja guardar programas.
echo Pasa la carpeta a Documentos (por ejemplo C:\Users\%USERNAME%\Documents\RUA EZ-Fac) y volve a abrir este archivo.
pause
exit /b 1

:sinpython
echo No pude conseguir Python 3.15.
echo Instalalo desde la pagina que se abre (Python install manager, sin permisos de administrador)
echo y despues volve a abrir este archivo.
start "" "https://www.python.org/downloads/windows/"
pause
exit /b 1

:falla
echo No pude instalar las librerias (falta internet o algo bloqueo la descarga).
echo Pedile a Matheo la carpeta wheels y ponela al lado de este archivo.
pause
exit /b 1
