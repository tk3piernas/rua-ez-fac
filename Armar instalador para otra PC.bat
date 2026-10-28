@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Armar instalador de RUA EZ-Fac
".venv\Scripts\python.exe" -W ignore -m appfacturas empaquetar
pause
