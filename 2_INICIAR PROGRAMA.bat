@echo off
title Coleccion de Latas - Pepe
echo.
echo  *** Abriendo la coleccion de Pepe... ***
echo.
cd /d "%~dp0"

set PYCMD=
where py >nul 2>nul && set PYCMD=py
if not defined PYCMD (
    python --version >nul 2>nul && set PYCMD=python
)
if not defined PYCMD (
    echo  *** No se encontro Python instalado en esta PC. ***
    echo  Corre primero "1_INSTALAR (primera vez).bat".
    pause
    exit /b 1
)

start "" cmd /c "timeout /t 2 >nul && start http://localhost:5000"
%PYCMD% app.py
pause
