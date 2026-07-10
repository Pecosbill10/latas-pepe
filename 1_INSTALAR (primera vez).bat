@echo off
title Instalando programa de Pepe...
echo.
echo  *** Instalando dependencias (solo se hace una vez) ***
echo.
cd /d "%~dp0"

set PYCMD=
where py >nul 2>nul && set PYCMD=py
if not defined PYCMD (
    python --version >nul 2>nul && set PYCMD=python
)
if not defined PYCMD (
    echo  *** No se encontro Python instalado en esta PC. ***
    echo  Instalalo desde https://www.python.org/downloads/ ^(tildando "Add python.exe to PATH"^) y volve a correr este archivo.
    pause
    exit /b 1
)

%PYCMD% -m pip install -r requirements.txt
echo.
echo  *** Instalacion completada! ***
echo  Ahora podes abrir el programa con "2_INICIAR PROGRAMA.bat"
echo.
pause
