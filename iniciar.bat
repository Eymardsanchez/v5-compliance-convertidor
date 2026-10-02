@echo off
setlocal enabledelayedexpansion

REM Verificar si Python está instalado
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Python detectado. Procediendo...
    goto ejecutar
)

REM Python no encontrado, descargar e instalar
echo Python no encontrado. Descargando e instalando la Ultima version...

set "PYTHON_INSTALLER=%TEMP%\python-installer.exe"
set "PYTHON_VERSION=3.13.7"
set "DOWNLOAD_URL=https://www.python.org/ftp/python/!PYTHON_VERSION!/python-!PYTHON_VERSION!-amd64.exe"

echo Descargando Python !PYTHON_VERSION!...
powershell -Command "Invoke-WebRequest -Uri '!DOWNLOAD_URL!' -OutFile '!PYTHON_INSTALLER!'" 2>nul

if not exist "!PYTHON_INSTALLER!" (
    echo Error al descargar Python.
    pause
    exit /b 1
)

echo Instalando Python...
"!PYTHON_INSTALLER!" /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1

REM Esperar a que termine la instalación
timeout /t 15 >nul

REM Verificar si la instalación fue exitosa
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: No fue posible instalar Python.
    pause
    exit /b 1
)

echo Python instalado correctamente.

:ejecutar
echo.
echo Instalando dependencias...
python -m pip install -r requirements.txt --quiet

if %errorlevel% neq 0 (
    echo Error al instalar las dependencias.
    pause
    exit /b 1
)

echo.
echo Ejecutando ConvertidorSantander...
echo.
python ConvertidorSantander.py

pause
exit /b 0
@echo off
setlocal enabledelayedexpansion

REM Verificar si Python está instalado
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Python detectado. Procediendo...
    goto ejecutar
)

REM Python no encontrado, descargar e instalar
echo Python no encontrado. Descargando e instalando la última versión...

set "PYTHON_INSTALLER=%TEMP%\python-installer.exe"
set "PYTHON_VERSION=3.13.7"
set "DOWNLOAD_URL=https://www.python.org/ftp/python/!PYTHON_VERSION!/python-!PYTHON_VERSION!-amd64.exe"

echo Descargando Python !PYTHON_VERSION!...
powershell -Command "Invoke-WebRequest -Uri '!DOWNLOAD_URL!' -OutFile '!PYTHON_INSTALLER!'" 2>nul

if not exist "!PYTHON_INSTALLER!" (
    echo Error al descargar Python.
    pause
    exit /b 1
)

echo Instalando Python...
"!PYTHON_INSTALLER!" /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1

REM Esperar a que termine la instalación
timeout /t 15 >nul

REM Verificar si la instalación fue exitosa
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: No fue posible instalar Python.
    pause
    exit /b 1
)

echo Python instalado correctamente.

:ejecutar
echo.
echo Instalando dependencias...
python -m pip install -r requirements.txt --quiet

if %errorlevel% neq 0 (
    echo Error al instalar las dependencias.
    pause
    exit /b 1
)

echo.
echo Ejecutando ConvertidorSantander...
echo.
python ConvertidorSantander.py

pause
exit /b 0
