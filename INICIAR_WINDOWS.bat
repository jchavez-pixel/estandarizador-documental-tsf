@echo off
setlocal EnableExtensions
cd /d "%~dp0"
echo ============================================================
echo ESTANDARIZADOR DOCUMENTAL TSF - INSTALADOR CORREGIDO
echo ============================================================
echo.

echo Buscando Python compatible (3.12 o 3.13)...
set "PY_CMD="
set "VENV_DIR="

where py >nul 2>nul
if not errorlevel 1 (
  py -3.12 --version >nul 2>nul
  if not errorlevel 1 (
    set "PY_CMD=py -3.12"
    set "VENV_DIR=.venv312"
    goto interpreter_found
  )
  py -3.13 --version >nul 2>nul
  if not errorlevel 1 (
    set "PY_CMD=py -3.13"
    set "VENV_DIR=.venv313"
    goto interpreter_found
  )
)

where python >nul 2>nul
if not errorlevel 1 (
  python -c "import sys; raise SystemExit(0 if sys.version_info[:2] in ((3,12),(3,13)) else 1)" >nul 2>nul
  if not errorlevel 1 (
    set "PY_CMD=python"
    set "VENV_DIR=.venv_supported"
    goto interpreter_found
  )
)

echo.
echo ERROR: No se encontro Python 3.12 o 3.13.
echo Tu instalacion parece estar usando Python 3.15, que no es compatible
 echo con las ruedas binarias de algunas dependencias de esta version.
echo.
echo Instala Python 3.12 de 64 bits desde:
echo https://www.python.org/downloads/windows/
echo Durante la instalacion marca "Add python.exe to PATH" e instala el Python Launcher.
echo Luego cierra esta ventana, abre una nueva y vuelve a ejecutar este archivo.
echo.
pause
exit /b 1

:interpreter_found
echo Interprete seleccionado: %PY_CMD%
echo Entorno virtual: %VENV_DIR%

if not exist "%VENV_DIR%\Scripts\python.exe" (
  echo Creando entorno virtual...
  %PY_CMD% -m venv "%VENV_DIR%"
  if errorlevel 1 goto venv_error
)

set "VENV_PY=%CD%\%VENV_DIR%\Scripts\python.exe"
echo.
echo Actualizando herramientas de instalacion...
"%VENV_PY%" -m pip install --upgrade pip setuptools wheel
if errorlevel 1 goto install_error

echo.
echo Instalando dependencias usando paquetes binarios precompilados...
echo No se intentara compilar C/C++ en este equipo.
"%VENV_PY%" -m pip install --only-binary=:all: -r requirements.txt
if errorlevel 1 goto install_error

echo.
echo Comprobando motor de exportacion PDF...
"%VENV_PY%" -c "from app import find_pdf_export_engine; print(find_pdf_export_engine() or 'MOTOR_PDF_NO_DETECTADO')"
"%VENV_PY%" -c "from app import find_pdf_export_engine; raise SystemExit(0 if find_pdf_export_engine() else 1)"
if errorlevel 1 (
  echo AVISO: No se detecto Microsoft Word COM ni LibreOffice.
  echo Se requiere Word de escritorio con COM habilitado o LibreOffice ya instalado.
)

echo.
echo Iniciando Estandarizador Documental TSF...
"%VENV_PY%" app.py
pause
exit /b 0

:venv_error
echo ERROR: No se pudo crear el entorno virtual.
pause
exit /b 1

:install_error
echo.
echo ERROR: No se pudieron instalar las dependencias.
echo Verifica conexion a internet y confirma que Python 3.12/3.13 sea de 64 bits.
echo Esta version evita compilar dependencias nativas; no instales Visual Studio para resolverlo.
pause
exit /b 1
