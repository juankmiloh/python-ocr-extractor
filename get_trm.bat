@echo off
echo "Vamos a obtener las TRM"
setlocal

REM Ruta relativa para lo demás (donde esté el .bat)
set "BASE_DIR=%~dp0"
set "VENV_DIR=%BASE_DIR%venv"
set "REQ_FILE=%BASE_DIR%requirements.txt"
set "SCRIPT=%BASE_DIR%main.py"

REM Ruta fija al Python embebido
set "PYTHON_DIR=%BASE_DIR%\python-3.13.9-embed-amd64\python.exe"

REM Imprimir la versión de Python embebido
echo Version de Python:
"%PYTHON_DIR%" --version

REM Crear entorno virtual si no existe
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo ----------------------------------------
    echo ** CREANDO ENTORNO VIRTUAL
    echo ----------------------------------------
    "%PYTHON_DIR%" -m virtualenv "%VENV_DIR%"
)

REM Activar entorno virtual
call "%VENV_DIR%\Scripts\activate.bat"

REM Instalar dependencias
if exist "%VENV_DIR%\Lib\site-packages\*" (
    echo ----------------------------------------
    echo ** LAS DEPENDENCIAS YA ESTAN INSTALADAS
    echo ----------------------------------------
) else (
    echo ----------------------------------------
    echo ** INSTALANDO DEPENDENCIAS DESDE requirements.txt
    echo ----------------------------------------
    pip install -r "%REQ_FILE%"
)

REM Ejecutar script principal
echo ----------------------------------------
echo ** EJECUTANDO OCR
echo ----------------------------------------
@REM python "%SCRIPT%" --ruta "C:\Users\JUANK\Desktop\TRM_PYTHON\PWC1"
set /p RUTA="NANA Por favor, ingresa la ruta de la carpeta a procesar: "
python "%SCRIPT%" --ruta "%RUTA%"

REM Final
echo.
echo PROCESO COMPLETADO
pause
endlocal