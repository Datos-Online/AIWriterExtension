@echo off
setlocal

REM --- Configuración ---
SET EXTENSION_NAME=AIWriterExtension.oxt
SET OUTPUT_DIR=dist

REM --- Archivos y carpetas a incluir en la extensión ---
SET FILES_TO_INCLUDE=^
    AIWriterExtension.py ^
    description.xml ^
    Addons.xcu ^
    Accelerators.xcu ^
    LICENSE ^
    README.md ^
    ai-writer-icon.png ^
    lang ^
    META-INF

REM Verificar si zip está instalado
where zip >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo Error: El comando 'zip' no esta instalado o no se encuentra en el PATH.
    echo Por favor, instala una utilidad zip (como la de 7-Zip) y asegurate de que este en el PATH del sistema.
    goto :eof
)

echo Creando el directorio de salida...
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

echo Eliminando la version anterior de la extension si existe...
if exist "%OUTPUT_DIR%\%EXTENSION_NAME%" del "%OUTPUT_DIR%\%EXTENSION_NAME%"

echo Creando el archivo de la extension: %EXTENSION_NAME%
zip -r "%OUTPUT_DIR%\%EXTENSION_NAME%" %FILES_TO_INCLUDE% -x "*.git*" "*__pycache__*"

echo.
echo Proceso completado. La extension se ha creado en: %OUTPUT_DIR%\%EXTENSION_NAME%
endlocal
