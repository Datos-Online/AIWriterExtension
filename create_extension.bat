@echo off
setlocal

REM --- Configuración ---
SET "EXTENSION_NAME=AIWriterExtension"
SET "OUTPUT_DIR=dist"

REM --- Archivos y carpetas a incluir en la extensión ---
SET FILES_TO_INCLUDE=AIWriterExtension.py description.xml Addons.xcu Accelerators.xcu LICENSE README.md ai-writer-icon.png lang META-INF

echo Creando el directorio de salida...
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

echo Eliminando la version anterior de la extension si existe...
if exist "%OUTPUT_DIR%\%EXTENSION_NAME%.zip" del "%OUTPUT_DIR%\%EXTENSION_NAME%.zip"
if exist "%OUTPUT_DIR%\%EXTENSION_NAME%.oxt" del "%OUTPUT_DIR%\%EXTENSION_NAME%.oxt"

echo Creando el archivo de la extension: %EXTENSION_NAME%

tar -a -c -f "%OUTPUT_DIR%\%EXTENSION_NAME%.zip" %FILES_TO_INCLUDE%
move "%OUTPUT_DIR%\%EXTENSION_NAME%.zip" "%OUTPUT_DIR%\%EXTENSION_NAME%.oxt"

echo.
echo Proceso completado. La extension se ha creado en: %OUTPUT_DIR%\%EXTENSION_NAME%.oxt

endlocal