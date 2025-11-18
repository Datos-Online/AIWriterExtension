#!/bin/bash

# --- Configuración ---
EXTENSION_NAME="AIWriterExtension.oxt"
OUTPUT_DIR="dist"

# --- Archivos y carpetas a incluir en la extensión ---
FILES_TO_INCLUDE="AIWriterExtension.py description.xml Accelerators.xcu Addons.xcu LICENSE README.md ai-writer-icon.png lang META-INF"

# Verificar si zip está instalado
if ! command -v zip &> /dev/null
then
    echo "Error: El comando 'zip' no está instalado. Por favor, instálalo para continuar."
    exit 1
fi

echo "Creando el directorio de salida..."
mkdir -p "$OUTPUT_DIR"

echo "Eliminando la version anterior de la extension si existe..."
rm -f "$OUTPUT_DIR/$EXTENSION_NAME"

# Create the extension file
echo "Creando el archivo de la extension: $EXTENSION_NAME"
zip -r "$OUTPUT_DIR/$EXTENSION_NAME" $FILES_TO_INCLUDE -x "*.git*" "*__pycache__*"

echo
echo "Proceso completado. La extension se ha creado en: $OUTPUT_DIR/$EXTENSION_NAME"
