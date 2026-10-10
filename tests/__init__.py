"""
Pruebas de la lógica de AI Writer que no depende de LibreOffice.

Ejecutar desde la raíz del repositorio (así se importa este paquete y se
configura la ruta antes de cargar las pruebas):
    python3 -m unittest -v
"""
import os
import sys

# Los módulos de la extensión viven en pythonpath/, igual que dentro del .oxt
PYTHONPATH_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "pythonpath")
if PYTHONPATH_DIR not in sys.path:
    sys.path.insert(0, PYTHONPATH_DIR)
