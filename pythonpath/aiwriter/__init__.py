"""
Lógica de AI Writer independiente de UNO.

LibreOffice añade automáticamente la carpeta pythonpath/ de la extensión a
sys.path. Estos módulos solo usan la librería estándar, no importan uno y
mantienen la sintaxis compatible con Python 3.9 (el incluido en LibreOffice
24.2 para Windows), para poder probarlos con unittest fuera de LibreOffice.
"""
