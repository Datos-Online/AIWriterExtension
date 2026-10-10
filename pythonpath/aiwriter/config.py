"""
Lectura, escritura y validación de la configuración de AI Writer.

La configuración se guarda como JSON en aiwriter.json, dentro del directorio
UserConfig de LibreOffice. Resolver ese directorio requiere UNO, así que aquí
las funciones reciben la ruta completa del archivo.
"""
import json
import os

SETTINGS_FILE = "aiwriter.json"


def config_path(user_config_dir):
    """
    Devuelve la ruta del archivo de configuración.

    :param user_config_dir: El directorio UserConfig de LibreOffice (ruta del sistema).
    :return: La ruta completa de aiwriter.json.
    """
    return os.path.join(user_config_dir, SETTINGS_FILE)


def load_config(path):
    """
    Carga la configuración completa.

    :param path: La ruta de aiwriter.json.
    :return: Un diccionario; vacío si el archivo no existe, no se puede leer o
             no contiene un objeto JSON válido.
    """
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)
    except (IOError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def get_value(path, key, default):
    """
    Obtiene un valor de configuración.

    :param path: La ruta de aiwriter.json.
    :param key: La clave de configuración a buscar.
    :param default: El valor a devolver si la clave no se encuentra.
    :return: El valor de la configuración o el valor por defecto.
    """
    return load_config(path).get(key, default)


def set_value(path, key, value):
    """
    Guarda un valor de configuración conservando el resto de claves.

    Si el archivo no existe, lo crea.

    :param path: La ruta de aiwriter.json.
    :param key: La clave de configuración a establecer.
    :param value: El valor a guardar.
    """
    data = load_config(path)
    data[key] = value
    try:
        with open(path, "w", encoding="utf-8") as file:
            json.dump(data, file, indent=4)
    except IOError as e:
        print(f"Error writing to {path}: {e}")


def validate_settings(values):
    """
    Valida y convierte los valores introducidos en el diálogo de configuración.

    :param values: Diccionario con los textos de los campos del diálogo.
    :return: Una tupla (ajustes, clave_de_error). Si todo es válido, ajustes contiene
             los valores convertidos y clave_de_error es None; si no, ajustes es None
             y clave_de_error es la clave del mensaje en lang/*.json.
    """
    openai_api_key = values["openai_api_key"].strip()
    if not openai_api_key:
        return None, "no_api_key"

    model = values["model"].strip()
    if not model:
        return None, "invalid_model"

    try:
        max_tokens = int(values["max_tokens"].strip())
    except ValueError:
        max_tokens = 0
    if max_tokens <= 0:
        return None, "invalid_max_tokens"

    try:
        # Se acepta la coma como separador decimal
        temperature = float(values["temperature"].strip().replace(",", "."))
    except ValueError:
        temperature = None
    # La comparación también descarta nan e inf
    if temperature is None or not 0 <= temperature <= 2:
        return None, "invalid_temperature"

    return {
        "openai_api_key": openai_api_key,
        "model": model,
        "max_tokens": max_tokens,
        "temperature": temperature,
    }, None
