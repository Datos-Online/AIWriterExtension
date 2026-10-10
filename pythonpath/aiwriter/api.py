"""
Comunicación con la API Chat Completions de OpenAI.

Construye el prompt y el payload, envía la petición e interpreta la respuesta.
Los errores se lanzan como ApiError con una clave de los archivos lang/*.json,
de modo que este módulo no depende del idioma de la interfaz.
"""
import json
import socket
import urllib.error
import urllib.request

from aiwriter import models

API_URL = "https://api.openai.com/v1/chat/completions"
API_TIMEOUT = 90  # Segundos máximos de espera de la respuesta de la API

# Clave del prompt de cada comando en los archivos de idioma
PROMPT_KEYS = {
    "complete": "prompt_complete",
    "summarize": "prompt_summarize",
    "improve": "prompt_improve",
    "expand": "prompt_expand",
    "translate": "prompt_translate",
}


class ApiError(Exception):
    """
    Error al llamar a la API, identificado por una clave de los archivos de idioma.

    :param key: Clave del mensaje (por ejemplo, "error_timeout").
    :param params: Datos para completar el mensaje (seconds, reason, code, detail).
    """

    def __init__(self, key, **params):
        super().__init__(key)
        self.key = key
        self.params = params

    def message(self, lang):
        """
        Construye el mensaje para el usuario en el idioma indicado.

        :param lang: Diccionario de textos cargado de lang/<idioma>.json.
        :return: El mensaje a mostrar en el diálogo de error.
        """
        text = lang[self.key]
        if self.key == "error_timeout":
            return text.format(seconds=self.params["seconds"])
        if self.key == "error_connection":
            return f"{text}: {self.params['reason']}"
        if self.key == "error_api":
            return f"{text} (HTTP {self.params['code']}): {self.params['detail']}"
        return text


def build_prompt(lang, command, text, target_language=""):
    """
    Construye el mensaje del usuario para un comando.

    :param lang: Diccionario de textos del idioma de la interfaz.
    :param command: El comando de IA (ej. 'complete', 'translate').
    :param text: El texto seleccionado.
    :param target_language: El idioma de destino, solo para 'translate'.
    :return: El prompt, o None si el comando no es válido.
    """
    if command not in PROMPT_KEYS:
        return None
    instruction = lang[PROMPT_KEYS[command]]
    if command == "translate":
        return f"{instruction} {target_language}: {text}"
    return f"{instruction}: {text}"


def build_payload(model, system_prompt, user_prompt, max_tokens, temperature):
    """
    Construye el cuerpo de la petición a Chat Completions.

    Usa max_completion_tokens (no max_tokens) y omite temperature en los modelos
    que no la admiten, según la tabla de aiwriter.models.

    :param model: El nombre del modelo.
    :param system_prompt: Las instrucciones generales del asistente.
    :param user_prompt: El prompt del comando con el texto.
    :param max_tokens: Máximo de tokens de la respuesta (int o texto numérico).
    :param temperature: Temperatura (float o texto numérico).
    :return: El payload como diccionario.
    """
    payload = {
        "model": model,
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": system_prompt
                    }
                ]
            }, {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": user_prompt
                    }
                ]
            }
        ],
        "max_completion_tokens": int(max_tokens),
        "temperature": float(temperature)
    }

    if not models.supports(model, "temperature"):
        del payload["temperature"]

    return payload


def parse_response(response_data):
    """
    Extrae el texto generado de la respuesta de la API.

    El contenido puede venir nulo o vacío; por ejemplo, en los modelos de
    razonamiento los tokens internos pueden agotar max_completion_tokens y la
    respuesta llega vacía con finish_reason "length".

    :param response_data: El JSON de respuesta ya decodificado.
    :return: El texto generado, sin espacios al principio ni al final.
    :raises ApiError: Si la respuesta no trae contenido.
    """
    try:
        choice = response_data["choices"][0]
    except (KeyError, IndexError, TypeError):
        raise ApiError("error_empty_response")

    content = (choice.get("message") or {}).get("content") or ""
    content = content.strip()
    if content:
        return content

    if choice.get("finish_reason") == "length":
        raise ApiError("error_token_limit")
    raise ApiError("error_empty_response")


def http_error_detail(body, reason):
    """
    Obtiene el detalle de un error HTTP de la API.

    OpenAI devuelve el detalle en el cuerpo JSON (`error.message`); si el cuerpo
    no es JSON o no tiene ese campo, se usa el texto recibido tal cual.

    :param body: El cuerpo de la respuesta de error, ya decodificado.
    :param reason: El texto de estado HTTP, usado si el cuerpo está vacío.
    :return: El detalle a mostrar.
    """
    detail = ""
    try:
        detail = json.loads(body).get("error", {}).get("message", "")
    except (ValueError, AttributeError):
        pass
    if not detail:
        detail = body.strip() or str(reason)
    return detail


def send(payload, api_key, url=API_URL, timeout=API_TIMEOUT):
    """
    Envía el payload a la API y devuelve el texto generado.

    :param payload: El cuerpo de la petición (ver build_payload).
    :param api_key: La clave de la API de OpenAI.
    :param url: El endpoint de Chat Completions.
    :param timeout: Segundos máximos de espera.
    :return: El texto generado.
    :raises ApiError: Por timeout, error de conexión, error HTTP o respuesta vacía.
    """
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")

    try:
        # urlopen lanza HTTPError para cualquier estado distinto de 2xx
        with urllib.request.urlopen(req, timeout=timeout) as response:
            response_data = json.loads(response.read().decode("utf-8"))
    except (socket.timeout, TimeoutError):
        # socket.timeout no es subclase de TimeoutError en Python < 3.10
        raise ApiError("error_timeout", seconds=timeout)
    except urllib.error.HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="replace")
        except Exception:
            body = ""
        raise ApiError("error_api", code=e.code, detail=http_error_detail(body, e.reason))
    except urllib.error.URLError as e:
        if isinstance(e.reason, (socket.timeout, TimeoutError)):
            raise ApiError("error_timeout", seconds=timeout)
        raise ApiError("error_connection", reason=e.reason)

    return parse_response(response_data)
