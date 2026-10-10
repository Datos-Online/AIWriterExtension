import io
import json
import os
import socket
import unittest
import urllib.error
from unittest import mock

from aiwriter import api

LANG_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "lang")


def load_lang(code):
    with open(os.path.join(LANG_DIR, f"{code}.json"), encoding="utf-8") as f:
        return json.load(f)


def fake_response(body):
    """Simula el objeto que devuelve urlopen usado como gestor de contexto."""
    response = mock.MagicMock()
    response.read.return_value = body
    context = mock.MagicMock()
    context.__enter__.return_value = response
    return context


def http_error(code, reason, body):
    return urllib.error.HTTPError(api.API_URL, code, reason, {}, io.BytesIO(body))


class TestBuildPrompt(unittest.TestCase):

    def setUp(self):
        self.lang = load_lang("es")

    def test_comandos(self):
        for command, key in api.PROMPT_KEYS.items():
            if command == "translate":
                continue
            with self.subTest(command=command):
                self.assertEqual(api.build_prompt(self.lang, command, "hola"), f"{self.lang[key]}: hola")

    def test_traducir_incluye_el_idioma(self):
        self.assertEqual(
            api.build_prompt(self.lang, "translate", "hola", "francés"),
            f"{self.lang['prompt_translate']} francés: hola",
        )

    def test_comando_desconocido(self):
        self.assertIsNone(api.build_prompt(self.lang, "hello", "hola"))

    def test_todos_los_prompts_existen_en_ambos_idiomas(self):
        for code in ("es", "en"):
            lang = load_lang(code)
            for key in list(api.PROMPT_KEYS.values()) + ["prompt_assistant"]:
                with self.subTest(lang=code, key=key):
                    self.assertTrue(lang[key])


class TestBuildPayload(unittest.TestCase):

    def test_estructura(self):
        payload = api.build_payload("gpt-4o-mini", "sistema", "usuario", 700, 0.3)
        self.assertEqual(payload, {
            "model": "gpt-4o-mini",
            "messages": [
                {"role": "assistant", "content": [{"type": "text", "text": "sistema"}]},
                {"role": "user", "content": [{"type": "text", "text": "usuario"}]},
            ],
            "max_completion_tokens": 700,
            "temperature": 0.3,
        })
        self.assertNotIn("max_tokens", payload)

    def test_convierte_valores_guardados_como_texto(self):
        # Configuraciones antiguas guardaban max_tokens como texto
        payload = api.build_payload("gpt-4o-mini", "s", "u", "1000", "0.5")
        self.assertIs(type(payload["max_completion_tokens"]), int)
        self.assertIs(type(payload["temperature"]), float)

    def test_omite_temperatura_si_el_modelo_no_la_admite(self):
        for model in ("gpt-5-mini", "gpt-6-sol"):
            with self.subTest(model=model):
                self.assertNotIn("temperature", api.build_payload(model, "s", "u", 100, 0.5))


class TestParseResponse(unittest.TestCase):

    def assertApiError(self, data, key):
        with self.assertRaises(api.ApiError) as ctx:
            api.parse_response(data)
        self.assertEqual(ctx.exception.key, key)

    def test_contenido_normal(self):
        data = {"choices": [{"message": {"content": "  Hola  "}, "finish_reason": "stop"}]}
        self.assertEqual(api.parse_response(data), "Hola")

    def test_contenido_nulo_por_limite_de_tokens(self):
        data = {"choices": [{"message": {"content": None}, "finish_reason": "length"}]}
        self.assertApiError(data, "error_token_limit")

    def test_respuestas_vacias(self):
        for data in (
            {"choices": [{"message": {"content": ""}, "finish_reason": "stop"}]},
            {"choices": [{"message": {"content": "   "}}]},
            {"choices": [{"message": None}]},
            {"choices": []},
            {},
            None,
        ):
            with self.subTest(data=data):
                self.assertApiError(data, "error_empty_response")


class TestHttpErrorDetail(unittest.TestCase):

    def test_mensaje_json_de_openai(self):
        body = '{"error": {"message": "Incorrect API key provided"}}'
        self.assertEqual(api.http_error_detail(body, "Unauthorized"), "Incorrect API key provided")

    def test_cuerpo_no_json(self):
        self.assertEqual(api.http_error_detail(" <html>bad gateway</html> ", "Bad Gateway"), "<html>bad gateway</html>")

    def test_json_que_no_es_un_objeto(self):
        self.assertEqual(api.http_error_detail('["x"]', "Bad Request"), '["x"]')

    def test_cuerpo_vacio_usa_reason(self):
        self.assertEqual(api.http_error_detail("", "Too Many Requests"), "Too Many Requests")


class TestApiErrorMessage(unittest.TestCase):

    def test_mensajes_en_ambos_idiomas(self):
        errors = [
            api.ApiError("error_timeout", seconds=90),
            api.ApiError("error_connection", reason="sin red"),
            api.ApiError("error_api", code=401, detail="Incorrect API key"),
            api.ApiError("error_token_limit"),
            api.ApiError("error_empty_response"),
        ]
        for code in ("es", "en"):
            lang = load_lang(code)
            for error in errors:
                with self.subTest(lang=code, key=error.key):
                    self.assertTrue(error.message(lang).startswith(lang[error.key].split("{")[0]))

    def test_formatos(self):
        lang = load_lang("es")
        self.assertIn("90 segundos", api.ApiError("error_timeout", seconds=90).message(lang))
        self.assertEqual(
            api.ApiError("error_connection", reason="sin red").message(lang),
            f"{lang['error_connection']}: sin red",
        )
        self.assertEqual(
            api.ApiError("error_api", code=401, detail="Incorrect API key").message(lang),
            f"{lang['error_api']} (HTTP 401): Incorrect API key",
        )


class TestSend(unittest.TestCase):

    PAYLOAD = {"model": "gpt-4o-mini", "messages": []}

    def send_with(self, effect):
        with mock.patch("urllib.request.urlopen") as urlopen:
            if isinstance(effect, BaseException):
                urlopen.side_effect = effect
            else:
                urlopen.return_value = fake_response(effect)
            result = api.send(self.PAYLOAD, "sk-test")
        return result, urlopen

    def assertSendError(self, effect, key):
        with self.assertRaises(api.ApiError) as ctx:
            self.send_with(effect)
        self.assertEqual(ctx.exception.key, key)
        return ctx.exception

    def test_peticion_correcta(self):
        body = b'{"choices": [{"message": {"content": "Resultado"}}]}'
        result, urlopen = self.send_with(body)
        self.assertEqual(result, "Resultado")

        request = urlopen.call_args.args[0]
        self.assertEqual(request.full_url, api.API_URL)
        self.assertEqual(request.get_method(), "POST")
        self.assertEqual(request.get_header("Authorization"), "Bearer sk-test")
        self.assertEqual(request.get_header("Content-type"), "application/json")
        self.assertEqual(json.loads(request.data), self.PAYLOAD)
        self.assertEqual(urlopen.call_args.kwargs["timeout"], api.API_TIMEOUT)

    def test_timeout_por_defecto(self):
        self.assertEqual(api.API_TIMEOUT, 90)

    def test_error_http(self):
        error = self.assertSendError(
            http_error(401, "Unauthorized", b'{"error": {"message": "Incorrect API key"}}'), "error_api")
        self.assertEqual(error.params, {"code": 401, "detail": "Incorrect API key"})

    def test_timeout_directo(self):
        for effect in (socket.timeout("timed out"), TimeoutError("timed out")):
            with self.subTest(effect=type(effect).__name__):
                error = self.assertSendError(effect, "error_timeout")
                self.assertEqual(error.params["seconds"], api.API_TIMEOUT)

    def test_timeout_dentro_de_urlerror(self):
        self.assertSendError(urllib.error.URLError(socket.timeout("timed out")), "error_timeout")

    def test_error_de_conexion(self):
        error = self.assertSendError(urllib.error.URLError("Name or service not known"), "error_connection")
        self.assertEqual(error.params["reason"], "Name or service not known")

    def test_respuesta_vacia(self):
        self.assertSendError(b'{"choices": []}', "error_empty_response")

    def test_json_invalido_se_propaga(self):
        # La clase UNO muestra cualquier otra excepción tal cual en un diálogo
        with self.assertRaises(ValueError):
            self.send_with(b"<html>")


if __name__ == "__main__":
    unittest.main()
