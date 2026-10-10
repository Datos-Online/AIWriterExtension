import json
import os
import shutil
import tempfile
import unittest

from aiwriter import config


class TestConfigFile(unittest.TestCase):

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.path = config.config_path(self.dir)

    def tearDown(self):
        shutil.rmtree(self.dir)

    def write(self, text):
        with open(self.path, "w", encoding="utf-8") as f:
            f.write(text)

    def read_json(self):
        with open(self.path, encoding="utf-8") as f:
            return json.load(f)

    def test_nombre_del_archivo(self):
        self.assertEqual(self.path, os.path.join(self.dir, "aiwriter.json"))

    def test_archivo_inexistente_devuelve_el_valor_por_defecto(self):
        self.assertEqual(config.load_config(self.path), {})
        self.assertEqual(config.get_value(self.path, "model", "gpt-4o-mini"), "gpt-4o-mini")

    def test_json_corrupto_devuelve_el_valor_por_defecto(self):
        for text in ("{corrupto", "", "[1, 2]", "null"):
            with self.subTest(text=text):
                self.write(text)
                self.assertEqual(config.get_value(self.path, "model", "def"), "def")

    def test_lee_valores_guardados(self):
        self.write('{"model": "gpt-5", "max_tokens": 500}')
        self.assertEqual(config.get_value(self.path, "model", "x"), "gpt-5")
        self.assertEqual(config.get_value(self.path, "max_tokens", 1000), 500)
        self.assertEqual(config.get_value(self.path, "temperature", "0.5"), "0.5")

    def test_set_value_crea_el_archivo(self):
        config.set_value(self.path, "language", "francés")
        self.assertEqual(self.read_json(), {"language": "francés"})

    def test_set_value_conserva_las_demas_claves(self):
        config.set_value(self.path, "openai_api_key", "sk-test")
        config.set_value(self.path, "model", "gpt-4o")
        config.set_value(self.path, "model", "gpt-5")
        self.assertEqual(self.read_json(), {"openai_api_key": "sk-test", "model": "gpt-5"})

    def test_set_value_conserva_los_tipos(self):
        config.set_value(self.path, "max_tokens", 1000)
        config.set_value(self.path, "temperature", 0.7)
        data = self.read_json()
        self.assertIs(type(data["max_tokens"]), int)
        self.assertIs(type(data["temperature"]), float)

    def test_set_value_sobre_json_corrupto_empieza_de_cero(self):
        self.write("{corrupto")
        config.set_value(self.path, "model", "gpt-4o")
        self.assertEqual(self.read_json(), {"model": "gpt-4o"})

    def test_formato_del_archivo(self):
        # Se mantiene el formato de siempre: sangría de 4 espacios y solo ASCII
        config.set_value(self.path, "language", "español")
        with open(self.path, encoding="utf-8") as f:
            text = f.read()
        self.assertEqual(text, '{\n    "language": "espa\\u00f1ol"\n}')


class TestValidateSettings(unittest.TestCase):

    BASE = {"openai_api_key": " sk-x ", "model": " gpt-4o-mini ", "max_tokens": "1000", "temperature": "0.5"}

    def validate(self, **changes):
        values = dict(self.BASE)
        values.update(changes)
        return config.validate_settings(values)

    def test_valores_validos_se_limpian_y_convierten(self):
        self.assertEqual(self.validate(), ({
            "openai_api_key": "sk-x",
            "model": "gpt-4o-mini",
            "max_tokens": 1000,
            "temperature": 0.5,
        }, None))

    def test_coma_decimal(self):
        settings, error = self.validate(temperature="0,7")
        self.assertIsNone(error)
        self.assertEqual(settings["temperature"], 0.7)

    def test_limites_de_temperatura(self):
        for value in ("0", "2", "1.5"):
            with self.subTest(value=value):
                self.assertIsNone(self.validate(temperature=value)[1])

    def test_temperatura_invalida(self):
        for value in ("abc", "3", "-1", "2.01", "", "nan", "inf"):
            with self.subTest(value=value):
                self.assertEqual(self.validate(temperature=value), (None, "invalid_temperature"))

    def test_max_tokens_con_espacios(self):
        self.assertEqual(self.validate(max_tokens=" 500 ")[0]["max_tokens"], 500)

    def test_max_tokens_invalido(self):
        for value in ("abc", "-1", "0", "", "1.5", "1e3"):
            with self.subTest(value=value):
                self.assertEqual(self.validate(max_tokens=value), (None, "invalid_max_tokens"))

    def test_modelo_vacio(self):
        self.assertEqual(self.validate(model="   "), (None, "invalid_model"))

    def test_clave_vacia(self):
        self.assertEqual(self.validate(openai_api_key=""), (None, "no_api_key"))

    def test_claves_de_error_existen_en_ambos_idiomas(self):
        lang_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "lang")
        for code in ("es", "en"):
            with open(os.path.join(lang_dir, f"{code}.json"), encoding="utf-8") as f:
                lang = json.load(f)
            for key in ("no_api_key", "invalid_model", "invalid_max_tokens", "invalid_temperature"):
                with self.subTest(lang=code, key=key):
                    self.assertIn(key, lang)


if __name__ == "__main__":
    unittest.main()
