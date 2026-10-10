import unittest

from aiwriter import models


class TestModelCapabilities(unittest.TestCase):

    def test_gpt5_no_admite_temperatura(self):
        for model in ("gpt-5", "gpt-5-mini", "gpt-5-nano", "gpt-5.1"):
            with self.subTest(model=model):
                self.assertFalse(models.supports(model, "temperature"))

    def test_gpt6_no_admite_temperatura(self):
        for model in ("gpt-6-astra", "gpt-6-sol", "gpt-6.1-sol", "gpt-6-luna"):
            with self.subTest(model=model):
                self.assertFalse(models.supports(model, "temperature"))

    def test_no_distingue_mayusculas_ni_espacios(self):
        self.assertFalse(models.supports("GPT-5-Mini", "temperature"))
        self.assertFalse(models.supports("  gpt-5  ", "temperature"))

    def test_otros_modelos_admiten_temperatura(self):
        for model in ("gpt-4o-mini", "gpt-4o", "gpt-4.1"):
            with self.subTest(model=model):
                self.assertTrue(models.supports(model, "temperature"))

    def test_modelo_desconocido_o_vacio_usa_valores_por_defecto(self):
        for model in ("modelo-inventado", "", None):
            with self.subTest(model=model):
                self.assertEqual(models.capabilities(model), models.DEFAULT_CAPABILITIES)

    def test_capacidad_desconocida_es_false(self):
        self.assertFalse(models.supports("gpt-4o-mini", "capacidad-inexistente"))

    def test_capabilities_devuelve_una_copia(self):
        models.capabilities("gpt-4o-mini")["temperature"] = False
        self.assertTrue(models.DEFAULT_CAPABILITIES["temperature"])


if __name__ == "__main__":
    unittest.main()
