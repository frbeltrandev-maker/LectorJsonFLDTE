import unittest

from dte.extractor import (
    clasificar_documento_sujeto_excluido,
    obtener_contenido_dte,
    obtener_fecha_formateada_dte,
    obtener_sello_recibido,
)


class ExtractorTests(unittest.TestCase):
    def test_obtiene_dte_en_raiz(self):
        datos = {"identificacion": {}, "emisor": {}, "resumen": {}}

        self.assertIs(obtener_contenido_dte(datos, "emisor"), datos)

    def test_obtiene_dte_anidado(self):
        contenido = {"identificacion": {}, "receptor": {}, "resumen": {}}
        datos = {"respuesta": contenido}

        self.assertIs(obtener_contenido_dte(datos, "receptor"), contenido)

    def test_retorna_none_cuando_no_hay_dte(self):
        self.assertIsNone(obtener_contenido_dte({"otro": {}}, "emisor"))

    def test_formatea_fecha_y_maneja_fecha_vacia(self):
        self.assertEqual(obtener_fecha_formateada_dte("2026-06-21")[1], "21/06/2026")
        self.assertEqual(obtener_fecha_formateada_dte(None), (None, ""))

    def test_clasifica_documento_sujeto_excluido(self):
        self.assertEqual(clasificar_documento_sujeto_excluido("123456789"), 2)
        self.assertEqual(clasificar_documento_sujeto_excluido("12345678901234"), 1)
        self.assertEqual(clasificar_documento_sujeto_excluido("123456789012345678"), 3)
        self.assertIsNone(clasificar_documento_sujeto_excluido("123"))
        self.assertIsNone(clasificar_documento_sujeto_excluido(None))

    def test_obtiene_sello_recibido_en_diferentes_campos(self):
        self.assertEqual(obtener_sello_recibido({"selloRecibido": "SELLO-1"}), "SELLO-1")
        self.assertEqual(
            obtener_sello_recibido({"responseMH": {"selloRecibido": "SELLO-2"}}),
            "SELLO-2",
        )
