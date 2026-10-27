import unittest

from dte.duplicates import ControlDuplicadosDTE


class DuplicadosTests(unittest.TestCase):
    def test_normaliza_codigo_y_registra_archivo_original(self):
        control = ControlDuplicadosDTE()
        self.assertFalse(control.es_duplicado({"codigoGeneracion": "GEN-001"}, "original.json"))
        self.assertTrue(control.es_duplicado({"codigoGeneracion": " gen-001 "}, "copia.json"))
        self.assertEqual(control.avisos[0][0], "copia.json")
        self.assertIn("original.json", control.avisos[0][1])

    def test_no_confunde_documentos_sin_codigo_o_con_codigos_distintos(self):
        control = ControlDuplicadosDTE()
        for codigo in (None, None, "", "", "  ", "GEN-1", "GEN-2"):
            self.assertFalse(control.es_duplicado({"codigoGeneracion": codigo}, "dte.json"))
        self.assertEqual(control.avisos, [])

    def test_no_comparte_codigos_entre_generaciones(self):
        identificacion = {"codigoGeneracion": "GEN-001"}
        for _ in range(2):
            self.assertFalse(ControlDuplicadosDTE().es_duplicado(identificacion, "dte.json"))
