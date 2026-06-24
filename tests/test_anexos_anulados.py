import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from reports.anexos_anulados import generar_anexo_anulados


class AnexosAnuladosTests(unittest.TestCase):
    def test_genera_anexo_de_documentos_anulados(self):
        with tempfile.TemporaryDirectory() as directorio:
            carpeta = Path(directorio)
            dte = {
                "identificacion": {
                    "tipoDte": "03",
                    "numeroControl": "DTE-03-001",
                    "codigoGeneracion": "ANU-001",
                },
                "emisor": {},
                "resumen": {},
                "selloRecibido": "SELLO-ANU",
            }
            (carpeta / "anulado.json").write_text(json.dumps(dte), encoding="utf-8")

            resumen = generar_anexo_anulados(carpeta, carpeta, "anulados")

            self.assertEqual(resumen[1], ("Comprobantes de Crédito Fiscal", 1))
            self.assertEqual(resumen[-1], ("Documentos Procesados", 1))
            libro = load_workbook(carpeta / "anulados.xlsx")
            fila = list(libro["Documentos Anulados"].values)[1]
            self.assertEqual(fila[0], "DTE-03-001")
            self.assertEqual(fila[4], "03")
            self.assertEqual(fila[6], "SELLO-ANU")
            self.assertEqual(fila[9], "ANU-001")
