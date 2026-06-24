import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from reports.anexos_sujeto_excluido import generar_anexo_sujeto_excluido


class AnexosSujetoExcluidoTests(unittest.TestCase):
    def test_genera_anexo_de_sujeto_excluido(self):
        with tempfile.TemporaryDirectory() as directorio:
            carpeta = Path(directorio)
            dte = {
                "identificacion": {
                    "fecEmi": "2026-06-21",
                    "tipoDte": "14",
                    "codigoGeneracion": "SE-001",
                },
                "sujetoExcluido": {
                    "numDocumento": "06141234567890",
                    "nombre": "Proveedor Excluido",
                },
                "resumen": {"totalCompra": 125.50},
                "selloRecibido": "SELLO-SE",
            }
            (carpeta / "sujeto.json").write_text(json.dumps(dte), encoding="utf-8")

            resumen = generar_anexo_sujeto_excluido(carpeta, carpeta, "sujeto")

            self.assertEqual(resumen[-2], ("Sujeto Excluido", 1))
            self.assertEqual(resumen[-1], ("Documentos Procesados", 1))
            libro = load_workbook(carpeta / "sujeto.xlsx")
            fila = list(libro["Sujeto Excluido"].values)[1]
            self.assertEqual(fila[0], 1)
            self.assertEqual(fila[3], "21/06/2026")
            self.assertEqual(fila[6], 125.50)
