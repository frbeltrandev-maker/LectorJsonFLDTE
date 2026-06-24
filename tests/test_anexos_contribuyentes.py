import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from reports.anexos_contribuyentes import generar_anexo_contribuyentes


class AnexosContribuyentesTests(unittest.TestCase):
    def test_genera_anexo_con_dte_03_05_y_06(self):
        with tempfile.TemporaryDirectory() as directorio:
            carpeta = Path(directorio)

            def crear_dte(tipo_dte, codigo, total_pagar, monto_operacion):
                return {
                    "identificacion": {
                        "fecEmi": "2026-06-21",
                        "tipoDte": tipo_dte,
                        "codigoGeneracion": codigo,
                        "numeroControl": f"DTE-{tipo_dte}-001",
                    },
                    "receptor": {"nrc": "123", "nombre": "Cliente Prueba"},
                    "resumen": {
                        "totalExenta": 10.00,
                        "descuExenta": 2.00,
                        "totalNoGravado": 4.00,
                        "descuNoSuj": 1.00,
                        "totalGravada": 20.00,
                        "descuGravada": 5.00,
                        "totalPagar": total_pagar,
                        "montoTotalOperacion": monto_operacion,
                        "tributos": [{"codigo": "20", "valor": 3.00}],
                    },
                    "selloRecibido": f"SELLO-{tipo_dte}",
                }

            dtes = [
                crear_dte("03", "GEN-03", 30.00, 31.00),
                crear_dte("05", "GEN-05", 40.00, 41.00),
                crear_dte("06", "GEN-06", 50.00, 51.00),
            ]
            for indice, dte in enumerate(dtes, start=1):
                (carpeta / f"dte-{indice}.json").write_text(json.dumps(dte), encoding="utf-8")

            resumen = generar_anexo_contribuyentes(carpeta, carpeta, "contribuyentes")

            self.assertEqual(
                resumen,
                [
                    ("Facturas", 0),
                    ("Notas de Crédito", 1),
                    ("Notas de Débito", 1),
                    ("Comprobantes de Crédito Fiscal", 1),
                    ("Comprobantes de Retención", 0),
                    ("Documentos Procesados", 3),
                ],
            )

            libro = load_workbook(carpeta / "contribuyentes.xlsx")
            filas = list(libro["Ventas a Contribuyentes"].values)
            self.assertEqual(len(filas), 4)
            self.assertEqual(filas[1][0], "21/06/2026")
            self.assertEqual(filas[1][2], "03")
            self.assertEqual(filas[1][11], 15.00)
            self.assertEqual(filas[1][15], 30.00)
            self.assertEqual(filas[2][2], "05")
            self.assertEqual(filas[2][15], 41.00)
            self.assertEqual(filas[3][2], "06")
            self.assertEqual(filas[3][15], 50.00)
