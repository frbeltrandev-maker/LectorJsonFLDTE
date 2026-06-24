import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from reports.anexos_consumidor_final import generar_anexo_consumidor_final


class AnexosConsumidorFinalTests(unittest.TestCase):
    def test_agrupa_facturas_por_fecha_y_tipo(self):
        with tempfile.TemporaryDirectory() as directorio:
            carpeta = Path(directorio)

            def crear_dte(tipo_dte, codigo, hora, gravada):
                return {
                    "identificacion": {
                        "fecEmi": "2026-06-21",
                        "horEmi": hora,
                        "tipoDte": tipo_dte,
                        "codigoGeneracion": codigo,
                    },
                    "receptor": {},
                    "resumen": {
                        "totalExenta": 1.00,
                        "totalNoSuj": 2.00,
                        "totalNoGravado": 0.00,
                        "totalGravada": gravada,
                        "montoTotalOperacion": gravada + 3.00,
                        "ivaRetenido": 0.00,
                    },
                }

            dtes = [
                crear_dte("01", "GEN-2", "12:00:00", 20.00),
                crear_dte("01", "GEN-1", "08:00:00", 10.00),
                crear_dte("11", "EXP-1", "09:00:00", 30.00),
            ]
            for indice, dte in enumerate(dtes, start=1):
                (carpeta / f"dte-{indice}.json").write_text(json.dumps(dte), encoding="utf-8")

            resumen = generar_anexo_consumidor_final(carpeta, carpeta, "consumidor")

            self.assertEqual(resumen[0], ("Facturas", 2))
            self.assertEqual(resumen[1], ("Facturas de exportacion", 1))
            self.assertEqual(resumen[-1], ("Documentos Procesados", 3))

            libro = load_workbook(carpeta / "consumidor.xlsx")
            filas = list(libro["Consumidor Final"].values)
            self.assertEqual(len(filas), 3)

            factura = filas[1]
            self.assertEqual(factura[2], "01")
            self.assertEqual(factura[7], "GEN-1")
            self.assertEqual(factura[8], "GEN-2")
            self.assertEqual(factura[10], 2.00)
            self.assertEqual(factura[12], 4.00)
            self.assertEqual(factura[13], 30.00)
            self.assertEqual(factura[19], 36.00)

            exportacion = filas[2]
            self.assertEqual(exportacion[2], "11")
            self.assertEqual(exportacion[7], "EXP-1")
            self.assertEqual(exportacion[14], 30.00)
