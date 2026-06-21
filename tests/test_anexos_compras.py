import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from reports.anexos_compras import generar_anexo_compras


class AnexosComprasTests(unittest.TestCase):
    def test_genera_anexo_y_resumen_de_compras(self):
        with tempfile.TemporaryDirectory() as directorio:
            carpeta = Path(directorio)
            dte = {
                "identificacion": {
                    "fecEmi": "2026-06-21",
                    "tipoDte": "03",
                    "codigoGeneracion": "GEN-001",
                },
                "emisor": {
                    "nit": "0614-000000-000-0",
                    "nrc": "12345",
                    "nombre": "Proveedor de Prueba",
                },
                "resumen": {
                    "totalExenta": 10.00,
                    "descuExenta": 2.00,
                    "totalNoSuj": 4.00,
                    "descuNoSuj": 1.00,
                    "totalGravada": 20.00,
                    "descuGravada": 5.00,
                    "montoTotalOperacion": 30.00,
                    "ivaPerci1": 0.30,
                    "tributos": [
                        {"codigo": "D1", "valor": 1.00},
                        {"codigo": "C8", "valor": 2.00},
                        {"codigo": "20", "valor": 3.00},
                    ],
                },
                "selloRecibido": "SELLO-001",
            }
            (carpeta / "dte.json").write_text(json.dumps(dte), encoding="utf-8")
            (carpeta / "nota.txt").write_text("archivo auxiliar", encoding="utf-8")

            progresos = []
            resumen = generar_anexo_compras(
                carpeta,
                carpeta,
                "compras",
                consultar_estado=lambda codigo, fecha: {
                    "estadoDoc": "Transmitido Satisfactoriamente",
                },
                progreso=lambda actual, total, archivo: progresos.append((actual, total, archivo)),
            )

            self.assertEqual(
                resumen,
                [
                    ("Facturas", 0),
                    ("Notas de Crédito", 0),
                    ("Notas de Débito", 0),
                    ("Comprobantes de Crédito Fiscal", 1),
                    ("Documentos Procesados", 1),
                ],
            )

            libro = load_workbook(carpeta / "compras.xlsx")
            self.assertEqual(libro.sheetnames, ["Resumen de DTEs", "Percepción", "Errores"])

            fila_resumen = list(libro["Resumen de DTEs"].values)[1]
            self.assertEqual(fila_resumen[0], "21/06/2026")
            self.assertEqual(fila_resumen[2], "03")
            self.assertEqual(fila_resumen[6], 11.00)
            self.assertEqual(fila_resumen[9], 15.00)
            self.assertEqual(fila_resumen[13], 3.00)

            fila_percepcion = list(libro["Percepción"].values)[1]
            self.assertEqual(fila_percepcion[3], "SELLO-001")
            self.assertEqual(fila_percepcion[6], 0.30)

            libro_errores = load_workbook(carpeta / "compras_errores.xlsx")
            self.assertEqual(list(libro_errores["Errores"].values)[1], ("nota.txt", "No es un archivo JSON"))
            self.assertEqual(len(progresos), 2)
            self.assertTrue(all(total == 2 for _, total, _ in progresos))

    def test_excluye_dte_invalidado_por_hacienda(self):
        with tempfile.TemporaryDirectory() as directorio:
            carpeta = Path(directorio)
            dte = {
                "identificacion": {
                    "fecEmi": "2026-05-26",
                    "tipoDte": "03",
                    "codigoGeneracion": "DTE-INVALIDADO",
                },
                "emisor": {"nit": "0614-000000-000-0", "nrc": "1", "nombre": "Proveedor"},
                "resumen": {"montoTotalOperacion": 10.00},
            }
            (carpeta / "invalidado.json").write_text(json.dumps(dte), encoding="utf-8")

            resumen = generar_anexo_compras(
                carpeta,
                carpeta,
                "compras",
                consultar_estado=lambda codigo, fecha: {
                    "estadoDoc": "Invalidado",
                    "descripcionEstado": "Documento Invalidado por el Emisor.",
                },
            )

            self.assertEqual(resumen[-1], ("Documentos Procesados", 0))
            libro = load_workbook(carpeta / "compras.xlsx")
            self.assertEqual(libro["Resumen de DTEs"].max_row, 1)

            self.assertFalse((carpeta / "compras_errores.xlsx").exists())
            libro_invalidos = load_workbook(carpeta / "InvalidosDTE.xlsx")
            self.assertEqual(
                list(libro_invalidos["DTE Invalidos"].values)[1],
                (
                    "invalidado.json",
                    "DTE-INVALIDADO",
                    "26/05/2026",
                    "Invalidado",
                    "Documento Invalidado por el Emisor.",
                ),
            )
