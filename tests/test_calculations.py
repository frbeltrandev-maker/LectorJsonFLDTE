import unittest

from dte.calculations import (
    calcular_valores_consumidor_final,
    calcular_valores_resumen,
    iterar_tributos_conocidos,
    obtener_total_compra_sujeto_excluido,
)


class CalculationsTests(unittest.TestCase):
    def setUp(self):
        self.resumen = {
            "totalExenta": 10.00,
            "descuExenta": 2.00,
            "totalNoSuj": 4.00,
            "totalNoGravado": 7.00,
            "descuNoSuj": 1.00,
            "totalGravada": 20.00,
            "descuGravada": 5.00,
            "montoTotalOperacion": 30.00,
            "totalPagar": 31.00,
            "ivaPerci1": 0.30,
            "ivaRete1": 0.50,
            "ivaRetenido": 0.75,
            "totalCompra": 99.00,
            "tributos": [
                {"codigo": "D1", "valor": 1.00},
                {"codigo": "C8", "valor": 2.00},
                {"codigo": "20", "valor": 3.00},
                {"codigo": "OTRO", "valor": 4.00},
            ],
        }

    def test_calcula_resumen_con_total_gravada(self):
        valores = calcular_valores_resumen(
            self.resumen,
            campo_no_sujeta="totalNoSuj",
            campo_monto_total="montoTotalOperacion",
        )

        self.assertEqual(valores["total_exenta"], 8.00)
        self.assertEqual(valores["total_no_sujeta"], 3.00)
        self.assertEqual(valores["total_gravada"], 15.00)
        self.assertEqual(valores["monto_total_operacion"], 30.00)
        self.assertEqual(valores["iva_perci1"], 0.30)

    def test_calcula_resumen_con_clave_gravada_y_retencion(self):
        valores = calcular_valores_resumen(
            self.resumen,
            campo_no_sujeta="totalNoGravado",
            campo_monto_total="totalPagar",
            clave_gravada="sub_total_ventas",
            incluir_retencion=True,
        )

        self.assertEqual(valores["total_no_sujeta"], 6.00)
        self.assertEqual(valores["sub_total_ventas"], 15.00)
        self.assertEqual(valores["monto_total_operacion"], 31.00)
        self.assertEqual(valores["iva_ret1"], 0.50)

    def test_extrae_solo_tributos_conocidos(self):
        tributos = list(iterar_tributos_conocidos(self.resumen))

        self.assertEqual(tributos, [("fovial", 1.00), ("cotrans", 2.00), ("iva", 3.00)])

    def test_calcula_consumidor_final_y_sujeto_excluido(self):
        valores = calcular_valores_consumidor_final(
            self.resumen,
            campo_iva="ivaRetenido",
            clave_iva="iva",
        )

        self.assertEqual(valores["total_exenta"], 8.00)
        self.assertEqual(valores["total_no_sujeta"], 3.00)
        self.assertEqual(valores["total_no_gravado"], 7.00)
        self.assertEqual(valores["total_gravada"], 15.00)
        self.assertEqual(valores["iva"], 0.75)
        self.assertEqual(obtener_total_compra_sujeto_excluido(self.resumen), 99.00)
