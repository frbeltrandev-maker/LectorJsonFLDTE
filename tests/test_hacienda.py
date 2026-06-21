import io
import unittest

from services.hacienda import (
    consultar_estado_dte,
    dte_esta_transmitido_satisfactoriamente,
)


class HaciendaTests(unittest.TestCase):
    def test_identifica_dte_transmitido_satisfactoriamente(self):
        respuesta = {
            "estadoDoc": "Transmitido Satisfactoriamente",
            "estadoDocInc": "SR",
            "accion": "OK",
        }

        self.assertTrue(dte_esta_transmitido_satisfactoriamente(respuesta))

    def test_rechaza_dte_invalidado(self):
        respuesta = {
            "estadoDoc": "Invalidado",
            "estadoDocInc": "INV",
            "accion": "OK",
        }

        self.assertFalse(dte_esta_transmitido_satisfactoriamente(respuesta))

    def test_no_usa_estado_doc_inc_para_filtrar(self):
        respuesta = {
            "estadoDoc": "Transmitido Satisfactoriamente",
            "estadoDocInc": "INV",
        }

        self.assertTrue(dte_esta_transmitido_satisfactoriamente(respuesta))

    def test_reintenta_consulta_despues_de_timeout(self):
        intentos = []

        class RespuestaFalsa(io.BytesIO):
            def __enter__(self):
                return self

            def __exit__(self, tipo, valor, traceback):
                self.close()

        def abrir_url_falsa(solicitud, timeout):
            intentos.append((solicitud.full_url, timeout))
            if len(intentos) < 3:
                raise TimeoutError("The read operation timed out")
            return RespuestaFalsa(b'{"estadoDoc": "Transmitido Satisfactoriamente"}')

        respuesta = consultar_estado_dte(
            "CODIGO-PRUEBA",
            "2026-06-21",
            timeout=5,
            reintentos=2,
            abrir_url=abrir_url_falsa,
            esperar=lambda segundos: None,
        )

        self.assertEqual(len(intentos), 3)
        self.assertEqual(respuesta["estadoDoc"], "Transmitido Satisfactoriamente")
        self.assertIn("codigoGeneracion=CODIGO-PRUEBA", intentos[0][0])
