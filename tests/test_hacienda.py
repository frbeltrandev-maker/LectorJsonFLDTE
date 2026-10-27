import io
import unittest
from urllib.error import HTTPError

from services.hacienda import (
    ConsultorHaciendaPorLotes,
    LimiteConsultasHaciendaError,
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

    def test_espera_despues_de_cada_lote_de_quince_consultas(self):
        consultas = []
        esperas = []
        avisos = []
        reanudaciones = []
        consultor = ConsultorHaciendaPorLotes(
            consultar=lambda codigo, fecha: consultas.append((codigo, fecha)) or {},
            max_consultas=15,
            espera_entre_lotes=60,
            esperar=esperas.append,
            al_esperar=lambda segundos, motivo: avisos.append((segundos, motivo)),
            al_reanudar=lambda: reanudaciones.append(True),
        )

        for numero in range(16):
            consultor.consultar(f"CODIGO-{numero}", "2026-09-07")

        self.assertEqual(len(consultas), 16)
        self.assertEqual(esperas, [60])
        self.assertEqual(avisos[0][0], 60)
        self.assertIn("lote de 15", avisos[0][1])
        self.assertEqual(reanudaciones, [True])

    def test_reintenta_una_vez_cuando_hacienda_responde_429(self):
        intentos = []
        esperas = []
        avisos = []
        reanudaciones = []

        class RespuestaFalsa(io.BytesIO):
            def __enter__(self):
                return self

            def __exit__(self, tipo, valor, traceback):
                self.close()

        def abrir_url_falsa(solicitud, timeout):
            intentos.append((solicitud.full_url, timeout))
            if len(intentos) == 1:
                raise HTTPError(solicitud.full_url, 429, "Too Many Requests", {}, None)
            return RespuestaFalsa(b'{"estadoDoc": "Transmitido Satisfactoriamente"}')

        respuesta = consultar_estado_dte(
            "CODIGO-PRUEBA",
            "2026-09-07",
            reintentos_limite=1,
            espera_limite=60,
            abrir_url=abrir_url_falsa,
            esperar=esperas.append,
            al_esperar=lambda segundos, motivo: avisos.append((segundos, motivo)),
            al_reanudar=lambda: reanudaciones.append(True),
        )

        self.assertEqual(len(intentos), 2)
        self.assertEqual(esperas, [60])
        self.assertEqual(avisos[0][0], 60)
        self.assertIn("HTTP 429", avisos[0][1])
        self.assertEqual(reanudaciones, [True])
        self.assertTrue(dte_esta_transmitido_satisfactoriamente(respuesta))

    def test_informa_limite_si_hacienda_mantiene_el_429(self):
        intentos = []
        esperas = []

        def abrir_url_falsa(solicitud, timeout):
            intentos.append((solicitud.full_url, timeout))
            raise HTTPError(solicitud.full_url, 429, "Too Many Requests", {}, None)

        with self.assertRaises(LimiteConsultasHaciendaError):
            consultar_estado_dte(
                "CODIGO-PRUEBA",
                "2026-09-07",
                reintentos_limite=1,
                espera_limite=60,
                abrir_url=abrir_url_falsa,
                esperar=esperas.append,
            )

        self.assertEqual(len(intentos), 2)
        self.assertEqual(esperas, [60])
