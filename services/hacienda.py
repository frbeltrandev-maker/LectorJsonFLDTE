import json
import socket
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


URL_CONSULTA_DTE = "https://admin.factura.gob.sv/prod/consultas/publica/simple/1"
ESTADO_TRANSMITIDO_SATISFACTORIAMENTE = "Transmitido Satisfactoriamente"


class ConsultaHaciendaError(Exception):
    pass


def consultar_estado_dte(
    codigo_generacion,
    fecha_emi,
    ambiente="01",
    timeout=20,
    reintentos=2,
    abrir_url=urlopen,
    esperar=time.sleep,
):
    parametros = urlencode({
        "codigoGeneracion": codigo_generacion,
        "fechaEmi": fecha_emi,
        "ambiente": ambiente,
    })
    solicitud = Request(
        f"{URL_CONSULTA_DTE}?{parametros}",
        headers={"Accept": "application/json", "User-Agent": "LJSON-V3/1.0"},
    )

    ultimo_error = None
    for intento in range(reintentos + 1):
        try:
            with abrir_url(solicitud, timeout=timeout) as respuesta:
                return json.load(respuesta)
        except HTTPError as error:
            raise ConsultaHaciendaError(str(error)) from error
        except json.JSONDecodeError as error:
            raise ConsultaHaciendaError(f"Respuesta JSON inválida: {error}") from error
        except (URLError, TimeoutError, socket.timeout) as error:
            ultimo_error = error
            if intento < reintentos:
                esperar(intento + 1)

    raise ConsultaHaciendaError(
        f"No hubo respuesta de Hacienda después de {reintentos + 1} intentos: {ultimo_error}"
    ) from ultimo_error


def dte_esta_transmitido_satisfactoriamente(respuesta):
    return respuesta.get("estadoDoc") == ESTADO_TRANSMITIDO_SATISFACTORIAMENTE
