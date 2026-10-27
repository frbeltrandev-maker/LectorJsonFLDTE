import json
import socket
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


URL_CONSULTA_DTE = "https://admin.factura.gob.sv/prod/consultas/publica/simple/1"
ESTADO_TRANSMITIDO_SATISFACTORIAMENTE = "Transmitido Satisfactoriamente"
MAX_CONSULTAS_POR_LOTE = 15
ESPERA_ENTRE_LOTES = 60
ESPERA_LIMITE_CONSULTAS = 60


class ConsultaHaciendaError(Exception):
    pass


class LimiteConsultasHaciendaError(ConsultaHaciendaError):
    pass


class ConsultorHaciendaPorLotes:
    def __init__(
        self,
        consultar,
        max_consultas=MAX_CONSULTAS_POR_LOTE,
        espera_entre_lotes=ESPERA_ENTRE_LOTES,
        esperar=time.sleep,
        al_esperar=None,
        al_reanudar=None,
    ):
        self.consultar_estado = consultar
        self.max_consultas = max_consultas
        self.espera_entre_lotes = espera_entre_lotes
        self.esperar = esperar
        self.al_esperar = al_esperar
        self.al_reanudar = al_reanudar
        self.consultas_realizadas = 0

    def consultar(self, codigo_generacion, fecha_emi):
        if (
            self.max_consultas > 0
            and self.consultas_realizadas > 0
            and self.consultas_realizadas % self.max_consultas == 0
        ):
            if self.al_esperar:
                self.al_esperar(
                    self.espera_entre_lotes,
                    f"Se completó un lote de {self.max_consultas} consultas.",
                )
            self.esperar(self.espera_entre_lotes)
            if self.al_reanudar:
                self.al_reanudar()

        self.consultas_realizadas += 1
        return self.consultar_estado(codigo_generacion, fecha_emi)


def _obtener_espera_limite(error, espera_predeterminada):
    retry_after = error.headers.get("Retry-After") if error.headers else None
    if retry_after:
        try:
            return max(0, int(retry_after))
        except (TypeError, ValueError):
            pass
    return espera_predeterminada


def consultar_estado_dte(
    codigo_generacion,
    fecha_emi,
    ambiente="01",
    timeout=20,
    reintentos=2,
    reintentos_limite=1,
    espera_limite=ESPERA_LIMITE_CONSULTAS,
    abrir_url=urlopen,
    esperar=time.sleep,
    al_esperar=None,
    al_reanudar=None,
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
    intentos_transitorios = 0
    intentos_limite = 0
    while True:
        try:
            with abrir_url(solicitud, timeout=timeout) as respuesta:
                return json.load(respuesta)
        except HTTPError as error:
            if error.code == 429:
                ultimo_error = error
                if intentos_limite < reintentos_limite:
                    intentos_limite += 1
                    segundos_espera = _obtener_espera_limite(error, espera_limite)
                    error.close()
                    if al_esperar:
                        al_esperar(
                            segundos_espera,
                            "Hacienda respondió HTTP 429 por exceso de consultas.",
                        )
                    esperar(segundos_espera)
                    if al_reanudar:
                        al_reanudar()
                    continue
                error.close()
                raise LimiteConsultasHaciendaError(
                    "Hacienda alcanzó el límite de consultas (HTTP 429)"
                ) from error
            error.close()
            raise ConsultaHaciendaError(str(error)) from error
        except json.JSONDecodeError as error:
            raise ConsultaHaciendaError(f"Respuesta JSON inválida: {error}") from error
        except (URLError, TimeoutError, socket.timeout) as error:
            ultimo_error = error
            if intentos_transitorios < reintentos:
                intentos_transitorios += 1
                esperar(intentos_transitorios)
                continue
            break

    raise ConsultaHaciendaError(
        f"No hubo respuesta de Hacienda después de {intentos_transitorios + 1} intentos: {ultimo_error}"
    ) from ultimo_error


def dte_esta_transmitido_satisfactoriamente(respuesta):
    return respuesta.get("estadoDoc") == ESTADO_TRANSMITIDO_SATISFACTORIAMENTE
