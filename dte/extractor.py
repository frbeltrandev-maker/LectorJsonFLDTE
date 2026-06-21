from datetime import datetime


def obtener_contenido_dte(datos, campo_participante):
    campos_requeridos = ("identificacion", campo_participante, "resumen")

    if all(campo in datos for campo in campos_requeridos):
        return datos

    for valor in datos.values():
        if isinstance(valor, dict) and all(campo in valor for campo in campos_requeridos):
            return valor
    return None


def obtener_fecha_formateada_dte(fec_emi):
    if not fec_emi:
        return None, ""

    fecha = datetime.strptime(fec_emi, "%Y-%m-%d")
    return fecha, fecha.strftime("%d/%m/%Y")


def clasificar_documento_sujeto_excluido(numero_documento):
    if not isinstance(numero_documento, str):
        return None

    longitud = len(numero_documento)
    if 9 <= longitud <= 10:
        return 2
    if 14 <= longitud <= 17:
        return 1
    if longitud > 17:
        return 3
    return None


def obtener_sello_recibido(datos):
    campos_posibles = [
        "SelloRecepcion",
        "selloRecepcion",
        "response",
        "responseMH",
        "ResponseMH",
        "respuestaHacienda",
    ]
    for campo in campos_posibles:
        if campo in datos:
            valor = datos[campo]
            if isinstance(valor, dict):
                return valor.get("selloRecibido")
            if isinstance(valor, str):
                return valor

    valor = datos.get("selloRecibido")
    if isinstance(valor, str):
        return valor
    return None
