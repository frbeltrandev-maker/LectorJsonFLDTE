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
