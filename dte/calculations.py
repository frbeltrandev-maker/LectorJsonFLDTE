TRIBUTOS_CONOCIDOS = {
    "D1": "fovial",
    "C8": "cotrans",
    "20": "iva",
}


def calcular_valores_resumen(
    resumen,
    *,
    campo_no_sujeta,
    campo_monto_total,
    clave_gravada="total_gravada",
    incluir_retencion=False,
):
    valores = {
        "total_exenta": resumen.get("totalExenta", 0.00) - resumen.get("descuExenta", 0.00),
        "total_no_sujeta": resumen.get(campo_no_sujeta, 0.00) - resumen.get("descuNoSuj", 0.00),
        clave_gravada: resumen.get("totalGravada", 0.00) - resumen.get("descuGravada", 0.00),
        "monto_total_operacion": resumen.get(campo_monto_total, 0.00),
        "iva_perci1": resumen.get("ivaPerci1", 0.00),
    }
    if incluir_retencion:
        valores["iva_ret1"] = resumen.get("ivaRete1", 0.00)
    return valores


def iterar_tributos_conocidos(resumen):
    for tributo in resumen.get("tributos") or []:
        clave = TRIBUTOS_CONOCIDOS.get(tributo.get("codigo"))
        if clave:
            yield clave, tributo.get("valor", 0.00)


def calcular_valores_consumidor_final(resumen, *, campo_iva, clave_iva):
    return {
        "total_exenta": resumen.get("totalExenta", 0.00) - resumen.get("descuExenta", 0.00),
        "total_no_sujeta": resumen.get("totalNoSuj", 0.00) - resumen.get("descuNoSuj", 0.00),
        "total_no_gravado": resumen.get("totalNoGravado", 0.00),
        "total_gravada": resumen.get("totalGravada", 0.00) - resumen.get("descuGravada", 0.00),
        "monto_total_operacion": resumen.get("montoTotalOperacion", 0.00),
        clave_iva: resumen.get(campo_iva, 0.00),
    }


def obtener_total_compra_sujeto_excluido(resumen):
    return resumen.get("totalCompra", 0.00)
