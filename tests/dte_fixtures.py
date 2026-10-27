"""Documentos sintéticos para comprobar reportes sin datos de clientes reales."""


def crear_dte(tipo="03", codigo="GEN-001", factor=1, fecha="2026-10-01"):
    return {
        "identificacion": {
            "tipoDte": tipo,
            "codigoGeneracion": codigo,
            "numeroControl": f"DTE-{tipo}-{codigo}",
            "fecEmi": fecha,
            "horEmi": "08:00:00",
        },
        "emisor": {"nombre": "Proveedor Prueba", "nrc": "123", "nit": "06140000000000"},
        "receptor": {"nombre": "Cliente Prueba", "nrc": "456"},
        "sujetoExcluido": {"nombre": "Sujeto Prueba", "numDocumento": "01234567-8"},
        "resumen": {
            "totalExenta": 2 * factor,
            "totalNoSuj": 2 * factor,
            "totalNoGravado": 2 * factor,
            "totalGravada": 3 * factor,
            "montoTotalOperacion": 7.39 * factor,
            "totalPagar": 7.39 * factor,
            "totalCompra": 7.39 * factor,
            "totalSujetoRetencion": 3 * factor,
            "totalIVAretenido": 0.03 * factor,
            "tributos": [{"codigo": "20", "valor": 0.39 * factor}],
        },
        "selloRecibido": "SELLO-PRUEBA",
    }
