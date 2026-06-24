import os

from openpyxl import Workbook
from openpyxl.styles import Font

from dte.calculations import calcular_valores_resumen, iterar_tributos_conocidos
from dte.extractor import (
    obtener_contenido_dte,
    obtener_fecha_formateada_dte,
    obtener_sello_recibido,
)
from dte.reader import cargar_datos_archivo
from dte.types import (
    DTE_COMPRA,
    DTE_COMPROBANTE_CREDITO_FISCAL,
    DTE_FACTURA,
    DTE_NOTA_CREDITO,
    DTE_NOTA_DEBITO,
)


def _crear_contadores():
    return {
        "sum_factura": 0,
        "sum_ntcdebito": 0,
        "sum_ccf": 0,
        "sum_ntccredito": 0,
        "sum_comproretencion": 0,
        "sum_grav": 0.0,
        "sum_iva": 0.0,
        "sum_total": 0.0,
        "sum_fovial": 0.0,
        "sum_cotrans": 0.0,
        "sum_docjson": 0,
    }


def _crear_valores():
    return {
        "fovial": 0,
        "cotrans": 0,
        "total_exenta": 0,
        "total_no_sujeta": 0,
        "sub_total_ventas": 0,
        "iva": 0,
        "monto_total_operacion": 0,
        "iva_perci1": 0,
    }


def generar_anexo_contribuyentes(carpeta, salida, nombre_archivo_salida, progreso=None):
    contadores = _crear_contadores()
    archivos_no_json = []
    wb = Workbook()
    ws_resumen = wb.active
    ws_resumen.title = "Ventas a Contribuyentes"
    ws_errores = wb.create_sheet(title="Errores")

    ws_resumen.append([
        "Fecha", "Clase de Documento", "Tipo de DTE", "Num_Control", "Num_Serie", "Num_Doc",
        "Num_Interno", "NRC", "Nom_Proveedor", "Exenta", "NoSuj", "Ventas Gravadas",
        "Debito Fiscal", "Venta a Terceros No Domiciliados ", "Debito Fiscal a terceros",
        "Total Ventas", "DUI", "Tipo_OP", "Tipo_I", "Num_Anexo",
    ])
    ws_errores.append(["Nombre del Archivo", "Error"])

    for hoja in (ws_resumen, ws_errores):
        for cell in hoja[1]:
            cell.font = Font(bold=True)

    nombres_archivos = os.listdir(carpeta)
    total_archivos = len(nombres_archivos)
    for indice, nombre_archivo in enumerate(nombres_archivos, start=1):
        if progreso:
            progreso(indice, total_archivos, nombre_archivo)

        ruta_archivo = os.path.join(carpeta, nombre_archivo)
        if not os.path.isfile(ruta_archivo) or not nombre_archivo.lower().endswith(".json"):
            archivos_no_json.append((nombre_archivo, "No es un archivo JSON"))
            continue

        print(f"Procesando archivo: {nombre_archivo}")
        datos = cargar_datos_archivo(ruta_archivo)
        contenido = obtener_contenido_dte(datos, "receptor") if datos else None
        if not contenido:
            archivos_no_json.append((nombre_archivo, "Error al cargar JSON o campos faltantes"))
            continue

        valores = _crear_valores()
        receptor = contenido["receptor"]
        identificacion = contenido["identificacion"]
        resumen = contenido["resumen"]
        tipo_dte = identificacion.get("tipoDte")
        if tipo_dte not in DTE_COMPRA:
            contadores["sum_docjson"] += 1
            continue

        if tipo_dte == DTE_NOTA_CREDITO:
            campo_monto_total = "montoTotalOperacion"
        else:
            campo_monto_total = "totalPagar"

        valores.update(calcular_valores_resumen(
            resumen,
            campo_no_sujeta="totalNoGravado",
            campo_monto_total=campo_monto_total,
            clave_gravada="sub_total_ventas",
        ))
        for clave, valor in iterar_tributos_conocidos(resumen):
            valores[clave] = valor
            contadores[f"sum_{clave}"] += valor

        if tipo_dte == DTE_COMPROBANTE_CREDITO_FISCAL:
            contadores["sum_ccf"] += 1
            contadores["sum_grav"] += valores["sub_total_ventas"]
            contadores["sum_total"] += valores["monto_total_operacion"]
        if tipo_dte == DTE_NOTA_CREDITO:
            contadores["sum_ntccredito"] += 1
            contadores["sum_grav"] -= valores["sub_total_ventas"]
            contadores["sum_total"] -= valores["monto_total_operacion"]
        if tipo_dte == DTE_NOTA_DEBITO:
            contadores["sum_ntcdebito"] += 1
            contadores["sum_grav"] += valores["sub_total_ventas"]
            contadores["sum_total"] += valores["monto_total_operacion"]
        if tipo_dte == DTE_FACTURA:
            contadores["sum_factura"] += 1

        _, fecha_formateada = obtener_fecha_formateada_dte(identificacion.get("fecEmi"))
        ws_resumen.append([
            fecha_formateada,
            "4",
            tipo_dte,
            identificacion.get("numeroControl"),
            obtener_sello_recibido(datos),
            identificacion.get("codigoGeneracion"),
            "",
            receptor.get("nrc"),
            receptor.get("nombre"),
            valores["total_exenta"],
            valores["total_no_sujeta"],
            valores["sub_total_ventas"],
            valores["iva"],
            0,
            0,
            valores["monto_total_operacion"],
            "",
            "",
            "",
            1,
        ])
        contadores["sum_docjson"] += 1

    wb.save(os.path.join(salida, f"{nombre_archivo_salida}.xlsx"))
    if archivos_no_json:
        for archivo in archivos_no_json:
            ws_errores.append(archivo)
        wb.save(os.path.join(salida, f"{nombre_archivo_salida}_errores.xlsx"))

    return [
        ("Facturas", contadores["sum_factura"]),
        ("Notas de Crédito", contadores["sum_ntccredito"]),
        ("Notas de Débito", contadores["sum_ntcdebito"]),
        ("Comprobantes de Crédito Fiscal", contadores["sum_ccf"]),
        ("Comprobantes de Retención", contadores["sum_comproretencion"]),
        ("Documentos Procesados", contadores["sum_docjson"]),
    ]
