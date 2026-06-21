import os

from openpyxl import Workbook
from openpyxl.styles import Font

from dte.extractor import (
    obtener_contenido_dte,
    obtener_fecha_formateada_dte,
    obtener_sello_recibido,
)
from dte.reader import cargar_datos_archivo
from dte.types import DTE_COMPROBANTE_RETENCION
from reports.invalidos_dte import guardar_dtes_invalidos
from services.hacienda import (
    ConsultaHaciendaError,
    consultar_estado_dte,
    dte_esta_transmitido_satisfactoriamente,
)


def generar_anexo_retencion(
    carpeta,
    salida,
    nombre_archivo_salida,
    consultar_estado=consultar_estado_dte,
    progreso=None,
):
    total_retenciones = 0
    total_documentos = 0
    archivos_no_json = []
    dtes_invalidos = []
    consultas_hacienda = {}
    wb = Workbook()
    ws_retencion = wb.active
    ws_retencion.title = "Comprobantes de Retención"
    ws_errores = wb.create_sheet(title="Errores")

    ws_retencion.append([
        "NIT", "Fecha", "Tipo DTE", "Sello Recepcion", "CodGeneracion",
        "Monto Sujeto", "Monto de Retención", "DUI", "Num_Anexo",
    ])
    ws_errores.append(["Nombre del Archivo", "Error"])

    for hoja in (ws_retencion, ws_errores):
        for cell in hoja[1]:
            cell.font = Font(bold=True)

    nombres_archivos = os.listdir(carpeta)
    total_archivos = len(nombres_archivos)

    for indice, nombre_archivo in enumerate(nombres_archivos, start=1):
        if progreso:
            progreso(indice, total_archivos, nombre_archivo)

        ruta_archivo = os.path.join(carpeta, nombre_archivo)

        if not nombre_archivo.lower().endswith(".json"):
            archivos_no_json.append((nombre_archivo, "No es un archivo JSON"))
            continue

        datos = cargar_datos_archivo(ruta_archivo)
        if not datos:
            archivos_no_json.append((nombre_archivo, "JSON inválido"))
            continue

        contenido = obtener_contenido_dte(datos, "emisor")
        if not contenido:
            archivos_no_json.append((nombre_archivo, "Estructura inválida"))
            continue

        identificacion = contenido["identificacion"]
        if identificacion.get("tipoDte") != DTE_COMPROBANTE_RETENCION:
            continue

        fecha_emi = identificacion.get("fecEmi")
        codigo_generacion = identificacion.get("codigoGeneracion")
        _, fecha_formateada = obtener_fecha_formateada_dte(fecha_emi)
        clave_consulta = (codigo_generacion, fecha_emi)
        try:
            respuesta_hacienda = consultas_hacienda.get(clave_consulta)
            if respuesta_hacienda is None:
                respuesta_hacienda = consultar_estado(codigo_generacion, fecha_emi)
                consultas_hacienda[clave_consulta] = respuesta_hacienda
        except ConsultaHaciendaError as error:
            archivos_no_json.append((
                nombre_archivo,
                f"No se pudo consultar Hacienda: {error}",
            ))
            continue

        if not dte_esta_transmitido_satisfactoriamente(respuesta_hacienda):
            dtes_invalidos.append([
                nombre_archivo,
                codigo_generacion,
                fecha_formateada,
                respuesta_hacienda.get("estadoDoc", "Sin estado"),
                respuesta_hacienda.get("descripcionEstado", ""),
            ])
            continue

        resumen = contenido["resumen"]
        emisor = contenido["emisor"]
        ws_retencion.append([
            emisor.get("nit"),
            fecha_formateada,
            DTE_COMPROBANTE_RETENCION,
            respuesta_hacienda.get("selloVal") or obtener_sello_recibido(contenido),
            codigo_generacion,
            resumen.get("totalSujetoRetencion", 0.00),
            resumen.get("totalIVAretenido", 0.00),
            "",
            7,
        ])
        total_retenciones += 1
        total_documentos += 1

    wb.save(os.path.join(salida, f"{nombre_archivo_salida}.xlsx"))

    if archivos_no_json:
        for archivo in archivos_no_json:
            ws_errores.append(archivo)
        wb.save(os.path.join(salida, f"{nombre_archivo_salida}_errores.xlsx"))

    guardar_dtes_invalidos(salida, dtes_invalidos)

    return [
        ("Comprobantes de Retención", total_retenciones),
        ("Documentos Procesados", total_documentos),
    ]
