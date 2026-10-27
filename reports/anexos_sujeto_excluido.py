import os

from openpyxl import Workbook
from openpyxl.styles import Font

from dte.calculations import obtener_total_compra_sujeto_excluido
from dte.extractor import (
    clasificar_documento_sujeto_excluido,
    obtener_contenido_dte,
    obtener_fecha_formateada_dte,
    obtener_sello_recibido,
)
from dte.reader import cargar_datos_archivo
from dte.duplicates import ControlDuplicadosDTE
from dte.types import DTE_SUJETO_EXCLUIDO


def generar_anexo_sujeto_excluido(carpeta, salida, nombre_archivo_salida, progreso=None):
    total_sujeto_excluido = 0
    duplicados = ControlDuplicadosDTE()
    total_documentos = 0
    archivos_no_json = []
    wb = Workbook()
    ws_resumen = wb.active
    ws_resumen.title = "Sujeto Excluido"
    ws_errores = wb.create_sheet(title="Errores")

    ws_resumen.append([
        "Tipo_Documento", "Num_NIT/DUI", "Razon Social", "F_Emision", "Num_Serie", "Num_Doc",
        "Monto_Operacion", "Monto_Retención(13%)", "Tipo_Op", "Clasificacion", "Sector",
        "Tipo Costo/Gasto", "Num_Anexo",
    ])
    ws_errores.append(["Nombre del Archivo", "Error"])
    for hoja in (ws_resumen, ws_errores):
        for cell in hoja[1]:
            cell.font = Font(bold=True)

    nombres_archivos = sorted(os.listdir(carpeta), key=str.casefold)
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
        contenido = obtener_contenido_dte(datos, "sujetoExcluido") if datos else None
        if not contenido:
            archivos_no_json.append((nombre_archivo, "Error al cargar JSON o campos faltantes"))
            continue

        sujeto_excluido = contenido["sujetoExcluido"]
        identificacion = contenido["identificacion"]
        tipo_dte = identificacion.get("tipoDte")
        numero_documento = sujeto_excluido.get("numDocumento")
        _, fecha_formateada = obtener_fecha_formateada_dte(identificacion.get("fecEmi"))

        if tipo_dte == DTE_SUJETO_EXCLUIDO:
            if duplicados.es_duplicado(identificacion, nombre_archivo):
                continue
            total_sujeto_excluido += 1
            ws_resumen.append([
                clasificar_documento_sujeto_excluido(numero_documento),
                numero_documento,
                sujeto_excluido.get("nombre"),
                fecha_formateada,
                obtener_sello_recibido(datos),
                identificacion.get("codigoGeneracion"),
                obtener_total_compra_sujeto_excluido(contenido["resumen"]),
                0,
                "",
                "",
                "",
                "",
                5,
            ])
        total_documentos += 1

    archivos_no_json.extend(duplicados.avisos)
    for archivo in archivos_no_json:
        ws_errores.append(archivo)
    wb.save(os.path.join(salida, f"{nombre_archivo_salida}.xlsx"))
    if archivos_no_json:
        wb.save(os.path.join(salida, f"{nombre_archivo_salida}_errores.xlsx"))

    return [
        ("Facturas", 0),
        ("Notas de Crédito", 0),
        ("Notas de Débito", 0),
        ("Comprobantes de Crédito Fiscal", 0),
        ("Comprobantes de Retención", 0),
        ("Sujeto Excluido", total_sujeto_excluido),
        ("Documentos Procesados", total_documentos),
    ]
