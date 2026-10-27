import os

from openpyxl import Workbook
from openpyxl.styles import Font

from dte.extractor import obtener_contenido_dte, obtener_sello_recibido
from dte.reader import cargar_datos_archivo
from dte.duplicates import ControlDuplicadosDTE
from dte.types import (
    DTE_ANULABLES,
    DTE_COMPROBANTE_CREDITO_FISCAL,
    DTE_FACTURA,
    DTE_FACTURA_EXPORTACION,
    DTE_NOTA_CREDITO,
    DTE_SUJETO_EXCLUIDO,
)


def generar_anexo_anulados(carpeta, salida, nombre_archivo_salida, progreso=None):
    duplicados = ControlDuplicadosDTE()
    contadores = {
        "facturas": 0,
        "credito_fiscal": 0,
        "notas_credito": 0,
        "notas_debito": 0,
        "exportacion": 0,
        "sujeto_excluido": 0,
        "documentos": 0,
    }
    archivos_no_json = []
    wb = Workbook()
    ws_resumen = wb.active
    ws_resumen.title = "Documentos Anulados"
    ws_errores = wb.create_sheet(title="Errores")

    ws_resumen.append([
        "Num_Resolucion", "Clase_Doc", "Desde(Impreso)", "Hasta(Impreso)", "Tipo_doc",
        "Tipo_det", "Serie", "Desde", "Hasta", "Cod_Generacion",
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
        if not datos:
            continue

        contenido = obtener_contenido_dte(datos, "emisor")
        if not contenido:
            print(f"Estructura no valida en el archivo {nombre_archivo}")
            continue

        identificacion = contenido["identificacion"]
        tipo_dte = identificacion.get("tipoDte")
        if tipo_dte not in DTE_ANULABLES:
            archivos_no_json.append((nombre_archivo, "Error al cargar JSON o campos faltantes"))
            continue

        if duplicados.es_duplicado(identificacion, nombre_archivo):
            continue

        if tipo_dte == DTE_FACTURA:
            contadores["facturas"] += 1
        if tipo_dte == DTE_COMPROBANTE_CREDITO_FISCAL:
            contadores["credito_fiscal"] += 1
        if tipo_dte == DTE_NOTA_CREDITO:
            contadores["notas_credito"] += 1
        if tipo_dte == DTE_FACTURA_EXPORTACION:
            contadores["exportacion"] += 1
        if tipo_dte == DTE_SUJETO_EXCLUIDO:
            contadores["sujeto_excluido"] += 1

        ws_resumen.append([
            identificacion.get("numeroControl"),
            "4",
            0,
            0,
            tipo_dte,
            "D",
            obtener_sello_recibido(contenido),
            0,
            0,
            identificacion.get("codigoGeneracion"),
        ])
        contadores["documentos"] += 1

    archivos_no_json.extend(duplicados.avisos)
    for archivo in archivos_no_json:
        ws_errores.append(archivo)
    wb.save(os.path.join(salida, f"{nombre_archivo_salida}.xlsx"))
    if archivos_no_json:
        wb.save(os.path.join(salida, f"{nombre_archivo_salida}_errores.xlsx"))

    return [
        ("Facturas", contadores["facturas"]),
        ("Comprobantes de Crédito Fiscal", contadores["credito_fiscal"]),
        ("Notas de Crédito", contadores["notas_credito"]),
        ("Notas de Débito", contadores["notas_debito"]),
        ("Facturas de exportacion", contadores["exportacion"]),
        ("Sujeto Excluido", contadores["sujeto_excluido"]),
        ("Documentos Procesados", contadores["documentos"]),
    ]
