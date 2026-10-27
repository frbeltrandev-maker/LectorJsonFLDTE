import os
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font

from dte.calculations import calcular_valores_consumidor_final
from dte.extractor import obtener_contenido_dte, obtener_fecha_formateada_dte
from dte.reader import cargar_datos_archivo
from dte.duplicates import ControlDuplicadosDTE
from dte.types import DTE_CONSUMIDOR_FINAL, DTE_FACTURA, DTE_FACTURA_EXPORTACION


def generar_anexo_consumidor_final(carpeta, salida, nombre_archivo_salida, progreso=None):
    total_facturas = 0
    duplicados = ControlDuplicadosDTE()
    total_exportaciones = 0
    total_documentos = 0
    archivos_no_json = []
    agrupados = defaultdict(list)

    wb = Workbook()
    ws_resumen = wb.active
    ws_resumen.title = "Consumidor Final"
    ws_errores = wb.create_sheet(title="Errores")

    ws_resumen.append([
        "Fecha", "Clase de Doc", "Tipo de DTE", "Num_Resolución", "Num_Serie",
        "Num_ControlInterno(DEL)", "Num_ControlInterno(AL)", "Num_Doc(Del)",
        "Num_Doc(Al)", "N_Maquina", "V_Exenta", "V_InternaNS", "V_NoSuj",
        "Ventas Gravadas", "Export Dentro C.A.", "Export Fuera C.A.", "Export Serv ",
        "Ventas a zonas francas", "Venta_Terceros", "Total Ventas", "Tipo_OP",
        "Tipo_I", "Num_Anexo",
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

        datos = cargar_datos_archivo(ruta_archivo)
        contenido = obtener_contenido_dte(datos, "receptor") if datos else None
        if not contenido:
            archivos_no_json.append((nombre_archivo, "Error al cargar JSON o campos faltantes"))
            continue

        identificacion = contenido["identificacion"]
        resumen = contenido["resumen"]
        tipo_dte = identificacion.get("tipoDte")
        fecha_emi = identificacion.get("fecEmi")
        if not fecha_emi or tipo_dte not in DTE_CONSUMIDOR_FINAL:
            continue

        if duplicados.es_duplicado(identificacion, nombre_archivo):
            continue

        _, fecha_formateada = obtener_fecha_formateada_dte(fecha_emi)
        valores = calcular_valores_consumidor_final(
            resumen,
            campo_iva="ivaRetenido",
            clave_iva="iva",
        )
        agrupados[(fecha_formateada, tipo_dte)].append({
            "codigo": identificacion.get("codigoGeneracion"),
            "horEmi": identificacion.get("horEmi"),
            **valores,
        })
        total_documentos += 1

    for (fecha, tipo_dte), documentos in agrupados.items():
        documentos_ordenados = sorted(documentos, key=lambda dte: dte.get("horEmi", "00:00:00"))
        codigo_inicio = documentos_ordenados[0]["codigo"]
        codigo_fin = documentos_ordenados[-1]["codigo"]
        total_exenta = sum(documento["total_exenta"] for documento in documentos)
        total_no_sujeta = sum(documento["total_no_sujeta"] for documento in documentos)
        total_gravada = sum(documento["total_gravada"] for documento in documentos)
        monto_total_operacion = sum(documento["monto_total_operacion"] for documento in documentos)

        fila_resumen = [
            fecha, "4", tipo_dte, "N/A", "N/A", "N/A", "N/A", codigo_inicio, codigo_fin, "",
        ]

        if tipo_dte == DTE_FACTURA:
            total_facturas += len(documentos)
            fila_resumen += [
                total_exenta, 0, total_no_sujeta, total_gravada, 0, 0, 0, 0, 0,
                monto_total_operacion, "", "", 2,
            ]
        elif tipo_dte == DTE_FACTURA_EXPORTACION:
            total_exportaciones += len(documentos)
            fila_resumen += [
                total_exenta, 0, total_no_sujeta, 0, total_gravada, 0, 0, 0, 0,
                monto_total_operacion, "", "", 2,
            ]

        ws_resumen.append(fila_resumen)

    archivos_no_json.extend(duplicados.avisos)
    for archivo in archivos_no_json:
        ws_errores.append(archivo)
    wb.save(os.path.join(salida, f"{nombre_archivo_salida}.xlsx"))
    if archivos_no_json:
        wb.save(os.path.join(salida, f"{nombre_archivo_salida}_errores.xlsx"))

    return [
        ("Facturas", total_facturas),
        ("Facturas de exportacion", total_exportaciones),
        ("Notas de Crédito", 0),
        ("Notas de Débito", 0),
        ("Comprobantes de Crédito Fiscal", 0),
        ("Comprobantes de Retención", 0),
        ("Documentos Procesados", total_documentos),
    ]
