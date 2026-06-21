import os
import queue
import threading
from openpyxl import Workbook
from collections import defaultdict
from openpyxl.styles import Font
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import sys
from libroIVA import create_secondary_window,centrar_ventana
from dte.calculations import (
    calcular_valores_consumidor_final,
    calcular_valores_resumen,
    iterar_tributos_conocidos,
    obtener_total_compra_sujeto_excluido,
)
from dte.extractor import (
    clasificar_documento_sujeto_excluido,
    obtener_contenido_dte,
    obtener_fecha_formateada_dte,
    obtener_sello_recibido,
)
from dte.reader import cargar_datos_archivo
from dte.types import (
    DTE_ANULABLES,
    DTE_COMPRA,
    DTE_COMPROBANTE_CREDITO_FISCAL,
    DTE_CONSUMIDOR_FINAL,
    DTE_FACTURA,
    DTE_FACTURA_EXPORTACION,
    DTE_NOTA_CREDITO,
    DTE_NOTA_DEBITO,
    DTE_SUJETO_EXCLUIDO,
    TIPO_DOCUMENTO_ANULADOS,
    TIPO_DOCUMENTO_COMPRAS,
    TIPO_DOCUMENTO_CONSUMIDOR_FINAL,
    TIPO_DOCUMENTO_CONTRIBUYENTES,
    TIPO_DOCUMENTO_RETENCION,
    TIPO_DOCUMENTO_SUJETO_EXCLUIDO,
    TIPOS_DOCUMENTO,
)
from licensing.license_manager import validar_licencia, activar_gui
from reports.anexos_compras import generar_anexo_compras
from reports.anexos_retencion import generar_anexo_retencion


def variables_crecientes():
    return {
        'sum_factura': 0,
        'sum_export': 0,
        'sum_ntcdebito': 0,
        'sum_ccf': 0,
        'sum_ntccredito': 0,
        'sum_comproretencion': 0,
        'sum_sujetoexcluido': 0,
        'sum_grav': 0.0,
        'sum_iva': 0.0,
        'sum_total': 0.0,
        'sum_fovial': 0.0,
        'sum_cotrans': 0.0,
        'sum_docjson': 0
    }

def inicializar_valores():
    return {
        'fovial': 0,
        'cotrans': 0,
        'total_exenta': 0,
        'total_gravada': 0,
        'total_descuento_gravado': 0,
        'total_no_sujeta': 0,
        'total_no_gravado:': 0,
        'sub_total_ventas': 0,
        'iva': 0,
        'monto_total_operacion': 0,
        'iva_perci1': 0,
        'total_sujeto_retencion': 0,
        'total_iva_retenido': 0,
        'total_Compra':0
    }

def procesar_carpeta(carpeta, salida, nombre_archivo_salida, tree, tipo_documento):
    var_creciente = variables_crecientes()
    archivos_no_json = []
    if tipo_documento == TIPO_DOCUMENTO_COMPRAS:
        doc_types = generar_anexo_compras(carpeta, salida, nombre_archivo_salida)
        tree.delete(*tree.get_children())
        for item in doc_types:
            tree.insert("", tk.END, values=item)
        messagebox.showinfo("Compras", "Los archivos han sido procesados exitosamente.")
    elif tipo_documento == TIPO_DOCUMENTO_RETENCION:
        doc_types = generar_anexo_retencion(carpeta, salida, nombre_archivo_salida)
        tree.delete(*tree.get_children())
        for item in doc_types:
            tree.insert("", tk.END, values=item)
        messagebox.showinfo("Retención", "Comprobantes de retención procesados correctamente.")
    elif tipo_documento == TIPO_DOCUMENTO_CONTRIBUYENTES:
        # Espacio reservado para implementar lógica de procesamiento para Contribuyentes
        # Nuevo Libro de Trabajo
        wb = Workbook()
        # Creando hojas
        ws_resumen = wb.active
        ws_resumen.title = "Ventas a Contribuyentes"
        ws_errores = wb.create_sheet(title="Errores")

        # Títulos de las hojas
        encabezados_resumen = ["Fecha", "Clase de Documento", "Tipo de DTE", "Num_Control", "Num_Serie", "Num_Doc","Num_Interno","NRC","Nom_Proveedor",
                               "Exenta","NoSuj", "Ventas Gravadas", "Debito Fiscal",
                               "Venta a Terceros No Domiciliados ", "Debito Fiscal a terceros",
                               "Total Ventas", "DUI", "Tipo_OP","Tipo_I", "Num_Anexo"]
        encabezados_errores = ["Nombre del Archivo", "Error"]

        ws_resumen.append(encabezados_resumen)
        ws_errores.append(encabezados_errores)

        for cell in ws_resumen[1]:
            cell.font = Font(bold=True)

        for cell in ws_errores[1]:
            cell.font = Font(bold=True)

        for nombre_archivo in os.listdir(carpeta):
            ruta_archivo = os.path.join(carpeta, nombre_archivo)
            if os.path.isfile(ruta_archivo) and nombre_archivo.lower().endswith('.json'):
                print(f"Procesando archivo: {nombre_archivo}")
                datos = cargar_datos_archivo(ruta_archivo)
                contenido = obtener_contenido_dte(datos, "receptor") if datos else None
                if contenido:
                    clase_doc = "4" #B
                    valores = inicializar_valores()
                    receptor = contenido['receptor']
                    identificacion = contenido['identificacion']
                    resumen = contenido['resumen']
                    fecEmi = identificacion.get('fecEmi') #A
                    tipoDte = identificacion.get('tipoDte') #C
                    num_control = identificacion.get('numeroControl') #D

                    sello_recibido = obtener_sello_recibido(datos)

                    codigoGeneracion = identificacion.get('codigoGeneracion')
                    nrc = receptor.get('nrc')
                    nombre_receptor = receptor.get('nombre')

                    _, formatted_date = obtener_fecha_formateada_dte(fecEmi)

                    if tipoDte in DTE_COMPRA:
                        if tipoDte == DTE_NOTA_CREDITO:
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
                            var_creciente[f"sum_{clave}"] += valor

                        if tipoDte == DTE_COMPROBANTE_CREDITO_FISCAL:
                            var_creciente['sum_ccf'] += 1
                            var_creciente['sum_grav'] += valores['sub_total_ventas']
                            var_creciente['sum_total'] += valores['monto_total_operacion']
                        if tipoDte == DTE_NOTA_CREDITO:
                            var_creciente['sum_ntccredito'] += 1
                            var_creciente['sum_grav'] -= valores['sub_total_ventas']
                            var_creciente['sum_total'] -= valores['monto_total_operacion']
                        if tipoDte == DTE_NOTA_DEBITO:
                            var_creciente['sum_ntcdebito'] += 1
                            var_creciente['sum_grav'] += valores['sub_total_ventas']
                            var_creciente['sum_total'] += valores['monto_total_operacion']
                        if tipoDte == DTE_FACTURA:
                            var_creciente['sum_factura'] += 1

                        # Agregar datos al resumen en Excel
                        fila_resumen = [formatted_date, clase_doc, tipoDte, num_control, sello_recibido, codigoGeneracion,"",
                                        nrc, nombre_receptor, valores['total_exenta'], valores['total_no_sujeta'],
                                        valores['sub_total_ventas'], valores['iva'],0,0, valores['monto_total_operacion'],"","","", 1]
                        ws_resumen.append(fila_resumen)
                    var_creciente['sum_docjson'] += 1
                else:
                    archivos_no_json.append((nombre_archivo, "Error al cargar JSON o campos faltantes"))
            else:
                archivos_no_json.append((nombre_archivo, "No es un archivo JSON"))
                # Actualizar Treeview
        doc_types = [
              ("Facturas", var_creciente['sum_factura']),
              ("Notas de Crédito", var_creciente['sum_ntccredito']),
              ("Notas de Débito", var_creciente['sum_ntcdebito']),
              ("Comprobantes de Crédito Fiscal", var_creciente['sum_ccf']),
              ("Comprobantes de Retención", var_creciente['sum_comproretencion']),
              ("Documentos Procesados", var_creciente['sum_docjson']),
        ]

        # Limpiar Treeview antes de insertar nuevos datos
        tree.delete(*tree.get_children())

        for item in doc_types:
            tree.insert("", tk.END, values=item)

        # Guardar el archivo Excel
        wb.save(os.path.join(salida, f"{nombre_archivo_salida}.xlsx"))

        if archivos_no_json:
            for archivo in archivos_no_json:
                ws_errores.append(archivo)
            wb.save(os.path.join(salida, f"{nombre_archivo_salida}_errores.xlsx"))

        messagebox.showinfo("Ventas a Contribuyentes", "Los archivos han sido procesados exitosamente.")

    elif tipo_documento == TIPO_DOCUMENTO_CONSUMIDOR_FINAL:
        # Espacio reservado para implementar lógica de procesamiento para Consumidor Final
        # Nuevo Libro de Trabajo
        wb = Workbook()
        # Creando hojas
        ws_resumen = wb.active
        ws_resumen.title = "Consumidor Final"
        ws_errores = wb.create_sheet(title="Errores")
        # Títulos de las hojas

        encabezados_resumen = ["Fecha", "Clase de Doc", "Tipo de DTE", "Num_Resolución", "Num_Serie", "Num_ControlInterno(DEL)","Num_ControlInterno(AL)",
                               "Num_Doc(Del)", "Num_Doc(Al)", "N_Maquina","V_Exenta","V_InternaNS", "V_NoSuj", "Ventas Gravadas","Export Dentro C.A.","Export Fuera C.A.",
                               "Export Serv ", "Ventas a zonas francas", "Venta_Terceros", "Total Ventas", "Tipo_OP","Tipo_I", "Num_Anexo"]
        encabezados_errores = ["Nombre del Archivo", "Error"]

        ws_resumen.append(encabezados_resumen)
        ws_errores.append(encabezados_errores)

        for cell in ws_resumen[1]:
            cell.font = Font(bold=True)

        for cell in ws_errores[1]:
            cell.font = Font(bold=True)

        agrupados = defaultdict(list)
        for nombre_archivo in os.listdir(carpeta):
            ruta_archivo = os.path.join(carpeta, nombre_archivo)
            if os.path.isfile(ruta_archivo) and nombre_archivo.lower().endswith('.json'):
                datos = cargar_datos_archivo(ruta_archivo)
                contenido = obtener_contenido_dte(datos, "receptor") if datos else None
                if contenido:
                    identificacion = contenido['identificacion']
                    resumen = contenido['resumen']
                    tipoDte = identificacion.get('tipoDte')
                    fecEmi = identificacion.get('fecEmi')
                    horEmi= identificacion.get('horEmi')
                    codigoGeneracion = identificacion.get('codigoGeneracion')

                    if fecEmi and tipoDte in DTE_CONSUMIDOR_FINAL:
                        _, formatted_date = obtener_fecha_formateada_dte(fecEmi)
                        clave = (formatted_date, tipoDte)
                        valores_consumidor = calcular_valores_consumidor_final(
                            resumen,
                            campo_iva="ivaRetenido",
                            clave_iva="iva",
                        )
                        agrupados[clave].append({
                            "codigo": codigoGeneracion,
                            "horEmi": horEmi,
                            **valores_consumidor,
                        })
                        var_creciente['sum_docjson'] += 1
                else:
                    archivos_no_json.append((nombre_archivo, "Error al cargar JSON o campos faltantes"))
            else:
                archivos_no_json.append((nombre_archivo, "No es un archivo JSON"))
                # Actualizar Treeview
        for (fecha, tipoDte), documentos in agrupados.items():
            documentos_ordenados = sorted(documentos, key=lambda x: x.get("horEmi", "00:00:00"))

            cod_inicio = documentos_ordenados[0]["codigo"]
            cod_fin = documentos_ordenados[-1]["codigo"]

            total_exenta = sum(d["total_exenta"] for d in documentos)
            total_no_sujeta = sum(d["total_no_sujeta"] for d in documentos)
            total_no_gravado = sum(d["total_no_gravado"] for d in documentos)
            total_gravada = sum(d["total_gravada"] for d in documentos)
            monto_total_operacion = sum(d["monto_total_operacion"] for d in documentos)
            iva_total = sum(d["iva"] for d in documentos)

            fila_resumen = [
                fecha, "4", tipoDte, "N/A", "N/A", "N/A", "N/A",
                cod_inicio, cod_fin, ""
            ]

            if tipoDte == DTE_FACTURA:
                var_creciente['sum_factura'] += len(documentos)
                fila_resumen += [
                    total_exenta, 0, total_no_sujeta,
                    total_gravada, 0, 0, 0, 0, 0,
                    monto_total_operacion, "", "", 2
                ]
            elif tipoDte == DTE_FACTURA_EXPORTACION:
                var_creciente['sum_export'] += len(documentos)
                fila_resumen += [
                    total_exenta, 0, total_no_sujeta,0,
                    total_gravada, 0,
                    0, 0,0,
                    monto_total_operacion,"","", 2
                ]

            ws_resumen.append(fila_resumen)

        doc_types = [
              ("Facturas", var_creciente['sum_factura']),
              ("Facturas de exportacion", var_creciente['sum_export']),
              ("Notas de Crédito", var_creciente['sum_ntccredito']),
              ("Notas de Débito", var_creciente['sum_ntcdebito']),
              ("Comprobantes de Crédito Fiscal", var_creciente['sum_ccf']),
              ("Comprobantes de Retención", var_creciente['sum_comproretencion']),
              ("Documentos Procesados", var_creciente['sum_docjson']),
        ]

        # Limpiar Treeview antes de insertar nuevos datos
        tree.delete(*tree.get_children())

        for item in doc_types:
            tree.insert("", tk.END, values=item)

        # Guardar el archivo Excel
        wb.save(os.path.join(salida, f"{nombre_archivo_salida}.xlsx"))

        if archivos_no_json:
            for archivo in archivos_no_json:
                ws_errores.append(archivo)
            wb.save(os.path.join(salida, f"{nombre_archivo_salida}_errores.xlsx"))

        messagebox.showinfo("Ventas a Consumidor Final", "Los archivos han sido procesados exitosamente.")

    elif tipo_documento == TIPO_DOCUMENTO_SUJETO_EXCLUIDO:
        wb = Workbook()
        # Creando hojas
        ws_resumen = wb.active
        ws_resumen.title = "Sujeto Excluido"
        ws_errores = wb.create_sheet(title="Errores")
        encabezados_resumen = ["Tipo_Documento","Num_NIT/DUI","Razon Social", "F_Emision","Num_Serie","Num_Doc","Monto_Operacion","Monto_Retención(13%)",
                               "Tipo_Op","Clasificacion","Sector","Tipo Costo/Gasto","Num_Anexo"]
        encabezados_errores = ["Nombre del Archivo", "Error"]

        ws_resumen.append(encabezados_resumen)
        ws_errores.append(encabezados_errores)

        for cell in ws_resumen[1]:
            cell.font = Font(bold=True)

        for cell in ws_errores[1]:
            cell.font = Font(bold=True)
        for nombre_archivo in os.listdir(carpeta):
            ruta_archivo = os.path.join(carpeta, nombre_archivo)
            if os.path.isfile(ruta_archivo) and nombre_archivo.lower().endswith('.json'):
                print(f"Procesando archivo: {nombre_archivo}")
                datos = cargar_datos_archivo(ruta_archivo)
                contenido = obtener_contenido_dte(datos, "sujetoExcluido") if datos else None
                if contenido:
                    clase_doc = "4" #B

                    valores = inicializar_valores()
                    sujetoExcluido = contenido['sujetoExcluido']
                    identificacion = contenido['identificacion']
                    resumen = contenido['resumen']
                    fecEmi = identificacion.get('fecEmi') #A
                    tipoDte = identificacion.get('tipoDte') #C
                    numDocumento = sujetoExcluido.get('numDocumento')
                    nombre = sujetoExcluido.get('nombre')
                    sello_recibido = obtener_sello_recibido(datos)
                    codigoGeneracion = identificacion.get('codigoGeneracion')

                    tipoDocumento = clasificar_documento_sujeto_excluido(numDocumento)
                    _, formatted_date = obtener_fecha_formateada_dte(fecEmi)

                    if tipoDte == DTE_SUJETO_EXCLUIDO:
                        var_creciente['sum_sujetoexcluido'] += 1
                        valores['total_Compra'] = obtener_total_compra_sujeto_excluido(resumen)

                        # Agregar datos al resumen en Excel
                        fila_resumen = [tipoDocumento,numDocumento,nombre,formatted_date,sello_recibido, codigoGeneracion, valores['total_Compra'],0,"","","","",5]
                        ws_resumen.append(fila_resumen)
                    var_creciente['sum_docjson'] += 1
                else:
                    archivos_no_json.append((nombre_archivo, "Error al cargar JSON o campos faltantes"))
            else:
                archivos_no_json.append((nombre_archivo, "No es un archivo JSON"))
                # Actualizar Treeview
        doc_types = [
              ("Facturas", var_creciente['sum_factura']),
              ("Notas de Crédito", var_creciente['sum_ntccredito']),
              ("Notas de Débito", var_creciente['sum_ntcdebito']),
              ("Comprobantes de Crédito Fiscal", var_creciente['sum_ccf']),
              ("Comprobantes de Retención", var_creciente['sum_comproretencion']),
              ("Sujeto Excluido", var_creciente['sum_sujetoexcluido']),
              ("Documentos Procesados", var_creciente['sum_docjson']),
        ]

        # Limpiar Treeview antes de insertar nuevos datos
        tree.delete(*tree.get_children())

        for item in doc_types:
            tree.insert("", tk.END, values=item)

        # Guardar el archivo Excel
        wb.save(os.path.join(salida, f"{nombre_archivo_salida}.xlsx"))

        if archivos_no_json:
            for archivo in archivos_no_json:
                ws_errores.append(archivo)
            wb.save(os.path.join(salida, f"{nombre_archivo_salida}_errores.xlsx"))

        messagebox.showinfo("Facturas Sujeto Excluido", "Los archivos han sido procesados exitosamente.")
    elif tipo_documento == TIPO_DOCUMENTO_ANULADOS:
        # Nuevo Libro de Trabajo
        wb = Workbook()
        # Creando hojas
        ws_resumen = wb.active
        ws_resumen.title = "Documentos Anulados"
        ws_errores = wb.create_sheet(title="Errores")

        # Títulos de las hojas
        encabezado_anulados = ["Num_Resolucion", "Clase_Doc","Desde(Impreso)","Hasta(Impreso)","Tipo_doc","Tipo_det","Serie","Desde","Hasta","Cod_Generacion"]
        encabezados_errores = ["Nombre del Archivo", "Error"]

        ws_resumen.append(encabezado_anulados)
        ws_errores.append(encabezados_errores)

        for cell in ws_resumen[1]:
            cell.font = Font(bold=True)

        for cell in ws_errores[1]:
            cell.font = Font(bold=True)

        for nombre_archivo in os.listdir(carpeta):
            ruta_archivo = os.path.join(carpeta, nombre_archivo)
            if os.path.isfile(ruta_archivo) and nombre_archivo.lower().endswith('.json'):
                print(f"Procesando archivo: {nombre_archivo}")
                datos = cargar_datos_archivo(ruta_archivo)
                if not datos:
                    continue

                contenido = obtener_contenido_dte(datos, "emisor")

                if not contenido:
                    print(f"⚠️ Estructura no válida en el archivo {nombre_archivo}")
                    continue

                identificacion = contenido['identificacion']

                numerocontrol = identificacion.get('numeroControl')
                clase_doc = "4"
                tipoDte = identificacion.get('tipoDte')
                tipo_detalle = "D"
                sello_recibido = obtener_sello_recibido(contenido)
                codigoGeneracion = identificacion.get('codigoGeneracion')

                if tipoDte in DTE_ANULABLES:
                    if tipoDte == DTE_FACTURA:
                        var_creciente['sum_factura'] += 1
                    if tipoDte == DTE_COMPROBANTE_CREDITO_FISCAL:
                        var_creciente['sum_ccf'] +=1
                    if tipoDte == DTE_NOTA_CREDITO:
                        var_creciente['sum_ntccredito'] += 1
                    if tipoDte == DTE_FACTURA_EXPORTACION:
                        var_creciente['sum_export'] += 1
                    if tipoDte == DTE_SUJETO_EXCLUIDO:
                        var_creciente['sum_sujetoexcluido'] += 1
                    fila_anulado = [numerocontrol,clase_doc,0,0,tipoDte,tipo_detalle,sello_recibido,0,0,codigoGeneracion]
                    ws_resumen.append(fila_anulado)
                    var_creciente['sum_docjson'] += 1

                else:
                    archivos_no_json.append((nombre_archivo, "Error al cargar JSON o campos faltantes"))
            else:
                archivos_no_json.append((nombre_archivo, "No es un archivo JSON"))
        doc_types = [
            ("Facturas", var_creciente['sum_factura']),
            ("Comprobantes de Crédito Fiscal", var_creciente['sum_ccf']),
            ("Notas de Crédito", var_creciente['sum_ntccredito']),
            ("Notas de Débito", var_creciente['sum_ntcdebito']),
            ("Facturas de exportacion", var_creciente['sum_export']),
            ("Sujeto Excluido", var_creciente['sum_sujetoexcluido']),
            ("Documentos Procesados", var_creciente['sum_docjson']),
        ]
        # Limpiar Treeview antes de insertar nuevos datos
        tree.delete(*tree.get_children())

        for item in doc_types:
            tree.insert("", tk.END, values=item)

        # Guardar el archivo Excel
        wb.save(os.path.join(salida, f"{nombre_archivo_salida}.xlsx"))

        if archivos_no_json:
            for archivo in archivos_no_json:
                ws_errores.append(archivo)
            wb.save(os.path.join(salida, f"{nombre_archivo_salida}_errores.xlsx"))

        messagebox.showinfo("DTEs Anulados", "Han sido procesados exitosamente.")
    elif tipo_documento not in TIPOS_DOCUMENTO:
        messagebox.showerror("Error", "Debe seleccionar un tipo de documento válido.")
        return

def procesar_reporte_en_segundo_plano(
    root,
    carpeta,
    salida,
    nombre_archivo_salida,
    tree,
    progreso_var,
    estado_var,
    boton_iniciar,
    generar_reporte,
    titulo_reporte,
    mensaje_exito,
):
    eventos = queue.Queue()
    boton_iniciar.config(state="disabled")
    progreso_var.set(0)
    estado_var.set(f"Preparando {titulo_reporte}...")

    def reportar_progreso(actual, total, archivo):
        eventos.put(("progreso", actual, total, archivo))

    def ejecutar_reporte():
        try:
            resumen = generar_reporte(
                carpeta,
                salida,
                nombre_archivo_salida,
                progreso=reportar_progreso,
            )
        except Exception as error:
            eventos.put(("error", str(error)))
        else:
            eventos.put(("completado", resumen))

    def procesar_eventos():
        finalizado = False
        while True:
            try:
                evento = eventos.get_nowait()
            except queue.Empty:
                break

            tipo_evento = evento[0]
            if tipo_evento == "progreso":
                _, actual, total, archivo = evento
                progreso_var.set((actual / total) * 100 if total else 0)
                estado_var.set(f"Procesando {actual} de {total}: {archivo}")
            elif tipo_evento == "completado":
                _, resumen = evento
                tree.delete(*tree.get_children())
                for item in resumen:
                    tree.insert("", tk.END, values=item)
                estado_var.set(f"Procesamiento de {titulo_reporte} finalizado.")
                messagebox.showinfo(titulo_reporte, mensaje_exito)
                finalizado = True
            elif tipo_evento == "error":
                _, mensaje = evento
                estado_var.set(f"El procesamiento de {titulo_reporte} no pudo completarse.")
                messagebox.showerror(titulo_reporte, f"No se pudo procesar la carpeta:\n{mensaje}")
                finalizado = True

        if finalizado:
            boton_iniciar.config(state="normal")
        else:
            root.after(100, procesar_eventos)

    threading.Thread(target=ejecutar_reporte, daemon=True).start()
    root.after(100, procesar_eventos)


def iniciar_procesamiento(
    root,
    carpeta_entrada,
    carpeta_salida,
    nombre_archivo,
    tipo_documento,
    tree,
    progreso_var,
    estado_var,
    boton_iniciar,
):
    carpeta = carpeta_entrada.get()
    salida = carpeta_salida.get()
    nombre_archivo_salida = nombre_archivo.get()
    tipo_documento_seleccionado = tipo_documento.get()  # Obtener el valor seleccionado del ComboBox
    print(tipo_documento_seleccionado)

    if carpeta and salida and nombre_archivo_salida and tipo_documento_seleccionado:
        if tipo_documento_seleccionado == TIPO_DOCUMENTO_COMPRAS:
            procesar_reporte_en_segundo_plano(
                root,
                carpeta,
                salida,
                nombre_archivo_salida,
                tree,
                progreso_var,
                estado_var,
                boton_iniciar,
                generar_anexo_compras,
                "Compras",
                "Los archivos han sido procesados exitosamente.",
            )
            return

        if tipo_documento_seleccionado == TIPO_DOCUMENTO_RETENCION:
            procesar_reporte_en_segundo_plano(
                root,
                carpeta,
                salida,
                nombre_archivo_salida,
                tree,
                progreso_var,
                estado_var,
                boton_iniciar,
                generar_anexo_retencion,
                "Retención",
                "Comprobantes de retención procesados correctamente.",
            )
            return

        procesar_carpeta(carpeta, salida, nombre_archivo_salida, tree, tipo_documento_seleccionado)
    else:
        # Mostrar advertencia si falta algún campo
        messagebox.showwarning(
            "Campos vacíos",
            "Por favor, complete todos los campos y seleccione un tipo de documento antes de iniciar el procesamiento."
        )

def limpiar_campos(tipo_documento,carpeta_entrada, carpeta_salida, nombre_archivo, tree):
    tipo_documento.set("Seleccionar...")
    carpeta_entrada.set("")  # Limpiar la variable StringVar
    carpeta_salida.set("")  # Limpiar la variable StringVar
    nombre_archivo.set("")  # Limpiar la variable StringVar
    tree.delete(*tree.get_children())  # Limpiar el árbol

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # Cuando está empaquetado con PyInstaller
    except Exception:
        base_path = os.path.abspath(".")  # Cuando se ejecuta como script .py
    return os.path.join(base_path, relative_path)

def primary_window():
    # Crear la interfaz gráfica
    root = tk.Tk()

    style = ttk.Style(root)
    style.theme_use('clam') 
    # Intentar usar .ico (Windows)
    try:
        root.iconbitmap(resource_path("iconos/icono_v3.ico"))
    except Exception:
        pass  # No romper la app

    # Usar PNG (SIEMPRE)
    root.icon_img = tk.PhotoImage(
        file=resource_path("iconos/icono_v3.png")
    )
    root.iconphoto(True, root.icon_img)
    root.title("Procesador de Archivos JSON")
    
    root.resizable(False, False) 
    #root.update_idletasks()
    centrar_ventana(root,600,460)
    
    
    # Configurar estilos para el Combobox
 # Usar el tema "clam" que permite personalización completa
    style.configure(
        "TCombobox",
        fieldbackground="red",  # Fondo de la caja de entrada
        background="white",  # Fondo del menú desplegable
        foreground="black",  # Texto negro
        bordercolor="gray",  # Bordes grises
        lightcolor="white",  # Efecto de luz
        darkcolor="white"  # Efecto oscuro
    )

    # Frame para los campos de entrada
    frame_entradas = tk.Frame(root)
    frame_entradas.pack(pady=10)
    frame_entradas.grid_columnconfigure(1, weight=1)

    # Crear las variables como StringVar
    tipo_documento = tk.StringVar()
    carpeta_entrada = tk.StringVar()
    carpeta_salida = tk.StringVar()
    nombre_archivo = tk.StringVar()

    # Etiqueta y ComboBox para seleccionar tipo de documento
    tk.Label(frame_entradas, text="Tipo de documentos:").grid(row=0, column=0, padx=5, pady=5)
    combo_tipo_documento = ttk.Combobox(
        frame_entradas,
        textvariable=tipo_documento,
        values=TIPOS_DOCUMENTO,
        state="readonly",
        width=47
    )
    combo_tipo_documento.grid(row=0, column=1, padx=5, pady=5)
    combo_tipo_documento.set("Seleccionar...")  # Valor predeterminado

    # Etiqueta y entrada para la carpeta de entrada
    tk.Label(frame_entradas, text="Carpeta de entrada:").grid(row=1, column=0, padx=5, pady=5)
    tk.Entry(frame_entradas, textvariable=carpeta_entrada, width=50).grid(row=1, column=1, padx=5, pady=5)
    tk.Button(frame_entradas, text="Buscar", command=lambda: carpeta_entrada.set(filedialog.askdirectory())).grid(row=1, column=2, padx=5, pady=5)

    # Etiqueta y entrada para la carpeta de salida
    tk.Label(frame_entradas, text="Carpeta de salida:").grid(row=2, column=0, padx=5, pady=5)
    tk.Entry(frame_entradas, textvariable=carpeta_salida, width=50).grid(row=2, column=1, padx=5, pady=5)
    tk.Button(frame_entradas, text="Buscar", command=lambda: carpeta_salida.set(filedialog.askdirectory())).grid(row=2, column=2, padx=5, pady=5)

    # Etiqueta y entrada para el nombre del archivo de salida
    tk.Label(frame_entradas, text="Nombre del archivo de salida (.xlsx):").grid(row=3, column=0, padx=5, pady=5)
    tk.Entry(frame_entradas, textvariable=nombre_archivo, width=50).grid(row=3, column=1, padx=5, pady=5)

    progreso_var = tk.DoubleVar(value=0)
    estado_var = tk.StringVar(value="")

    boton_iniciar = tk.Button(
        frame_entradas,
        text="Iniciar procesamiento",
        command=lambda: iniciar_procesamiento(
            root,
            carpeta_entrada,
            carpeta_salida,
            nombre_archivo,
            tipo_documento,
            tree,
            progreso_var,
            estado_var,
            boton_iniciar,
        ),
    )
    boton_iniciar.grid(row=4, column=1, pady=10)

    # Botón para limpiar los campos
    tk.Button(frame_entradas, text="Limpiar", command=lambda: limpiar_campos(tipo_documento, carpeta_entrada, carpeta_salida, nombre_archivo, tree)).grid(row=4, column=2, pady=10)

    progreso = ttk.Progressbar(
        frame_entradas,
        variable=progreso_var,
        maximum=100,
        mode="determinate",
        length=360,
    )
    progreso.grid(row=5, column=1, pady=(0, 3))
    tk.Label(
        frame_entradas,
        textvariable=estado_var,
        anchor="w",
        justify="left",
        wraplength=360,
    ).grid(
        row=6,
        column=1,
        pady=(0, 5),
    )

    # Frame para el Treeview
    frame_treeview = tk.Frame(root)
    frame_treeview.pack(pady=10)

    # Treeview para mostrar el resumen de documentos procesados
    tree = ttk.Treeview(frame_treeview, columns=("Tipo de Documento", "Cantidad"), show="headings")
    tree.heading("Tipo de Documento", text="Tipo de Documento")
    tree.heading("Cantidad", text="Cantidad")
    tree.pack()

    # Botón para generar libro de IVA
    tk.Button(frame_entradas, text="Generar Libro de IVA", command=lambda: create_secondary_window(root)).grid(row=4, column=0, pady=10)

    root.mainloop()


if __name__ == "__main__":
    # 🔐 BLOQUEO DE LICENCIA
    if not validar_licencia():
        activar_gui()
    # 🚀 TU APP REAL
    primary_window()
