import os
import uuid
import hashlib
import json
from datetime import datetime
from openpyxl import Workbook
from collections import defaultdict
from openpyxl.styles import Font
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import sys
from libroIVA import create_secondary_window,centrar_ventana

SECRET_KEY = "348fjeqposaui%//FpouLSDFGkc7513i-*/TRWERa"
APP_NAME = "LJSON_V3"

def get_machine_id():
    return str(uuid.getnode())

def generar_licencia(machine_id):
    data = machine_id + SECRET_KEY
    return hashlib.sha256(data.encode()).hexdigest()

def ruta_licencia():
    base = os.getenv("APPDATA")
    carpeta = os.path.join(base, APP_NAME)
    os.makedirs(carpeta, exist_ok=True)
    return os.path.join(carpeta, "license.key")

def validar_licencia():
    ruta = ruta_licencia()

    if not os.path.exists(ruta):
        return False

    with open(ruta, "r") as f:
        licencia_guardada = f.read().strip()

    return licencia_guardada == generar_licencia(get_machine_id())

def activar_gui():

    machine_id = get_machine_id()

    root = tk.Tk()
    root.title("Activación LJSON V3")
    root.geometry("420x220")
    root.resizable(False, False)

    tk.Label(root, text="ID DEL EQUIPO", font=("Arial", 10, "bold")).pack(pady=(10, 2))
    id_entry = tk.Entry(root, width=50)
    id_entry.insert(0, machine_id)
    id_entry.config(state="readonly")
    id_entry.pack(pady=5)

    tk.Label(root, text="CLAVE DE ACTIVACIÓN").pack(pady=(10, 2))
    key_entry = tk.Entry(root, width=50)
    key_entry.pack(pady=5)

    def activar():
        clave = key_entry.get().strip()
        if clave == generar_licencia(machine_id):
            with open(ruta_licencia(), "w") as f:
                f.write(clave)
            messagebox.showinfo("Éxito", "Activación exitosa")
            root.destroy()
        else:
            messagebox.showerror("Error", "Clave inválida")
            root.destroy()
            sys.exit(1)

    tk.Button(root, text="ACTIVAR", command=activar).pack(pady=15)
    root.mainloop()

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

def cargar_datos_archivo(ruta):
    codificaciones = ['utf-8-sig', 'latin-1', 'cp1252']
    for codificacion in codificaciones:
        try:
            with open(ruta, 'r', encoding=codificacion) as archivo:
                contenido = archivo.read()
                if not contenido.strip():
                    print(f"El archivo {ruta} está vacío.")
                    return None
                electronicas_coasso = json.loads(contenido)
                return electronicas_coasso
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Error decodificando JSON en el archivo {ruta} con codificación {codificacion}: {e}")
        except Exception as e:
            print(f"Error inesperado en el archivo {ruta} con codificación {codificacion}: {e}")
    return None

def procesar_carpeta(carpeta, salida, nombre_archivo_salida, tree, tipo_documento):
    var_creciente = variables_crecientes()
    archivos_no_json = []
    if tipo_documento == "Compras":
        # Nuevo Libro de Trabajo
        wb = Workbook()
        # Creando hojas
        ws_resumen = wb.active
        ws_resumen.title = "Resumen de DTEs"
        #ws_retencion = wb.create_sheet(title="Comprobantes de Retención")
        ws_percepcion = wb.create_sheet(title="Percepción")
        ws_errores = wb.create_sheet(title="Errores")

        # Títulos de las hojas
        encabezados_resumen = ["Fecha", "Clase de Documento", "Tipo de DTE", "Num Doc", "NRC", "Nombre de Proveedor",
                               "Exenta/NoSuj", "IntE/S", "ImpE/S", "Internas Gravadas", "Internaciones Gravadas",
                               "Importaciones Gravadas de Bienes", "Importaciones Gravadas de Servicio",
                               "Credito Fiscal", "Monto Total", "Fovial", "Cotrans", "Percibido 1%", "Num_Anexo"]
        #encabezados_retencion = ["NIT", "Fecha", "Tipo DTE", "Sello Recepcion", "CodGeneracion", "Monto Sujeto", "Monto de Retención","DUI", "Num_Anexo"]
        encabezados_percepcion = ["NIT", "Fecha", "Tipo DTE", "Sello Recepcion", "CodGeneracion", "Monto Sujeto", "Monto de Percepción",
                                 "DUI", "Num_Anexo"]
        encabezados_errores = ["Nombre del Archivo", "Error"]

        ws_resumen.append(encabezados_resumen)
        #ws_retencion.append(encabezados_retencion)
        ws_percepcion.append(encabezados_percepcion)
        ws_errores.append(encabezados_errores)

        for cell in ws_resumen[1]:
            cell.font = Font(bold=True)

        #for cell in ws_retencion[1]:
        #    cell.font = Font(bold=True)

        for cell in ws_percepcion[1]:
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

                # Detectar si los datos están en la raíz o en un nivel anidado
                if 'identificacion' in datos and 'emisor' in datos and 'resumen' in datos:
                    contenido = datos
                else:
                    contenido = None
                    for valor in datos.values():
                        if isinstance(valor,
                                      dict) and 'identificacion' in valor and 'emisor' in valor and 'resumen' in valor:
                            contenido = valor
                            break

                if not contenido:
                    print(f"⚠️ Estructura no válida en el archivo {nombre_archivo}")
                    continue

                clase_doc = "4"
                valores = inicializar_valores()

                emisor = contenido['emisor']
                identificacion = contenido['identificacion']
                resumen = contenido['resumen']

                fecEmi = identificacion.get('fecEmi')
                tipoDte = identificacion.get('tipoDte')
                sello_recibido = obtener_sello_recibido(contenido)

                codigoGeneracion = identificacion.get('codigoGeneracion')
                nrc = emisor.get('nrc')
                nombre_emisor = emisor.get('nombre')
                nit = emisor.get('nit')

                if fecEmi:
                    date_obj = datetime.strptime(fecEmi, '%Y-%m-%d')
                    formatted_date = date_obj.strftime('%d/%m/%Y')
                else:
                    formatted_date = ""

                if tipoDte in ["03", "05", "06", "01"]:
                    valores['total_exenta'] = resumen.get('totalExenta', 0.00)-resumen.get('descuExenta', 0.00)
                    valores['total_no_sujeta'] = resumen.get('totalNoSuj', 0.00)-resumen.get('descuNoSuj', 0.00)
                    valores['total_gravada'] = resumen.get('totalGravada', 0.00)-resumen.get('descuGravada', 0.00)
                    valores['monto_total_operacion'] = resumen.get('montoTotalOperacion', 0.00)
                    valores['iva_perci1'] = resumen.get('ivaPerci1', 0.00)

                    if 'tributos' in resumen and resumen['tributos'] is not None:
                        for tributo in resumen['tributos']:
                            if tributo.get('codigo') == 'D1':
                                valores['fovial'] = tributo.get('valor', 0.00)
                                var_creciente['sum_fovial'] += valores['fovial']
                            if tributo.get('codigo') == 'C8':
                                valores['cotrans'] = tributo.get('valor', 0.00)
                                var_creciente['sum_cotrans'] += valores['cotrans']
                            if tributo.get('codigo') == '20':
                                valores['iva'] = tributo.get('valor', 0.00)
                                var_creciente['sum_iva'] += valores['iva']

                    if tipoDte == "03":
                        var_creciente['sum_ccf'] += 1
                        var_creciente['sum_grav'] += valores['total_gravada']
                        var_creciente['sum_total'] += valores['monto_total_operacion']
                    if tipoDte == "05":
                        var_creciente['sum_ntccredito'] += 1
                        var_creciente['sum_grav'] -= valores['total_gravada']
                        var_creciente['sum_total'] -= valores['monto_total_operacion']
                    if tipoDte == "06":
                        var_creciente['sum_ntcdebito'] += 1
                        var_creciente['sum_grav'] += valores['total_gravada']
                        var_creciente['sum_total'] += valores['monto_total_operacion']
                    if tipoDte == "01":
                        var_creciente['sum_factura'] += 1                      

                    # Agregar datos al resumen en Excel
                    fila_resumen = [
                        formatted_date, clase_doc, tipoDte, codigoGeneracion, nrc, nombre_emisor,
                        valores['total_exenta'] + valores['total_no_sujeta'], 0, 0, valores['total_gravada'], 0,
                        0, 0, valores['iva'], valores['monto_total_operacion'], valores['fovial'],
                        valores['cotrans'], valores['iva_perci1'], 3
                    ]
                    ws_resumen.append(fila_resumen)

                var_creciente['sum_docjson'] += 1

                if valores['iva_perci1'] > 0:
                    fila_percepcion = [
                        nit, formatted_date, tipoDte, sello_recibido, codigoGeneracion,
                        valores['total_gravada'], valores['iva_perci1'], "", 8
                    ]
                    ws_percepcion.append(fila_percepcion)

            else:
                archivos_no_json.append((nombre_archivo, "No es un archivo JSON"))

        # Actualizar Treeview
        doc_types = [
            ("Facturas", var_creciente['sum_factura']),
            ("Notas de Crédito", var_creciente['sum_ntccredito']),
            ("Notas de Débito", var_creciente['sum_ntcdebito']),
            ("Comprobantes de Crédito Fiscal", var_creciente['sum_ccf']),
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

        messagebox.showinfo("Compras", "Los archivos han sido procesados exitosamente.")
    elif tipo_documento == "Comprobante de Retención":
        wb = Workbook()
        ws_retencion = wb.active
        ws_retencion.title = "Comprobantes de Retención"
        ws_errores = wb.create_sheet(title="Errores")

        encabezados_retencion = [
            "NIT", "Fecha", "Tipo DTE", "Sello Recepcion",
            "CodGeneracion", "Monto Sujeto", "Monto de Retención",
            "DUI", "Num_Anexo"
        ]
        encabezados_errores = ["Nombre del Archivo", "Error"]

        ws_retencion.append(encabezados_retencion)
        ws_errores.append(encabezados_errores)

        for cell in ws_retencion[1]:
            cell.font = Font(bold=True)

        for cell in ws_errores[1]:
            cell.font = Font(bold=True)

        for nombre_archivo in os.listdir(carpeta):
            ruta_archivo = os.path.join(carpeta, nombre_archivo)

            if not nombre_archivo.lower().endswith(".json"):
                archivos_no_json.append((nombre_archivo, "No es un archivo JSON"))
                continue

            datos = cargar_datos_archivo(ruta_archivo)
            if not datos:
                archivos_no_json.append((nombre_archivo, "JSON inválido"))
                continue

            contenido = datos if 'identificacion' in datos else None
            if not contenido:
                archivos_no_json.append((nombre_archivo, "Estructura inválida"))
                continue

            identificacion = contenido['identificacion']
            resumen = contenido['resumen']
            emisor = contenido['emisor']

            if identificacion.get('tipoDte') != "07":
                continue

            fecEmi = identificacion.get('fecEmi')
            formatted_date = datetime.strptime(fecEmi, '%Y-%m-%d').strftime('%d/%m/%Y') if fecEmi else ""

            fila = [
                emisor.get('nit'),
                formatted_date,
                "07",
                obtener_sello_recibido(contenido),
                identificacion.get('codigoGeneracion'),
                resumen.get('totalSujetoRetencion', 0.00),
                resumen.get('totalIVAretenido', 0.00),
                "",
                7
            ]
            ws_retencion.append(fila)
            var_creciente['sum_comproretencion'] += 1
            var_creciente['sum_docjson'] += 1
        # Treeview
        tree.delete(*tree.get_children())
        tree.insert("", tk.END, values=("Comprobantes de Retención", var_creciente['sum_comproretencion']))
        tree.insert("", tk.END, values=("Documentos Procesados", var_creciente['sum_docjson']))

        wb.save(os.path.join(salida, f"{nombre_archivo_salida}.xlsx"))

        if archivos_no_json:
            for e in archivos_no_json:
                ws_errores.append(e)
            wb.save(os.path.join(salida, f"{nombre_archivo_salida}_errores.xlsx"))

        messagebox.showinfo("Retención", "Comprobantes de retención procesados correctamente.")
    elif tipo_documento == "Contribuyentes":
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
                if datos and 'identificacion' in datos and 'receptor' in datos and 'resumen' in datos:
                    clase_doc = "4" #B
                    valores = inicializar_valores()
                    receptor = datos['receptor']
                    identificacion = datos['identificacion']
                    resumen = datos['resumen']
                    fecEmi = identificacion.get('fecEmi') #A
                    tipoDte = identificacion.get('tipoDte') #C
                    num_control = identificacion.get('numeroControl') #D

                    sello_recibido = obtener_sello_recibido(datos)

                    codigoGeneracion = identificacion.get('codigoGeneracion')
                    nrc = receptor.get('nrc')
                    nombre_receptor = receptor.get('nombre')

                    if fecEmi:
                        date_obj = datetime.strptime(fecEmi, '%Y-%m-%d')
                        formatted_date = date_obj.strftime('%d/%m/%Y')
                    else:
                        formatted_date = ""

                    if tipoDte in ["03", "05", "06",]:
                        total_exenta = resumen.get('totalExenta', 0.00)
                        descu_exenta = resumen.get('descuExenta', 0.00)

                        total_no_suj = resumen.get('totalNoGravado', 0.00)
                        descu_no_suj = resumen.get('descuNoSuj', 0.00)

                        total_gravada = resumen.get('totalGravada', 0.00)
                        descu_gravada = resumen.get('descuGravada', 0.00)
                        if tipoDte == "05":
                            monto_total = resumen.get('montoTotalOperacion', 0.00)
                        else:
                            monto_total = resumen.get('totalPagar', 0.00)

                        valores['total_exenta'] = total_exenta - descu_exenta
                        valores['total_no_sujeta'] = total_no_suj - descu_no_suj
                        valores['sub_total_ventas'] = total_gravada - descu_gravada
                        valores['monto_total_operacion'] = monto_total
                        valores['iva_perci1'] = resumen.get('ivaPerci1', 0.00)
                        if 'tributos' in resumen and resumen['tributos'] is not None:
                            for tributo in resumen['tributos']:
                                if tributo.get('codigo') == 'D1':
                                    valores['fovial'] = tributo.get('valor', 0.00)
                                    var_creciente['sum_fovial'] += valores['fovial']
                                if tributo.get('codigo') == 'C8':
                                    valores['cotrans'] = tributo.get('valor', 0.00)
                                    var_creciente['sum_cotrans'] += valores['cotrans']
                                if tributo.get('codigo') == '20':
                                    valores['iva'] = tributo.get('valor', 0.00)
                                    var_creciente['sum_iva'] += valores['iva']

                        if tipoDte == "03":
                            var_creciente['sum_ccf'] += 1
                            var_creciente['sum_grav'] += valores['sub_total_ventas']
                            var_creciente['sum_total'] += valores['monto_total_operacion']
                        if tipoDte == "05":
                            var_creciente['sum_ntccredito'] += 1
                            var_creciente['sum_grav'] -= valores['sub_total_ventas']
                            var_creciente['sum_total'] -= valores['monto_total_operacion']
                        if tipoDte == "06":
                            var_creciente['sum_ntcdebito'] += 1
                            var_creciente['sum_grav'] += valores['sub_total_ventas']
                            var_creciente['sum_total'] += valores['monto_total_operacion']
                        if tipoDte == "01":
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

    elif tipo_documento == "Consumidor Final":
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
                if datos and 'identificacion' in datos and 'receptor' in datos and 'resumen' in datos:
                    identificacion = datos['identificacion']
                    resumen = datos['resumen']
                    tipoDte = identificacion.get('tipoDte')
                    fecEmi = identificacion.get('fecEmi')
                    horEmi= identificacion.get('horEmi')
                    codigoGeneracion = identificacion.get('codigoGeneracion')

                    if fecEmi and tipoDte in ["01", "11"]:
                        formatted_date = datetime.strptime(fecEmi, '%Y-%m-%d').strftime('%d/%m/%Y')
                        clave = (formatted_date, tipoDte)
                        total_exenta = resumen.get('totalExenta', 0.00)
                        descu_exenta = resumen.get('descuExenta', 0.00)

                        total_no_suj = resumen.get('totalNoSuj', 0.00)
                        descu_no_suj = resumen.get('descuNoSuj', 0.00)

                        total_no_gravado = resumen.get('totalNoGravado', 0.00)
                        # normalmente no lleva descuento, pero lo dejo separado por claridad

                        total_gravada = resumen.get('totalGravada', 0.00)
                        descu_gravada = resumen.get('descuGravada', 0.00)
                        agrupados[clave].append({
                            "codigo": codigoGeneracion,
                            "horEmi": horEmi,
                            "total_exenta": total_exenta - descu_exenta,
                            "total_no_sujeta": total_no_suj - descu_no_suj,
                            "total_no_gravado": total_no_gravado,
                            "total_gravada": total_gravada - descu_gravada,
                            "monto_total_operacion": resumen.get('montoTotalOperacion', 0.00),
                            "iva": resumen.get('ivaRetenido', 0.00)
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

            if tipoDte == "01":
                var_creciente['sum_factura'] += len(documentos)
                fila_resumen += [
                    total_exenta, 0, total_no_sujeta,
                    total_gravada, 0, 0, 0, 0, 0,
                    monto_total_operacion, "", "", 2
                ]
            elif tipoDte == "11":
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

    elif tipo_documento == "Sujeto Excluido":
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
                if datos and 'identificacion' in datos and 'sujetoExcluido' in datos and 'resumen' in datos:
                    clase_doc = "4" #B

                    valores = inicializar_valores()
                    sujetoExcluido = datos['sujetoExcluido']
                    identificacion = datos['identificacion']
                    resumen = datos['resumen']
                    fecEmi = identificacion.get('fecEmi') #A
                    tipoDte = identificacion.get('tipoDte') #C
                    numDocumento = sujetoExcluido.get('numDocumento')
                    nombre = sujetoExcluido.get('nombre')
                    sello_recibido = obtener_sello_recibido(datos)
                    codigoGeneracion = identificacion.get('codigoGeneracion')

                    if numDocumento and isinstance(numDocumento, str):
                        longitud = len(numDocumento)
                        if 9 <= longitud <= 10:
                            tipoDocumento = 2
                        elif 14 <= longitud <= 17:
                            tipoDocumento = 1
                        elif longitud > 17:
                            tipoDocumento = 3
                        else:
                            tipoDocumento = None  # o algún valor por defecto si no cae en ningún rango
                    else:
                        tipoDocumento = None  # si numDocumento no es string o es None

                    if fecEmi:
                        date_obj = datetime.strptime(fecEmi, '%Y-%m-%d')
                        formatted_date = date_obj.strftime('%d/%m/%Y')
                    else:
                        formatted_date = ""

                    if tipoDte in ["14"]:
                        var_creciente['sum_sujetoexcluido'] += 1
                        valores['total_Compra']=resumen.get('totalCompra',0.00)

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
    elif tipo_documento == "Documentos Anulados":
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

                # Detectar si los datos están en la raíz o en un nivel anidado
                if 'identificacion' in datos and 'emisor' in datos and 'resumen' in datos:
                    contenido = datos
                else:
                    contenido = None
                    for valor in datos.values():
                        if isinstance(valor,
                                      dict) and 'identificacion' in valor and 'emisor' in valor and 'resumen' in valor:
                            contenido = valor
                            break

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

                if tipoDte in ["03", "05", "11", "01", "14" ]:
                    if tipoDte == "01":
                        var_creciente['sum_factura'] += 1
                    if tipoDte == "03":
                        var_creciente['sum_ccf'] +=1
                    if tipoDte == "05":
                        var_creciente['sum_ntccredito'] += 1
                    if tipoDte == "11":
                        var_creciente['sum_export'] += 1
                    if tipoDte == "14":
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
    elif tipo_documento not in ["Compras", "Contribuyentes", "Consumidor Final","Sujeto Excluido"]:
        messagebox.showerror("Error", "Debe seleccionar un tipo de documento válido.")
        return

# Función para obtener el sello recibido
def obtener_sello_recibido(datos):
    campos_posibles = ['SelloRecepcion','selloRecepcion',"response", 'responseMH', 'ResponseMH', 'respuestaHacienda']
    for campo in campos_posibles:
        if campo in datos:
            valor = datos[campo]
            if isinstance(valor, dict):
                return valor.get('selloRecibido')
            elif isinstance(valor, str):
                # Si el campo directamente es un string que representa el sello
                return valor
    # También puede estar directo en la raíz
    valor = datos.get('selloRecibido')
    if isinstance(valor, str):
        return valor
    return None

def iniciar_procesamiento(carpeta_entrada, carpeta_salida, nombre_archivo, tipo_documento, tree):
    carpeta = carpeta_entrada.get()
    salida = carpeta_salida.get()
    nombre_archivo_salida = nombre_archivo.get()
    tipo_documento_seleccionado = tipo_documento.get()  # Obtener el valor seleccionado del ComboBox
    print(tipo_documento_seleccionado)

    if carpeta and salida and nombre_archivo_salida and tipo_documento_seleccionado:
        # Llamar a procesar_carpeta con el tipo de documento seleccionado
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
    centrar_ventana(root,600,400)
    
    
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
        values=["Compras", "Contribuyentes", "Consumidor Final","Comprobante de Retención", "Sujeto Excluido","Documentos Anulados"],
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

    # Botón para iniciar el procesamiento
    tk.Button(
        frame_entradas,
        text="Iniciar procesamiento",
        command=lambda: iniciar_procesamiento(carpeta_entrada, carpeta_salida, nombre_archivo, tipo_documento, tree)
    ).grid(row=4, column=1, pady=10)

    # Botón para limpiar los campos
    tk.Button(frame_entradas, text="Limpiar", command=lambda: limpiar_campos(tipo_documento, carpeta_entrada, carpeta_salida, nombre_archivo, tree)).grid(row=4, column=2, pady=10)

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
