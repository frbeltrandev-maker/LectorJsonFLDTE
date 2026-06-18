import json
import os
import sys
import tkinter as tk
from datetime import datetime
from collections import defaultdict
from tkinter import ttk, filedialog, messagebox
import openpyxl
from openpyxl.styles import Alignment, Font,Border, Side
from openpyxl.utils import get_column_letter
from clientes_store import cargar_clientes, guardar_cliente,editar_cliente, eliminar_cliente

def variables_crecientes():
    return {
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
        'total_no_sujeta': 0,
        'sub_total_ventas': 0,
        'iva': 0,
        'monto_total_operacion': 0,
        'iva_perci1': 0,
        'total_sujeto_retencion': 0,
        'total_iva_retenido': 0
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
def format_nit_input(*args):
    text = nit_var.get().replace('-', '')  # Eliminar guiones
    if len(text) > 3:
        text = text[:4] + '-' + text[4:]
    if len(text) > 10:
        text = text[:11] + '-' + text[11:]
    if len(text) > 14:
        text = text[:15] + '-' + text[15:]
    text = text[:17]
    nit_var.set(text)
    entry_index = len(text)
    if len(text) > 5:
        entry_index += 1
    if len(text) > 12:
        entry_index += 1
    if len(text) > 17:
        entry_index += 1
    nit_entry.icursor(entry_index)

def format_dui_input(*args):
    text = dui_var.get().replace('-', '')  # Eliminar guiones
    if len(text) > 7:
        text = text[:8] + '-' + text[8:]
    text = text[:10]
    dui_var.set(text)
    entry_index = len(text)
    if len(text) > 10:
        entry_index += 1
    dui_entry.icursor(entry_index)

def validar_nombre_contri(*args):
    return bool(nombre_contri_var.get().strip())


def seleccionar_ruta_entrada():
    ruta = filedialog.askdirectory()
    if ruta:  # Verificar que el usuario no haya cancelado el diálogo
        ruta_entrada_var.set(ruta)
    else:
        messagebox.showwarning("Selección de carpeta", "No se ha seleccionado ninguna carpeta.")
def seleccionar_ruta_salida():
    ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    ruta_salida_var.set(ruta)

def centrar_ventana(win, width, height):
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()

    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)

    win.geometry(f"{width}x{height}+{x}+{y}")
    win.deiconify()  

   

def ventana_cliente(parent, callback_actualizar, cliente_seleccionado=None):
    """
    Ventana para agregar o editar clientes.
    - cliente_seleccionado: si se pasa un diccionario, se carga para edición.
    """
    win = tk.Toplevel(parent)
    win.title("Cliente")
    win.resizable(False, False)
    win.update_idletasks()
    centrar_ventana(win,325,175)

    # Labels
    tk.Label(win, text="Nombre:", font='Arial 10 bold').grid(row=0, column=0, padx=10, pady=5, sticky="e")
    tk.Label(win, text="NRC:", font='Arial 10 bold').grid(row=1, column=0, padx=10, pady=5, sticky="e")
    tk.Label(win, text="NIT:", font='Arial 10 bold').grid(row=2, column=0, padx=10, pady=5, sticky="e")
    tk.Label(win, text="DUI:", font='Arial 10 bold').grid(row=3, column=0, padx=10, pady=5, sticky="e")

    # Variables
    nombre_var = tk.StringVar(value=cliente_seleccionado["nombre"] if cliente_seleccionado else "")
    nrc_var = tk.StringVar(value=cliente_seleccionado.get("nrc", "") if cliente_seleccionado else "")
    nit_var = tk.StringVar(value=cliente_seleccionado.get("nit", "") if cliente_seleccionado else "")
    dui_var = tk.StringVar(value=cliente_seleccionado.get("dui", "") if cliente_seleccionado else "")

    # Entrys
    nombre_entry = tk.Entry(win, width=30, textvariable=nombre_var)
    nrc_entry = tk.Entry(win, width=30, textvariable=nrc_var)
    nit_entry = tk.Entry(win, width=30, textvariable=nit_var)
    dui_entry = tk.Entry(win, width=30, textvariable=dui_var)

    nombre_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
    nrc_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
    nit_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
    dui_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

    # --- Formateo NIT/DUI ---
    def format_nit(*args):
        text = nit_var.get().replace('-', '')  # Eliminar guiones
        if len(text) > 3:
            text = text[:4] + '-' + text[4:]
        if len(text) > 10:
            text = text[:11] + '-' + text[11:]
        if len(text) > 14:
            text = text[:15] + '-' + text[15:]
        text = text[:17]
        nit_var.set(text)
        entry_index = len(text)
        if len(text) > 5:
            entry_index += 1
        if len(text) > 12:
            entry_index += 1
        if len(text) > 17:
            entry_index += 1
        nit_entry.icursor(entry_index)    

    def format_dui(*args):
        text = dui_var.get().replace('-', '')
        if len(text) > 7:
            text = text[:8] + '-' + text[8:]
        text = text[:10]
        dui_var.set(text)

    nit_var.trace_add("write", format_nit)
    dui_var.trace_add("write", format_dui)

    # --- Guardar ---
    def guardar():
        if not nombre_var.get().strip():
            messagebox.showwarning("Campo obligatorio", "El nombre es obligatorio")
            return

        cliente_data = {
            "nombre": nombre_var.get(),
            "nrc": nrc_var.get(),
            "nit": nit_var.get(),
            "dui": dui_var.get()
        }

        if cliente_seleccionado:
            editar_cliente(cliente_seleccionado["nombre"], cliente_data)
        else:
            guardar_cliente(cliente_data)

        callback_actualizar()
        win.destroy()

    tk.Button(win, text="Guardar", command=guardar).grid(row=4, column=0, columnspan=2, pady=15)


def crear_libro_excel(encabezado_1, encabezado_2, titulo_libro, nombre_contribuyente, nrc, nit, dui, periodo, anio,filas_datos,totales,nombre_firmador,salida,resumen_operaciones):
    wb = openpyxl.Workbook()
    ws = wb.active

    bold_font = Font(bold=True, size=11)
    centered_alignment = Alignment(horizontal="center", vertical="center")
    right_alignment = Alignment(horizontal="left")
    thin_border = Border(top=Side(style='thin'))
    borde_cuadro = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))

    for col in range(1, len(encabezado_1) + 1):
        ws.cell(row=11, column=col).border = thin_border
        ws.cell(row=13, column=col).border = thin_border

    # Encabezados
    for col, header in enumerate(encabezado_1, start=1):
        cell = ws.cell(row=11, column=col, value=header)
        cell.font = bold_font
        cell.alignment = centered_alignment

    for col, header in enumerate(encabezado_2, start=1):
        cell = ws.cell(row=12, column=col, value=header)
        cell.font = bold_font
        cell.alignment = centered_alignment

    # Combinar celdas comunes
    ws.merge_cells('A11:A12')
    ws.merge_cells('B11:B12')


    # Ajustes dinámicos según tipo de libro
    if titulo_libro == 'LIBRO DE COMPRAS':
        ws.merge_cells('F11:G11')
        ws.merge_cells('H11:I11')
        ws.merge_cells('D11:D12')
        ws.merge_cells('E11:E12')
    elif titulo_libro == 'LIBRO DE VENTAS A CONTRIBUYENTES':
        ws.merge_cells('F11:H11')
        ws.merge_cells('I11:K11')
        ws.merge_cells('D11:D12')
        ws.merge_cells('E11:E12')
    else:
        ws.merge_cells('E11:E12')
        ws.merge_cells('F11:I11')

    # Ajustar ancho de columnas
    column_lengths = {}
    for row in [11, 12]:
        for col_idx in range(1, len(encabezado_1) + 1):
            col_letter = get_column_letter(col_idx)
            cell = ws.cell(row=row, column=col_idx)
            value_length = len(str(cell.value)) if cell.value else 0
            column_lengths[col_letter] = max(column_lengths.get(col_letter, 0), value_length)

    for col_letter, length in column_lengths.items():
        ws.column_dimensions[col_letter].width = length + 2

    if titulo_libro == 'LIBRO DE VENTAS A CONSUMIDOR':
        ws.merge_cells('A2:K2')
        ws['A2'] = f'CONTRIBUYENTE: {nombre_contribuyente}'
        ws['A2'].alignment = centered_alignment
        ws['A2'].font = bold_font

        ws['C3'] = f'NRC: {nrc}'
        ws['C3'].alignment = right_alignment
        ws['C3'].font = bold_font

        ws['D3'] = f'NIT: {nit}'
        ws['D3'].alignment = right_alignment
        ws['D3'].font = bold_font

        ws['F3'] = f'DUI: {dui}'
        ws['F3'].alignment = right_alignment
        ws['F3'].font = bold_font

        ws.merge_cells('A5:K5')
        ws['A5'] = titulo_libro
        ws['A5'].alignment = centered_alignment
        ws['A5'].font = bold_font

        ws.merge_cells('A6:K6')
        ws['A6'] = '(EXPRESADO EN DOLARES DE LOS ESTADOS UNIDOS DE AMERICA)'
        ws['A6'].alignment = centered_alignment

        ws['A8'] = f'PERIODO TRIBUTARIO: {periodo.upper()}'
        ws['A8'].alignment = right_alignment
        ws['A8'].font = bold_font

        ws['F8'] = f'AÑO: {anio}'
        ws['F8'].alignment = centered_alignment
        ws['F8'].font = bold_font
    else:
        ws.merge_cells('A2:N2')
        ws['A2'] = f'CONTRIBUYENTE: {nombre_contribuyente}'
        ws['A2'].alignment = centered_alignment
        ws['A2'].font = bold_font

        ws['E3'] = f'NRC: {nrc}'
        ws['E3'].alignment = right_alignment
        ws['E3'].font = bold_font

        ws['G3'] = f'NIT: {nit}'
        ws['G3'].alignment = right_alignment
        ws['G3'].font = bold_font

        ws['J3'] = f'DUI: {dui}'
        ws['J3'].alignment = right_alignment
        ws['J3'].font = bold_font

        ws.merge_cells('A5:N5')
        ws['A5'] = titulo_libro
        ws['A5'].alignment = centered_alignment
        ws['A5'].font = bold_font

        ws.merge_cells('A6:N6')
        ws['A6'] = '(EXPRESADO EN DOLARES DE LOS ESTADOS UNIDOS DE AMERICA)'
        ws['A6'].alignment = centered_alignment

        ws['A8'] = f'PERIODO TRIBUTARIO: {periodo.upper()}'
        ws['A8'].alignment = right_alignment
        ws['A8'].font = bold_font

        ws['F8'] = f'AÑO: {anio}'
        ws['F8'].alignment = centered_alignment
        ws['F8'].font = bold_font

    # === REGISTROS DE DATOS ===
    fila_actual = 13
    filas_datos_ordenadas = sorted(filas_datos, key=lambda x: x[0])  # x[0] es la fecha

    for num_registro, (_, fila) in enumerate(filas_datos_ordenadas, start=1):
        ws.cell(row=fila_actual, column=1, value=num_registro)
        for col_idx, valor in enumerate(fila[1:], start=2):
            celda = ws.cell(row=fila_actual, column=col_idx, value=valor)
            if 6 <= col_idx <= 14:
                celda.alignment = centered_alignment
                celda.number_format = '0.00'
            else:
                celda.alignment = centered_alignment
        fila_actual += 1

    # === FILA DE TOTALES ===
    fila_totales = fila_actual + 1
    ws.cell(row=fila_totales, column=5, value="Totales:").font = bold_font
    for idx, valor in enumerate(totales, start=6):
        ws.cell(row=fila_totales, column=idx, value=valor)

    borde_fila_total = Border(
        top=Side(border_style='thin', color='000000'),
        bottom=Side(border_style='thin', color='000000')
    )

    if titulo_libro == 'LIBRO DE VENTAS A CONSUMIDOR':
        for col in range(1, 12):
            celda = ws.cell(row=fila_totales, column=col)
            celda.font = bold_font
            celda.alignment = centered_alignment
            celda.border = borde_fila_total
            celda.number_format = '0.00'
    else:
        for col in range(1, 15):
            celda = ws.cell(row=fila_totales, column=col)
            celda.font = bold_font
            celda.alignment = centered_alignment
            celda.border = borde_fila_total
            celda.number_format = '0.00'

    
    fila_resumen = fila_totales + 3 

    if resumen_operaciones and titulo_libro.startswith('LIBRO DE VENTAS'):

        fila_resumen = fila_totales + 3

        # Definir columnas según tipo de libro
        if titulo_libro == 'LIBRO DE VENTAS A CONTRIBUYENTES':
            col_inicio = 10   # B
        else:  # CONSUMIDOR FINAL
            col_inicio = 7   # F

        col_titulo = get_column_letter(col_inicio)
        col_ex = col_inicio
        col_gr = col_inicio + 1
        col_de = col_inicio + 2
        col_exp = col_inicio + 3
        col_tot = col_inicio + 4
        col_titulo_inicio = col_inicio - 1
        # Título
        ws.merge_cells(
            f'{get_column_letter(col_titulo_inicio)}{fila_resumen}:'
            f'{get_column_letter(col_tot)}{fila_resumen}'
        )
        ws[f'{get_column_letter(col_titulo_inicio)}{fila_resumen}'] = 'Resumen de Operaciones'
        ws[f'{get_column_letter(col_titulo_inicio)}{fila_resumen}'].font = bold_font
        ws[f'{get_column_letter(col_titulo_inicio)}{fila_resumen}'].alignment = centered_alignment

        headers = [
            "Ventas\nExentas",
            "Ventas\nGravadas",
            "Débito\nFiscal",
            "Ventas\nExportación",
            "Ventas\nTotales"
        ]

        for i, h in enumerate(headers):
            cell = ws.cell(row=fila_resumen + 1, column=col_inicio + i, value=h)
            cell.font = bold_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        # Contribuyentes
        ws.cell(row=fila_resumen + 2, column=col_inicio - 1, value='Contribuyentes').font = bold_font
        c = resumen_operaciones["contribuyentes"]

        ws.cell(row=fila_resumen + 2, column=col_ex, value=c["exentas"])
        ws.cell(row=fila_resumen + 2, column=col_gr, value=c["gravadas"])
        ws.cell(row=fila_resumen + 2, column=col_de, value=c["debito"])
        ws.cell(row=fila_resumen + 2, column=col_exp, value=c["exportacion"])
        ws.cell(row=fila_resumen + 2, column=col_tot, value=c["total"])

        # Consumidor Final
        ws.cell(row=fila_resumen + 3, column=col_inicio - 1, value='Consumidor Final').font = bold_font
        cf = resumen_operaciones["consumidor_final"]

        ws.cell(row=fila_resumen + 3, column=col_ex, value=cf["exentas"])
        ws.cell(row=fila_resumen + 3, column=col_gr, value=cf["gravadas"])
        ws.cell(row=fila_resumen + 3, column=col_de, value=cf["debito"])
        ws.cell(row=fila_resumen + 3, column=col_exp, value=cf["exportacion"])
        ws.cell(row=fila_resumen + 3, column=col_tot, value=cf["total"])

        for fila in (fila_resumen + 2, fila_resumen + 3):  # Contribuyentes y Consumidor Final
            for col in (col_ex, col_gr, col_de, col_exp, col_tot):
                celda = ws.cell(row=fila, column=col)
                celda.number_format = '0.00'
                celda.alignment = centered_alignment

        fila_inicio = fila_resumen
        fila_fin = fila_resumen + 3
        col_inicio_borde = col_inicio - 1
        col_fin_borde = col_tot

        for fila in range(fila_inicio, fila_fin + 1):
            for col in range(col_inicio_borde, col_fin_borde + 1):
                ws.cell(row=fila, column=col).border = borde_cuadro

        fila_firma = fila_resumen + 8
    else:
        fila_firma = fila_totales + 3
    
    # === FIRMA CONTADOR ===
    ws[f'A{fila_firma}'] = 'Firma del Contador: '
    ws[f'A{fila_firma}'].alignment = right_alignment
    ws[f'A{fila_firma}'].font = bold_font

    ws[f'G{fila_firma}'] = f'Nombre de contador:  {nombre_firmador}'
    ws[f'G{fila_firma}'].alignment = right_alignment
    ws[f'G{fila_firma}'].font = bold_font

    # Configurar impresión
    ws.print_title_rows = '1:12'

    # === GUARDAR ARCHIVO ===
    wb.save(salida)



def generar_libro_excel(carpeta,salida,tipo_libro):
    var_creciente = variables_crecientes()
    if tipo_libro == "Libro de Compras":
        nombre_contribuyente = nombre_contri_var.get()
        nrce = nrc_entry.get()
        nit = nit_entry.get()
        dui = dui_entry.get()
        periodo = month_combobox.get()
        anio = year_combobox.get()
        nombre_firmador = nombre_firma_var.get()
        filas_datos = []

        encabezado_compras_1 = ["#", "Fecha", "No", "N.R.C.", "Nombre del Proveedor", "Compras Exentas", "",
                                "Compras Gravadas", "", "Fovial", "Credito", "Total", "Percepcion", "Sujetos"]
        encabezado_compras_2 = ["", "", "Documento", "", "", "Internas", "Import.", "Internas", "Import.", "Cotrans",
                                "Fiscal", "Compra", "IVA", "Excluidos"]

        sum_compras_exentas_internas = 0
        sum_compras_exentas_import=0
        sum_compras_gravadas_internas = 0
        sum_compras_gravadas_import=0
        sum_fovial_cotrans = 0
        sum_credito_fiscal = 0
        sum_total_compra = 0
        sum_percepcion_iva = 0
        sum_sujetos_excluidos=0

        totales = [
            sum_compras_exentas_internas,
            sum_compras_exentas_import,
            sum_compras_gravadas_internas,
            sum_compras_gravadas_import,
            sum_fovial_cotrans,
            sum_credito_fiscal,
            sum_total_compra,
            sum_percepcion_iva,
            sum_sujetos_excluidos
        ]

        for nombre_de_archivo in os.listdir(carpeta):
            ruta_archivo = os.path.join(carpeta, nombre_de_archivo)
            if os.path.isfile(ruta_archivo) and nombre_de_archivo.lower().endswith('.json'):
                archivo_json = cargar_datos_archivo(ruta_archivo)
                if not archivo_json:
                    continue

                # Detectar si los datos están en la raíz o en un nivel anidado
                if 'identificacion' in archivo_json and 'emisor' in archivo_json and 'resumen' in archivo_json:
                    contenido = archivo_json
                else:
                    contenido = None
                    for valor in archivo_json.values():
                        if isinstance(valor,
                                      dict) and 'identificacion' in valor and 'emisor' in valor and 'resumen' in valor:
                            contenido = valor
                            break

                if not contenido:
                    print(f"Estructura no válida en el archivo {nombre_de_archivo}")
                    continue

                valores = inicializar_valores()
                identificacion = contenido['identificacion']
                emisor = contenido['emisor']
                resumen = contenido['resumen']

                codigoGeneracion = identificacion.get('codigoGeneracion')
                nrc = emisor.get('nrc')
                nombre_emisor = emisor.get('nombre')
                tipoDte = identificacion.get('tipoDte')
                fecEmi = identificacion.get('fecEmi')

                if fecEmi:
                    date_obj = datetime.strptime(fecEmi, '%Y-%m-%d')
                    formatted_date = date_obj.strftime('%d/%m/%Y')
                else:
                    formatted_date = ""

                if tipoDte in ["03", "05", "06"]:
                    valores['total_exenta'] = resumen.get('totalExenta', 0.00) - resumen.get('descuExenta', 0.00)
                    valores['total_no_sujeta'] = resumen.get('totalNoSuj', 0.00) - resumen.get('descuNoSuj', 0.00)
                    valores['total_gravada'] = resumen.get('totalGravada', 0.00) - resumen.get('descuGravada', 0.00)
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

                    if tipoDte in ["03", "06"]:
                        var_creciente['sum_grav'] += valores['total_gravada']
                        var_creciente['sum_total'] += valores['monto_total_operacion']

                    if tipoDte == "05":  # Nota de crédito
                        valores = {key: -abs(value) for key, value in valores.items()}
                        var_creciente['sum_grav'] -= valores['total_gravada']
                        var_creciente['sum_total'] -= valores['monto_total_operacion']

                    fila = [
                        var_creciente['sum_docjson'] + 1,
                        formatted_date,
                        codigoGeneracion,
                        nrc,
                        nombre_emisor,
                        valores['total_exenta'] + valores['total_no_sujeta'] if valores['total_exenta'] + valores[
                            'total_no_sujeta'] != 0 else 0.00,
                        "0.00",
                        valores['total_gravada'] if valores['total_gravada'] != 0 else 0.00,
                        "0.00",
                        valores['fovial'] + valores['cotrans'],
                        valores['iva'],
                        valores['monto_total_operacion']+valores['iva_perci1'],
                        valores['iva_perci1'],
                        "0.00",
                    ]

                    filas_datos.append((date_obj, fila))

                    var_creciente['sum_docjson'] += 1
                    # Totales
                    totales[0] += valores.get('total_exenta', 0)
                    totales[1] += valores.get('total_exenta_import', 0)
                    totales[2] += valores.get('total_gravada', 0)
                    totales[3] += valores.get('sub_total_ventas_import', 0)
                    totales[4] += valores.get('fovial', 0) + valores.get('cotrans', 0)
                    totales[5] += valores.get('iva', 0)
                    totales[6] += valores.get('monto_total_operacion', 0) + valores.get('iva_perci1', 0)
                    totales[7] += valores.get('iva_perci1', 0)
                    totales[8] += valores.get('sujeto_excluido', 0)
           
        crear_libro_excel(
            encabezado_compras_1,
            encabezado_compras_2,
            'LIBRO DE COMPRAS',
            nombre_contribuyente,
            nrce,
            nit,
            dui,
            periodo,
            anio,
            filas_datos,
            totales,
            nombre_firmador,
            salida,
            resumen_operaciones=None
        )
        messagebox.showinfo("Éxito", "El libro de Compras se ha generado exitosamente.")
    elif tipo_libro == "Ventas a Contribuyentes":
        nombre_contribuyente = nombre_contri_var.get()
        nrce = nrc_entry.get()
        nit = nit_entry.get()
        dui = dui_entry.get()
        periodo = month_combobox.get()
        anio = year_combobox.get()
        nombre_firmador = nombre_firma_var.get()
        filas_datos = []

        # Encabezado de registros.
        encabezado_registro = ["#", "Fecha", "No", "N.R.C.", "Nombre del Proveedor", "Ventas Propias", "","",
                               "Ventas a Terceros", "","", "Total", "Percepcion", "Retencion"]
        encabezado_registro1 = ["", "", "Documento", "", "", "Exentas", "Gravadas", "Debito", "Exentas","Gravadas","Debito","Venta",
                                "IVA", "IVA"]

        sum_compras_exentas_internas = 0.0
        sum_compras_gravadas_internas = 0.0
        sum_credito_fiscal = 0.0
        sum_ventas_exentas_ter=0.0
        sum_ventas_gravadas_ter = 0.0
        sum_ventas_debito_ter = 0.0
        sum_total_compra = 0.0
        sum_percepcion_iva = 0.0
        sum_retencion_iva = 0.0

        totales = [
            sum_compras_exentas_internas,
            sum_compras_gravadas_internas,
            sum_credito_fiscal,
            sum_ventas_exentas_ter ,
            sum_ventas_gravadas_ter,
            sum_ventas_debito_ter,
            sum_total_compra,
            sum_percepcion_iva,
            sum_retencion_iva,
        ]

        resumen_operaciones = {
            "contribuyentes": {
                "exentas": None,
                "gravadas": None,
                "debito": None,
                "exportacion": None,
                "total": None
            },
            "consumidor_final": {
                "exentas": None,
                "gravadas": None,
                "debito": None,
                "exportacion": None,
                "total": None
            }
        }

        for nombre_de_archivo in os.listdir(carpeta):
            ruta_archivo = os.path.join(carpeta, nombre_de_archivo)
            if os.path.isfile(ruta_archivo) and nombre_de_archivo.lower().endswith('.json'):
                archivo_json = cargar_datos_archivo(ruta_archivo)
                if not archivo_json:
                    continue

                    # Detectar si los datos están en la raíz o en un nivel anidado
                if 'identificacion' in archivo_json and 'receptor' in archivo_json and 'resumen' in archivo_json:
                    contenido = archivo_json
                else:
                    contenido = None
                    for valor in archivo_json.values():
                        if isinstance(valor,
                                      dict) and 'identificacion' in valor and 'receptor' in valor and 'resumen' in valor:
                            contenido = valor
                            break

                if not contenido:
                    print(f"⚠️ Estructura no válida en el archivo {nombre_de_archivo}")
                    continue

                valores = inicializar_valores()
                identificacion = contenido['identificacion']
                receptor = contenido['receptor']
                resumen = contenido['resumen']

                codigoGeneracion = identificacion.get('codigoGeneracion')  # C13
                nrc = receptor.get('nrc')  # D13
                nombre_emisor = receptor.get('nombre')  # E13
                tipoDte = identificacion.get('tipoDte')
                fecEmi = identificacion.get('fecEmi')  # B13

                if fecEmi:
                    date_obj = datetime.strptime(fecEmi, '%Y-%m-%d')
                    formatted_date = date_obj.strftime('%d/%m/%Y')
                else:
                    formatted_date = ""

                if tipoDte in ["03", "05", "06"]:
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
                    valores['monto_total_operacion'] = monto_total # L13
                    valores['iva_perci1'] = resumen.get('ivaPerci1', 0.00)  # M13
                    valores['iva_ret1'] = resumen.get('ivaRete1', 0.00)  # M13

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

                    if tipoDte in ["03", "06"]:
                        var_creciente['sum_grav'] += valores['sub_total_ventas']
                        var_creciente['sum_total'] += valores['monto_total_operacion']

                    if tipoDte == "05":
                        valores = {key: -abs(value) for key, value in valores.items()}
                        var_creciente['sum_grav'] -= valores['sub_total_ventas']
                        var_creciente['sum_total'] -= valores['monto_total_operacion']

                    fila_datos = [
                        var_creciente['sum_docjson'] + 1,
                        formatted_date,
                        codigoGeneracion,
                        nrc,
                        nombre_emisor,
                        valores['total_exenta'] + valores['total_no_sujeta'],  # Exentas
                        valores['sub_total_ventas'],
                        valores['iva'],
                        "0.00",
                        "0.00",
                        "0.00",
                        valores['monto_total_operacion'],
                        valores['iva_perci1'],
                        valores['iva_ret1'],
                    ]

                    filas_datos.append((date_obj, fila_datos))

                    var_creciente['sum_docjson'] += 1
                    totales[0] += valores['total_exenta']
                    totales[1] += valores['sub_total_ventas']
                    totales[2] += valores['iva']
                    totales[3] = 0
                    totales[4] = 0
                    totales[5] = 0
                    totales[6] += valores['monto_total_operacion']
                    totales[7] += valores['iva_perci1']
                    totales[8] += valores['iva_ret1']

        resumen_operaciones["contribuyentes"] = {
            "exentas": totales[0],
            "gravadas": totales[1],
            "debito": totales[2],
            "exportacion": 0,
            "total": totales[6]
        }            

        crear_libro_excel(
            encabezado_registro,
            encabezado_registro1,
            'LIBRO DE VENTAS A CONTRIBUYENTES',
            nombre_contribuyente,
            nrce,
            nit,
            dui,
            periodo,
            anio,
            filas_datos,
            totales,
            nombre_firmador,
            salida,
            resumen_operaciones
        )
        messagebox.showinfo("Éxito", "Libro de Ventas a Contribuyentes se ha generado exitosamente.")
    elif tipo_libro == "Ventas a Consumidor Final":
        nombre_contribuyente = nombre_contri_var.get()
        nrce = nrc_entry.get()
        nit = nit_entry.get()
        dui = dui_entry.get()
        periodo = month_combobox.get()
        anio = year_combobox.get()
        nombre_firmador = nombre_firma_var.get()
        filas_datos = []

        # Encabezado de registros.
        encabezado_registro = ["#", "Fecha", "No Del", "No Al", "Nombre", "Ventas Consumidores","","","","Retencion","Total"]
        encabezado_registro1 = ["", "", "Documento", "Documento", "","Exentas","Gravadas","Exportacion","Terceros","IVA","Venta" ]

        sum_ventas_exentas = 0.0
        sum_ventas_gravadas = 0.0
        sum_ventas_export =0.0
        sum_ventas_terceros=0.0
        sum_retencion_iva=0.0
        sum_monto_total=0.0
        totales=[
            sum_ventas_exentas,
            sum_ventas_gravadas,
            sum_ventas_export,
            sum_ventas_terceros,
            sum_retencion_iva,
            sum_monto_total
        ]
        resumen_operaciones = {
            "contribuyentes": {
                "exentas": None,
                "gravadas": None,
                "debito": None,
                "exportacion": None,
                "total": None
            },
            "consumidor_final": {
                "exentas": None,
                "gravadas": None,
                "debito": None,
                "exportacion": None,
                "total": None
            }
        }
        agrupados = defaultdict(list)
        for nombre_archivo in os.listdir(carpeta):
            ruta_archivo = os.path.join(carpeta, nombre_archivo)
            if os.path.isfile(ruta_archivo) and nombre_archivo.lower().endswith('.json'):
                datos = cargar_datos_archivo(ruta_archivo)
                if datos and 'identificacion' in datos and 'receptor' in datos and 'resumen' in datos:
                    identificacion = datos['identificacion']
                    receptor = datos['receptor']
                    resumen = datos['resumen']
                    tipoDte = identificacion.get('tipoDte')
                    fecEmi = identificacion.get('fecEmi')
                    horEmi = identificacion.get('horEmi')
                    nombre = receptor.get('nombre')
                    codigoGeneracion = identificacion.get('codigoGeneracion')

                    if fecEmi and tipoDte in ["01", "11"]:
                        formatted_date = datetime.strptime(fecEmi, '%Y-%m-%d').strftime('%d/%m/%Y')
                        clave = (formatted_date, tipoDte)
                        total_exenta = resumen.get('totalExenta', 0.00)
                        descu_exenta = resumen.get('descuExenta', 0.00)

                        total_no_suj = resumen.get('totalNoSuj', 0.00)
                        descu_no_suj = resumen.get('descuNoSuj', 0.00)

                        total_no_gravado = resumen.get('totalNoGravado', 0.00)
                        total_gravada = resumen.get('totalGravada', 0.00)
                        descu_gravada = resumen.get('descuGravada', 0.00)
                        agrupados[clave].append({
                            "codigo": codigoGeneracion,
                            "nombre": nombre,
                            "horEmi": horEmi,
                            "total_exenta": total_exenta - descu_exenta,
                            "total_no_sujeta": total_no_suj - descu_no_suj,
                            "total_no_gravado": total_no_gravado,
                            "total_gravada": total_gravada - descu_gravada,
                            "monto_total_operacion": resumen.get('montoTotalOperacion', 0.00),
                            "ivaRete1": resumen.get('ivaRete1', 0.00)
                        })
                        var_creciente['sum_docjson'] += 1

        for (fecha, tipoDte), documentos in agrupados.items():
            documentos_ordenados = sorted(documentos, key=lambda x: x.get("horEmi", "00:00:00"))
            cod_inicio = documentos_ordenados[0]["codigo"]
            cod_fin = documentos_ordenados[-1]["codigo"]
            # Condición que pedís:
            if cod_inicio == cod_fin:
                nombre_proveedor = documentos_ordenados[0]["nombre"]
            else:
                nombre_proveedor = ""
            total_exenta = sum(d["total_exenta"] for d in documentos)
            total_gravada = sum(d["total_gravada"] for d in documentos)
            monto_total_operacion = sum(d["monto_total_operacion"] for d in documentos)
            retencioniva = sum(d["ivaRete1"] for d in documentos)

            fecha_dt = datetime.strptime(fecha, '%d/%m/%Y')

            fila_resumen = [
                var_creciente['sum_docjson'] + 1, fecha, cod_inicio,cod_fin, nombre_proveedor,
            ]

            if tipoDte == "01":
                fila_resumen += [
                    total_exenta, total_gravada, 0,0, retencioniva,
                    monto_total_operacion,
                ]
                totales[0] += total_exenta
                totales[1] += total_gravada
                totales[2] += 0
                totales[3] += 0
                totales[4] += retencioniva
                totales [5] +=monto_total_operacion

            elif tipoDte == "11":
                fila_resumen += [
                    0, 0,total_gravada, 0,retencioniva,
                    monto_total_operacion,
                ]
                totales[0] += total_exenta
                totales[1] += 0
                totales[2] += total_gravada
                totales[3] += 0
                totales[4] += retencioniva
                totales[5] += monto_total_operacion
            filas_datos.append((fecha_dt, fila_resumen))

        filas_datos.sort(key=lambda x: x[0])

        base_gravada = round(totales[1] / 1.13, 2)
        debito_fiscal = round(base_gravada * 0.13, 2)

        resumen_operaciones["consumidor_final"] = {
            "exentas": totales[0],
            "gravadas": base_gravada,
            "debito": debito_fiscal,
            "exportacion": totales[2],
            "total": totales[5]
        }

        crear_libro_excel(
            encabezado_registro,
            encabezado_registro1,
            'LIBRO DE VENTAS A CONSUMIDOR',
            nombre_contribuyente,
            nrce,
            nit,
            dui,
            periodo,
            anio,
            filas_datos,
            totales,
            nombre_firmador,
            salida,
            resumen_operaciones
        )
        messagebox.showinfo("Éxito", "Libro de Ventas a Consumidor Final se ha generado exitosamente.")
    elif tipo_libro not in ["Libro de Compras", "Ventas a Contribuyentes", "Ventas a Consumidor Final"]:
        messagebox.showerror("Error", "Debe seleccionar un tipo de libro válido.")

        return
def on_generate_book(ruta_entrada_var, ruta_salida_var, tipo_libro):
    carpeta = ruta_entrada_var.get()
    salida = ruta_salida_var.get()
    tipo_libro_seleccionado = tipo_libro.get()  # Esto obtiene el texto seleccionado

    if not validar_nombre_contri():
        messagebox.showwarning("Advertencia", "Por favor, complete el campo obligatorio: Nombre del Contribuyente.")
        return

    if not carpeta or not salida:
        messagebox.showwarning("Campos vacíos", "Por favor, complete todos los campos de entrada y salida.")
        return

    if tipo_libro_seleccionado not in ["Libro de Compras", "Ventas a Contribuyentes", "Ventas a Consumidor Final"]:
        messagebox.showerror("Error", "Debe seleccionar un tipo de libro válido.")
        return
    generar_libro_excel(carpeta, salida, tipo_libro_seleccionado)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # Cuando está empaquetado con PyInstaller
    except Exception:
        base_path = os.path.abspath(".")  # Cuando se ejecuta como script .py
    return os.path.join(base_path, relative_path)

def create_secondary_window(ventanaprincipal):

    global nit_var, nit_entry, dui_var, dui_entry, nombre_comercial_entry, nombre_contri_var, error_label, nrc_entry, month_combobox, year_combobox, ruta_salida_var, ruta_entrada_var, nombre_firma_var

    ventanaprincipal.withdraw()
    secondary_root = tk.Toplevel(ventanaprincipal)
    secondary_root.transient(ventanaprincipal)  # ✅ PADRE correcto
    secondary_root.grab_set()  
    # Intentar usar .ico (Windows)
    try:
        secondary_root.iconbitmap(resource_path("iconos/icono_v3.ico"))
    except Exception:
        pass  # No romper la app

    # Usar PNG (SIEMPRE)
    secondary_root.icon_img = tk.PhotoImage(
        file=resource_path("iconos/icono_v3.png")
    )
    secondary_root.iconphoto(True, secondary_root.icon_img)
    secondary_root.title('Generación de Libro de IVA')
    secondary_root.protocol(
        "WM_DELETE_WINDOW",
        lambda: volver_a_interfaz1(secondary_root, ventanaprincipal)
    )    
    secondary_root.resizable(False, False)

    centrar_ventana(secondary_root,575,400)
    
    header_frame = tk.Frame(secondary_root)
    header_frame.pack(fill='x', pady=20, anchor='n')

    header_label = tk.Label(header_frame, text='Detalles del encabezado del Libro', font='Arial 12 bold')
    header_label.pack(side='right', padx=(0, 120))

    back_button = tk.Button(header_frame, text='←', command=lambda: volver_a_interfaz1(secondary_root,ventanaprincipal), width=3)

    back_button.pack(side='left', padx=(20, 0))

    datos_frame = tk.Frame(secondary_root)
    datos_frame.pack(fill='x', pady=5, anchor='n')

    tipolibro = tk.Label(datos_frame, text='Libro de IVA a generar: ', font='Arial 10 bold')
    tipolibro.grid(row=1, column=0, padx=(40, 0), sticky='w', pady=0)

    tipo_libro = tk.StringVar()
    combo_tipo_libro = ttk.Combobox(
        datos_frame,
        textvariable=tipo_libro,
        values=["Libro de Compras", "Ventas a Contribuyentes", "Ventas a Consumidor Final"],
        state="readonly",
        width=25
    )
    combo_tipo_libro.grid(row=1, column=1, padx=(0, 180), pady=0)
    combo_tipo_libro.set("Seleccionar...")
    cliente_frame = tk.Frame(datos_frame)
    cliente_frame.grid(row=2, column=1, sticky="w", pady=5)

    namecontri = tk.Label(datos_frame, text='Nombre de Contribuyente*: ', font='Arial 10 bold')
    namecontri.grid(row=2, column=0, padx=(40, 0), sticky='w', pady=(12,12))
    nombre_contri_var = tk.StringVar()
    nrc_var = tk.StringVar()
    nit_var = tk.StringVar()
    dui_var = tk.StringVar()
    clientes = cargar_clientes()
    clientes_dict = {c["nombre"]: c for c in clientes}

    combo_clientes = ttk.Combobox( cliente_frame, textvariable=nombre_contri_var, values=list(clientes_dict.keys()), width=35, state="readonly" )
    combo_clientes.grid(row=0, column=0, sticky="w")
    
    btn_agregar_cliente = tk.Button(
        cliente_frame,
        text="➕",
        width=3,
        command=lambda: ventana_cliente(secondary_root, actualizar_clientes)
    )
    
    btn_agregar_cliente.grid(row=0, column=1, padx=(18, 0))

    def editar_cliente_seleccionado():
        nombre_seleccionado = combo_clientes.get()
        if not nombre_seleccionado or nombre_seleccionado not in clientes_dict:
            messagebox.showwarning("Seleccionar cliente", "Debe seleccionar un cliente para editar")
            return

        cliente_sel = clientes_dict[nombre_seleccionado]
        ventana_cliente(secondary_root, actualizar_clientes, cliente_seleccionado=cliente_sel)

    btn_editar_cliente = tk.Button(
        cliente_frame,
        text="✏️",
        width=3,
        command=lambda: editar_cliente_seleccionado()
    )
    btn_editar_cliente.grid(row=0, column=2, padx=(5,0))

    def eliminar_cliente_seleccionado():
        nombre_seleccionado = combo_clientes.get()
        if not nombre_seleccionado or nombre_seleccionado not in clientes_dict:
            messagebox.showwarning("Seleccionar cliente", "Debe seleccionar un cliente para eliminar")
            return

        if not messagebox.askyesno("Confirmar", f"¿Desea eliminar al cliente '{nombre_seleccionado}'?"):
            return

        eliminar_cliente(nombre_seleccionado)
        actualizar_clientes()
        messagebox.showinfo("Cliente eliminado", f"Cliente '{nombre_seleccionado}' eliminado correctamente")
    btn_eliminar_cliente = tk.Button(
        cliente_frame,
         text="❌",
        width=3,
        anchor="center",
        justify="center",
        command=eliminar_cliente_seleccionado
    )
    btn_eliminar_cliente.grid(row=0, column=3, padx=(5,0))

    
    def on_cliente_select(event=None):
        nombre = nombre_contri_var.get()
        cliente = clientes_dict.get(nombre)

        if not cliente:
            return

        nrc_var.set(cliente.get("nrc", ""))
        nit_var.set(cliente.get("nit", ""))
        dui_var.set(cliente.get("dui", ""))

    combo_clientes.bind("<<ComboboxSelected>>", on_cliente_select)

    def actualizar_clientes():
        nonlocal clientes_dict
        clientes = cargar_clientes()
        clientes_dict = {c["nombre"]: c for c in clientes}

        nombres = list(clientes_dict.keys())
        combo_clientes["values"] = nombres

        if nombres:
            combo_clientes.set(nombres[-1])
            on_cliente_select()

    #nombre_contri_var = tk.StringVar()
    #nombre_contri.grid(row=2, column=1, padx=(0, 75), sticky='ew', pady=5)
    #nombre_contri_var.trace_add("write", validar_nombre_contri)

    info_frame = tk.Frame(datos_frame)
    info_frame.grid(row=3, column=0, columnspan=2, padx=40, pady=(5, 5), sticky='w')

    nrc_nit_dui_frame = tk.Frame(info_frame)
    nrc_nit_dui_frame.grid(row=0, column=0, padx=0, pady=(0, 10), sticky='w')

    nrc_label = tk.Label(nrc_nit_dui_frame, text='NRC: ', font='Arial 10 bold')
    nrc_label.grid(row=0, column=0, padx=(0, 10), sticky='w')

    nrc_entry = tk.Entry(nrc_nit_dui_frame, textvariable=nrc_var, width=10)
    nrc_entry.grid(row=0, column=1, padx=(0, 10), sticky='ew')

    nit_label = tk.Label(nrc_nit_dui_frame, text='NIT: ', font='Arial 10 bold')
    nit_label.grid(row=0, column=2, padx=(0, 10), sticky='w')

    nit_entry = tk.Entry(nrc_nit_dui_frame, textvariable=nit_var, width=19)
    nit_entry.grid(row=0, column=3, padx=(0, 21), sticky='ew')
    nit_var.trace_add("write", format_nit_input)

    dui_label = tk.Label(nrc_nit_dui_frame, text='DUI: ', font='Arial 10 bold')
    dui_label.grid(row=0, column=4, padx=(0, 10), sticky='w')
    dui_entry = tk.Entry(nrc_nit_dui_frame, textvariable=dui_var, width=16)
    dui_entry.grid(row=0, column=5, padx=(3, 15), sticky='ew')
    dui_var.trace_add("write", format_dui_input)

    periodo_frame = tk.Frame(info_frame)
    periodo_frame.grid(row=1, column=0, padx=0, pady=0, sticky='w')

    period_label = tk.Label(periodo_frame, text='Periodo Tributario: ', font='Arial 10 bold')
    period_label.grid(row=0, column=0, padx=(0, 10), sticky='w')

    months = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    month_combobox = ttk.Combobox(periodo_frame, values=months, width=10,state='readonly')
    month_combobox.grid(row=0, column=1, padx=(0, 10), sticky='ew')
    month_combobox.set('Enero')

    year_label = tk.Label(periodo_frame, text='Año: ', font='Arial 10 bold')
    year_label.grid(row=0, column=2, padx=(0, 10), sticky='w')

    years = [str(year) for year in range(2000, 2031)]
    year_combobox = ttk.Combobox(periodo_frame, values=years, width=5,state='readonly')
    year_combobox.grid(row=0, column=3, padx=(0, 10), sticky='ew')
    year_combobox.set('2023')

    # Campo de Nombre del Contador (movido antes de Ruta de Entrada)
    namefirma_label = tk.Label(datos_frame, text='Nombre del Contador: ', font='Arial 10 bold')
    namefirma_label.grid(row=4, column=0, padx=(40, 0), sticky='w', pady=12)

    nombre_firma_var = tk.StringVar()
    nombre_fir = tk.Entry(datos_frame, textvariable=nombre_firma_var, width=45 )
    nombre_fir.grid(row=4, column=1, padx=(0, 75), sticky='ew', pady=5)

    # Frame para los campos de entrada
    frame_entradas = tk.Frame(info_frame)
    frame_entradas.grid(row=2, column=0, padx=0, pady=(10, 0), sticky='w')

    # Etiqueta y entrada para la carpeta de entrada
    ruta_entrada_label = tk.Label(frame_entradas, text='Ruta de Entrada: ', font='Arial 10 bold')
    ruta_entrada_label.grid(row=1, column=0, padx=(0, 3), sticky='w')

    ruta_entrada_var = tk.StringVar()
    ruta_entrada_entry = tk.Entry(frame_entradas, textvariable=ruta_entrada_var, width=42)
    ruta_entrada_entry.grid(row=1, column=1, padx=(0, 10), sticky='ew')

    seleccionar_rutae_button = tk.Button(frame_entradas, text='Seleccionar', command=seleccionar_ruta_entrada)
    seleccionar_rutae_button.grid(row=1, column=2, padx=(0, 5), sticky='ew')

    # Ruta de salida
    ruta_salida_frame = tk.Frame(info_frame)
    ruta_salida_frame.grid(row=3, column=0, padx=0, pady=(10, 0), sticky='w')

    ruta_salida_label = tk.Label(ruta_salida_frame, text='Ruta de Salida: ', font='Arial 10 bold')
    ruta_salida_label.grid(row=1, column=0, padx=(0, 10), sticky='w')

    ruta_salida_var = tk.StringVar()
    ruta_salida_entry = tk.Entry(ruta_salida_frame, textvariable=ruta_salida_var, width=42)
    ruta_salida_entry.grid(row=1, column=1, padx=(0, 10), sticky='ew')

    seleccionar_ruta_button = tk.Button(ruta_salida_frame, text='Seleccionar', command=seleccionar_ruta_salida)
    seleccionar_ruta_button.grid(row=1, column=2, padx=(0, 5), sticky='ew')

    generate_button = tk.Button(datos_frame, text='Generar Libro', command=lambda: on_generate_book(ruta_entrada_var, ruta_salida_var, tipo_libro))
    generate_button.grid(row=5, column=0, columnspan=2, pady=15)

    secondary_root.mainloop()

# Función para volver a la primera interfaz
def volver_a_interfaz1(ventana2, ventana_principal):
    ventana2.destroy()
    ventana_principal.deiconify()  # Ejecutar la primera interfaz













