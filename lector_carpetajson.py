import os
import inspect
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys
from libroIVA import create_secondary_window,centrar_ventana
from dte.types import (
    TIPO_DOCUMENTO_ANTICIPO_IVA,
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
from reports.anexos_anticipo_iva import generar_anexo_anticipo_iva
from reports.anexos_anulados import generar_anexo_anulados
from reports.anexos_contribuyentes import generar_anexo_contribuyentes
from reports.anexos_consumidor_final import generar_anexo_consumidor_final
from reports.anexos_retencion import generar_anexo_retencion
from reports.anexos_sujeto_excluido import generar_anexo_sujeto_excluido


REPORTES = {
    TIPO_DOCUMENTO_ANTICIPO_IVA: (
        generar_anexo_anticipo_iva,
        "Anticipo a cuenta IVA 2%",
        "Anexo 6 generado. Revise el archivo de errores si fue creado.",
    ),
    TIPO_DOCUMENTO_COMPRAS: (
        generar_anexo_compras,
        "Compras",
        "Los archivos han sido procesados exitosamente.",
    ),
    TIPO_DOCUMENTO_RETENCION: (
        generar_anexo_retencion,
        "Retención",
        "Comprobantes de retención procesados correctamente.",
    ),
    TIPO_DOCUMENTO_CONTRIBUYENTES: (
        generar_anexo_contribuyentes,
        "Ventas a Contribuyentes",
        "Los archivos han sido procesados exitosamente.",
    ),
    TIPO_DOCUMENTO_CONSUMIDOR_FINAL: (
        generar_anexo_consumidor_final,
        "Ventas a Consumidor Final",
        "Los archivos han sido procesados exitosamente.",
    ),
    TIPO_DOCUMENTO_SUJETO_EXCLUIDO: (
        generar_anexo_sujeto_excluido,
        "Sujeto Excluido",
        "Los archivos han sido procesados exitosamente.",
    ),
    TIPO_DOCUMENTO_ANULADOS: (
        generar_anexo_anulados,
        "Documentos Anulados",
        "Han sido procesados exitosamente.",
    ),
}

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

    def reportar_estado(mensaje):
        eventos.put(("estado", mensaje))

    def ejecutar_reporte():
        try:
            argumentos = {"progreso": reportar_progreso}
            if "estado" in inspect.signature(generar_reporte).parameters:
                argumentos["estado"] = reportar_estado
            resumen = generar_reporte(
                carpeta,
                salida,
                nombre_archivo_salida,
                **argumentos,
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
            elif tipo_evento == "estado":
                _, mensaje = evento
                estado_var.set(mensaje)
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

    if not carpeta or not salida or not nombre_archivo_salida:
        messagebox.showwarning(
            "Campos vacíos",
            "Por favor, complete todos los campos y seleccione un tipo de documento antes de iniciar el procesamiento."
        )
        return

    reporte = REPORTES.get(tipo_documento_seleccionado)
    if not reporte:
        messagebox.showerror("Error", "Debe seleccionar un tipo de documento válido.")
        return

    generar_reporte, titulo_reporte, mensaje_exito = reporte
    procesar_reporte_en_segundo_plano(
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
        root.iconbitmap(resource_path("iconos/iconofldte.ico"))
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
