import hashlib
import os
import sys
import uuid
import tkinter as tk
from tkinter import messagebox


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
    root.overrideredirect(True)
    root.resizable(False, False)
    width = 420
    height = 220
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

    tk.Label(root, text="Activación LJSON V3", font=("Arial", 10, "bold")).pack(pady=(12, 4))
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
