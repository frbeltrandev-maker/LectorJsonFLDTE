import json
import os
import sys


BASE_DIR = os.path.dirname(sys.executable)
ARCHIVO_CLIENTES = os.path.join(BASE_DIR, "clientes.json")


def cargar_clientes():
    if not os.path.exists(ARCHIVO_CLIENTES):
        return []
    with open(ARCHIVO_CLIENTES, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar_cliente(cliente):
    clientes = cargar_clientes()
    clientes.append(cliente)
    with open(ARCHIVO_CLIENTES, "w", encoding="utf-8") as f:
        json.dump(clientes, f, indent=4, ensure_ascii=False)


def editar_cliente(nombre_original, cliente_actualizado):
    """
    Edita un cliente existente.
    Si no existe, lo agrega como nuevo.
    """
    clientes = cargar_clientes()
    for i, c in enumerate(clientes):
        if c["nombre"] == nombre_original:
            clientes[i] = cliente_actualizado
            break
    else:
        clientes.append(cliente_actualizado)

    with open(ARCHIVO_CLIENTES, "w", encoding="utf-8") as f:
        json.dump(clientes, f, indent=4, ensure_ascii=False)


def eliminar_cliente(nombre):
    clientes = cargar_clientes()
    clientes = [c for c in clientes if c["nombre"] != nombre]
    with open(ARCHIVO_CLIENTES, "w", encoding="utf-8") as f:
        json.dump(clientes, f, indent=4, ensure_ascii=False)
