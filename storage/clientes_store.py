import json
import os
import sys


APP_NAME = "LJSON_V3"
CARPETA_DATOS = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), APP_NAME)
ARCHIVO_CLIENTES = os.path.join(CARPETA_DATOS, "clientes.json")
ARCHIVO_CLIENTES_LEGADO = os.path.join(os.path.dirname(sys.executable), "clientes.json")
RAIZ_PROYECTO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARCHIVO_CLIENTES_DESARROLLO = os.path.join(RAIZ_PROYECTO, "dist", "clientes.json")


def ruta_clientes_existente():
    if os.path.exists(ARCHIVO_CLIENTES):
        return ARCHIVO_CLIENTES
    if os.path.exists(ARCHIVO_CLIENTES_LEGADO):
        return ARCHIVO_CLIENTES_LEGADO
    if os.path.exists(ARCHIVO_CLIENTES_DESARROLLO):
        return ARCHIVO_CLIENTES_DESARROLLO
    return None


def guardar_clientes(clientes):
    os.makedirs(CARPETA_DATOS, exist_ok=True)
    with open(ARCHIVO_CLIENTES, "w", encoding="utf-8") as f:
        json.dump(clientes, f, indent=4, ensure_ascii=False)


def cargar_clientes():
    ruta = ruta_clientes_existente()
    if ruta is None:
        return []
    with open(ruta, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar_cliente(cliente):
    clientes = cargar_clientes()
    clientes.append(cliente)
    guardar_clientes(clientes)


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

    guardar_clientes(clientes)


def eliminar_cliente(nombre):
    clientes = cargar_clientes()
    clientes = [c for c in clientes if c["nombre"] != nombre]
    guardar_clientes(clientes)
