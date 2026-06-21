import json
import tempfile
import unittest
from pathlib import Path

from storage import clientes_store


class ClientesStoreTests(unittest.TestCase):
    def setUp(self):
        self.directorio_temporal = tempfile.TemporaryDirectory()
        self.ruta_temporal = Path(self.directorio_temporal.name)
        self.rutas_originales = {
            "CARPETA_DATOS": clientes_store.CARPETA_DATOS,
            "ARCHIVO_CLIENTES": clientes_store.ARCHIVO_CLIENTES,
            "ARCHIVO_CLIENTES_LEGADO": clientes_store.ARCHIVO_CLIENTES_LEGADO,
            "ARCHIVO_CLIENTES_DESARROLLO": clientes_store.ARCHIVO_CLIENTES_DESARROLLO,
        }
        clientes_store.CARPETA_DATOS = str(self.ruta_temporal / "datos")
        clientes_store.ARCHIVO_CLIENTES = str(self.ruta_temporal / "datos" / "clientes.json")
        clientes_store.ARCHIVO_CLIENTES_LEGADO = str(self.ruta_temporal / "legado" / "clientes.json")
        clientes_store.ARCHIVO_CLIENTES_DESARROLLO = str(self.ruta_temporal / "desarrollo" / "clientes.json")

    def tearDown(self):
        for nombre, valor in self.rutas_originales.items():
            setattr(clientes_store, nombre, valor)
        self.directorio_temporal.cleanup()

    def test_guarda_edita_y_elimina_cliente(self):
        cliente = {"nombre": "Cliente Uno", "nrc": "1", "nit": "2", "dui": "3"}
        cliente_editado = {"nombre": "Cliente Editado", "nrc": "4", "nit": "5", "dui": "6"}

        clientes_store.guardar_cliente(cliente)
        self.assertEqual(clientes_store.cargar_clientes(), [cliente])

        clientes_store.editar_cliente("Cliente Uno", cliente_editado)
        self.assertEqual(clientes_store.cargar_clientes(), [cliente_editado])

        clientes_store.eliminar_cliente("Cliente Editado")
        self.assertEqual(clientes_store.cargar_clientes(), [])

    def test_lee_datos_legados_y_los_migra_al_guardar(self):
        ruta_legado = Path(clientes_store.ARCHIVO_CLIENTES_LEGADO)
        ruta_legado.parent.mkdir(parents=True)
        cliente_legado = {"nombre": "Cliente Legado"}
        ruta_legado.write_text(json.dumps([cliente_legado]), encoding="utf-8")

        self.assertEqual(clientes_store.cargar_clientes(), [cliente_legado])

        clientes_store.guardar_cliente({"nombre": "Cliente Nuevo"})
        self.assertTrue(Path(clientes_store.ARCHIVO_CLIENTES).exists())
        self.assertEqual(
            clientes_store.cargar_clientes(),
            [cliente_legado, {"nombre": "Cliente Nuevo"}],
        )
