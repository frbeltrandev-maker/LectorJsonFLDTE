import json
import io
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

from dte.reader import cargar_datos_archivo


class ReaderTests(unittest.TestCase):
    def test_carga_json_utf8_sig(self):
        with tempfile.TemporaryDirectory() as directorio:
            ruta = Path(directorio) / "dte.json"
            ruta.write_text(json.dumps({"identificacion": {}}), encoding="utf-8-sig")

            self.assertEqual(cargar_datos_archivo(ruta), {"identificacion": {}})

    def test_retorna_none_para_archivo_vacio(self):
        with tempfile.TemporaryDirectory() as directorio:
            ruta = Path(directorio) / "vacio.json"
            ruta.write_text("", encoding="utf-8")

            self.assertIsNone(cargar_datos_archivo(ruta))

    def test_retorna_none_para_json_invalido(self):
        with tempfile.TemporaryDirectory() as directorio:
            ruta = Path(directorio) / "invalido.json"
            ruta.write_text("{invalido", encoding="utf-8")

            self.assertIsNone(cargar_datos_archivo(ruta))

    def test_carga_json_cp1252_sin_mostrar_error_intermedio(self):
        with tempfile.TemporaryDirectory() as directorio:
            ruta = Path(directorio) / "cp1252.json"
            ruta.write_bytes(json.dumps({"nombre": "Café"}, ensure_ascii=False).encode("cp1252"))
            salida = io.StringIO()

            with redirect_stdout(salida):
                datos = cargar_datos_archivo(ruta)

            self.assertEqual(datos, {"nombre": "Café"})
            self.assertEqual(salida.getvalue(), "")
