class ControlDuplicadosDTE:
    """Conserva la primera aparición de cada código dentro de una generación.

    Los archivos se recorren por nombre para que la selección sea reproducible.
    Sin un código no se puede establecer identidad: no se descartan documentos
    por compartir una fecha, un importe o un nombre de archivo parecido.
    """

    def __init__(self):
        self.archivos_por_codigo = {}
        self.avisos = []

    def es_duplicado(self, identificacion, nombre_archivo):
        codigo = identificacion.get("codigoGeneracion")
        if not isinstance(codigo, str) or not codigo.strip():
            return False

        clave = codigo.strip().upper()
        if clave in self.archivos_por_codigo:
            original = self.archivos_por_codigo[clave]
            self.avisos.append((
                nombre_archivo,
                f"DTE duplicado omitido: {clave}. Primera aparición: {original}.",
            ))
            return True

        self.archivos_por_codigo[clave] = nombre_archivo
        return False
