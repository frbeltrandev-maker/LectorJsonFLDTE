import json


def cargar_datos_archivo(ruta):
    codificaciones = ["utf-8-sig", "latin-1", "cp1252"]
    for codificacion in codificaciones:
        try:
            with open(ruta, "r", encoding=codificacion) as archivo:
                contenido = archivo.read()
                if not contenido.strip():
                    print(f"El archivo {ruta} está vacío.")
                    return None
                return json.loads(contenido)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Error decodificando JSON en el archivo {ruta} con codificación {codificacion}: {e}")
        except Exception as e:
            print(f"Error inesperado en el archivo {ruta} con codificación {codificacion}: {e}")
    return None
