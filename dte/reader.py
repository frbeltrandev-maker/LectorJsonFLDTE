import json


def cargar_datos_archivo(ruta):
    codificaciones = ["utf-8-sig", "latin-1", "cp1252"]
    errores = []

    for codificacion in codificaciones:
        try:
            with open(ruta, "r", encoding=codificacion) as archivo:
                contenido = archivo.read()
                if not contenido.strip():
                    print(f"El archivo {ruta} está vacío.")
                    return None
                return json.loads(contenido)

        except FileNotFoundError:
            print(f"No se encontró el archivo {ruta}.")
            return None
        except (UnicodeDecodeError, json.JSONDecodeError) as error:
            errores.append(f"{codificacion}: {error}")
        except OSError as error:
            print(f"No se pudo leer el archivo {ruta}: {error}")
            return None

    if errores:
        print(
            f"No se pudo cargar el JSON {ruta}. "
            f"Codificaciones probadas: {', '.join(codificaciones)}. "
            f"Último error: {errores[-1]}"
        )
    return None
