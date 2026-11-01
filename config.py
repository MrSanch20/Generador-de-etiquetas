"""
Ajustes que el programa recuerda entre usos (en %APPDATA%\\EtiquetadorCuadros\\config.json):
  - SW-CODE por defecto de los elementos (181 si no se ha cambiado)
  - SW-CODE de cada regleta (XALI, XSEN…), el último que se puso
"""
import json
import os

SWCODE_ELEMENTOS_DEFECTO = "181"


def _ruta():
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "EtiquetadorCuadros", "config.json")


def cargar():
    try:
        with open(_ruta(), encoding="utf-8") as f:
            datos = json.load(f)
    except (OSError, ValueError):
        datos = {}
    datos.setdefault("swcode_elementos", SWCODE_ELEMENTOS_DEFECTO)
    datos.setdefault("swcode_regletas", {})
    return datos


def guardar(datos):
    try:
        os.makedirs(os.path.dirname(_ruta()), exist_ok=True)
        with open(_ruta(), "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except OSError:
        pass   # si no se puede guardar, el programa sigue funcionando igual
