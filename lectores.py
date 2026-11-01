"""
Lectores de archivos.

Cada lector convierte un archivo (PDF, DXF/DWG, imagen escaneada) en una lista
de páginas con sus textos y posiciones, en un formato común. Así la lógica de
bornas y elementos (extraccion.py) es la misma para todos los formatos.

Coordenadas: como en un PDF, con el origen arriba a la izquierda y la Y hacia
abajo, y escaladas para que el alto de página sea ~842 (un A3/A4 en puntos).
"""
import glob
import math
import os
import re
from dataclasses import dataclass, field

ALTO_REF = 841.89          # alto de página de referencia (puntos)
HORIZONTAL = (1.0, 0.0)
VERTICAL = (0.0, -1.0)     # girado 90°, se lee de abajo arriba

EXT_PDF = (".pdf",)
EXT_DXF = (".dxf",)
EXT_DWG = (".dwg",)
EXT_IMAGEN = (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp")
EXT_TABLA = (".xlsx", ".xlsm", ".xls", ".csv", ".txt")

# Formatos admitidos, para mostrarlos en la ventana: (nombre, extensiones, nota)
FORMATOS = (
    ("PDF del programa de esquemas", EXT_PDF, "El más fiable."),
    ("PDF escaneado o imagen", EXT_PDF + EXT_IMAGEN,
     "Se lee con reconocimiento de texto (OCR). Puede equivocarse: revisa la lista. ~10 s por página."),
    ("AutoCAD DXF", EXT_DXF, "Textos y atributos de bloques (AutoCAD Electrical)."),
    ("AutoCAD DWG", EXT_DWG, "Necesita «ODA File Converter» (gratuito) instalado; si no, guárdalo como DXF."),
    ("Listado Excel / CSV", (".xlsx", ".xlsm", ".csv", ".txt"),
     "Lista de bornas o de aparatos exportada de EPLAN, SEE, AutoCAD Electrical… "
     "El .xls antiguo no: guárdalo como .xlsx."),
)


def extensiones_admitidas():
    """Todas las extensiones que se pueden abrir, sin repetir y en orden: '.pdf .dxf …'."""
    vistas = []
    for _, exts, _ in FORMATOS:
        for e in exts:
            if e not in vistas:
                vistas.append(e)
    return vistas


class ErrorLectura(Exception):
    """Error con un mensaje pensado para el usuario."""


@dataclass
class Texto:
    texto: str
    x0: float
    y0: float
    x1: float
    y1: float
    dir: tuple = HORIZONTAL
    tam: float = 0.0          # tamaño de letra
    estilo: object = None     # color, capa… lo que distinga el estilo del texto


@dataclass
class Pagina:
    numero: int
    ancho: float
    alto: float
    lineas: list = field(default_factory=list)   # textos agrupados por línea
    trozos: list = field(default_factory=list)   # textos sueltos (spans)
    origen: str = "pdf"                          # pdf | dxf | ocr
    sin_ocr: bool = False                        # página escaneada que no se ha podido leer (falta el OCR)

    def texto_zona(self, y_desde=0.0, y_hasta=None):
        """Todo el texto cuyo centro está entre dos alturas, unido con espacios."""
        y_hasta = self.alto if y_hasta is None else y_hasta
        return " ".join(t.texto for t in self.trozos if y_desde <= (t.y0 + t.y1) / 2 <= y_hasta)


def _dir(d):
    return tuple(round(v, 1) + 0.0 for v in d)   # +0.0 evita el -0.0


# --------------------------------------------------------------------------
# PDF (con texto, o escaneado -> OCR)
# --------------------------------------------------------------------------
def leer_pdf(ruta, progreso=None):
    import pymupdf

    try:
        doc = pymupdf.open(ruta)
    except Exception as e:
        raise ErrorLectura(f"No se puede abrir el PDF: {e}")
    if doc.needs_pass:
        raise ErrorLectura("El PDF está protegido con contraseña.")
    paginas = []
    try:
        for num, page in enumerate(doc, start=1):
            if progreso:
                progreso(f"Leyendo página {num} de {doc.page_count}…")
            esc = ALTO_REF / page.rect.height
            pag = Pagina(num, page.rect.width * esc, ALTO_REF)
            for b in page.get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    spans = [s for s in l["spans"] if s["text"].strip()]
                    if not spans:
                        continue
                    d = _dir(l["dir"])
                    for s in spans:
                        x0, y0, x1, y1 = (v * esc for v in s["bbox"])
                        pag.trozos.append(Texto(s["text"].strip(), x0, y0, x1, y1, d, s["size"] * esc, s["color"]))
                    x0, y0, x1, y1 = (v * esc for v in l["bbox"])
                    texto = "".join(s["text"] for s in l["spans"]).strip()
                    pag.lineas.append(Texto(texto, x0, y0, x1, y1, d, spans[0]["size"] * esc, spans[0]["color"]))
            if len(pag.trozos) < 3:
                # Página sin texto: probablemente escaneada -> OCR
                if progreso:
                    progreso(f"Página {num}: sin texto, aplicando OCR…")
                if ocr_disponible():
                    pix = page.get_pixmap(dpi=200)
                    from PIL import Image
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    pag = _ocr_pagina(img, num)
                else:
                    # sin OCR: se salta esta página y se sigue con las demás
                    pag.sin_ocr = True
            paginas.append(pag)
    finally:
        doc.close()
    return paginas


# --------------------------------------------------------------------------
# Imágenes escaneadas -> OCR
# --------------------------------------------------------------------------
_OCR = None


def ocr_disponible():
    """True si este programa incluye el reconocimiento de texto (OCR)."""
    global _OCR
    try:
        _motor_ocr()
        return True
    except Exception:          # no instalado, o instalado pero roto (DLL que no carga…)
        _OCR = False
        return False


def _motor_ocr():
    """Motor de OCR: RapidOCR 3.x (cualquier Python) o rapidocr_onnxruntime 1.x (Python <= 3.12).

    Devuelve una función imagen(numpy) -> [(caja, texto, confianza), …].
    """
    global _OCR
    if _OCR is False:
        raise ErrorLectura("El reconocimiento de texto (OCR) no está instalado en este programa.\n"
                           "Vuelve a crear el .exe con crear_exe.bat (usa Python 3.12 de 64 bits).")
    if _OCR is None:
        try:
            import logging
            from rapidocr import RapidOCR
            logging.getLogger("RapidOCR").setLevel(logging.WARNING)
            motor = RapidOCR()

            def _ocr(img):
                r = motor(img)
                if r is None or r.txts is None:
                    return []
                return list(zip(r.boxes.tolist(), r.txts, [float(x) for x in r.scores]))
        except ImportError:
            try:
                from rapidocr_onnxruntime import RapidOCR
            except ImportError:
                _OCR = False
                raise ErrorLectura("El reconocimiento de texto (OCR) no está instalado en este programa.\n"
                                   "Vuelve a crear el .exe con crear_exe.bat (usa Python 3.12 de 64 bits).")
            motor = RapidOCR()

            def _ocr(img):
                res, _ = motor(img)
                return res or []
        _OCR = _ocr
    return _OCR


def _ocr_pagina(img, numero):
    """OCR de una imagen. Lee el texto horizontal y, girando la imagen, el vertical."""
    import numpy as np

    ocr = _motor_ocr()
    w, h = img.size
    esc = ALTO_REF / h
    pag = Pagina(numero, w * esc, ALTO_REF, origen="ocr")

    def cajas(im):
        for box, txt, conf in ocr(np.array(im)):
            txt = txt.strip()
            if txt and conf >= 0.5:
                xs = [p[0] for p in box]
                ys = [p[1] for p in box]
                yield txt, min(xs), min(ys), max(xs), max(ys), conf

    # 1) texto vertical (de abajo arriba): girar la imagen 90° a la derecha
    verticales = []
    rot = img.rotate(-90, expand=True)          # tamaño (h, w)
    for txt, rx0, ry0, rx1, ry1, conf in cajas(rot):
        ancho, alto = rx1 - rx0, ry1 - ry0
        # en la imagen girada el texto vertical es horizontal (más ancho que alto);
        # los textos muy cortos ("6", "20") son casi cuadrados
        if ancho < alto * (0.6 if len(txt) <= 3 else 1.2):
            continue
        # punto (rx, ry) de la imagen girada -> (x, y) original: x = ry, y = h - rx
        verticales.append((txt, ry0, h - rx1, ry1, h - rx0, conf, ry1 - ry0))
    # 2) texto horizontal: cajas más anchas que altas (o textos cortos casi cuadrados)
    horizontales = []
    for txt, x0, y0, x1, y1, conf in cajas(img):
        if (x1 - x0) >= (y1 - y0) * (0.6 if len(txt) <= 3 else 0.9):
            horizontales.append((txt, x0, y0, x1, y1, conf, y1 - y0))

    def solapa(a, b):
        ix = max(0, min(a[3], b[3]) - max(a[1], b[1]))
        iy = max(0, min(a[4], b[4]) - max(a[2], b[2]))
        inter = ix * iy
        menor = min((a[3] - a[1]) * (a[4] - a[2]), (b[3] - b[1]) * (b[4] - b[2])) or 1
        return inter / menor > 0.5

    # Un texto corto puede salir en las dos pasadas: se queda la lectura más segura
    for lista, otra, d in ((verticales, horizontales, VERTICAL), (horizontales, verticales, HORIZONTAL)):
        for c in lista:
            if any(solapa(c, o) and o[5] > c[5] for o in otra):
                continue
            txt, x0, y0, x1, y1, conf, grueso = c
            t = Texto(txt, x0 * esc, y0 * esc, x1 * esc, y1 * esc, d, grueso * esc * 0.8, "ocr")
            pag.trozos.append(t)
            pag.lineas.append(t)
    return pag


def leer_imagen(ruta, progreso=None):
    from PIL import Image, ImageSequence

    try:
        im = Image.open(ruta)
    except Exception as e:
        raise ErrorLectura(f"No se puede abrir la imagen: {e}")
    paginas = []
    for num, frame in enumerate(ImageSequence.Iterator(im), start=1):   # TIFF multipágina
        if progreso:
            progreso(f"Reconociendo texto (OCR) de la página {num}…")
        paginas.append(_ocr_pagina(frame.convert("RGB"), num))
    return paginas


# --------------------------------------------------------------------------
# DXF / DWG (AutoCAD)
# --------------------------------------------------------------------------
def _textos_dxf(entidades, capa_por_defecto=""):
    """TEXT, MTEXT y ATTRIB (también dentro de bloques) -> (texto, x, y, rot, alto, estilo, ancho)."""
    from ezdxf.math import Vec3

    for e in entidades:
        tipo = e.dxftype()
        try:
            if tipo == "INSERT":
                for a in e.attribs:
                    if not a.is_invisible:
                        yield from _textos_dxf([a])
                yield from _textos_dxf(e.virtual_entities())
            elif tipo in ("TEXT", "ATTRIB"):
                txt = e.plain_text().strip() if hasattr(e, "plain_text") else e.dxf.text.strip()
                if not txt:
                    continue
                alto = e.dxf.height
                rot = e.dxf.rotation
                # punto de arranque del texto (abajo a la izquierda, sin girar)
                p = e.dxf.insert
                halign = e.dxf.get("halign", 0)
                if halign or e.dxf.get("valign", 0):
                    p = e.dxf.get("align_point", p)
                ancho = len(txt) * alto * 0.75
                if halign in (1, 4):          # centrado
                    p = Vec3(p) - Vec3.from_deg_angle(rot, ancho / 2)
                elif halign == 2:             # derecha
                    p = Vec3(p) - Vec3.from_deg_angle(rot, ancho)
                yield txt, p.x, p.y, rot, alto, (round(alto, 2), e.dxf.layer, e.dxf.color), ancho
            elif tipo == "MTEXT":
                txt = e.plain_text().strip()
                if not txt:
                    continue
                alto = e.dxf.char_height
                rot = e.get_rotation()
                p = Vec3(e.dxf.insert)
                ancho = max(len(l) for l in txt.splitlines()) * alto * 0.75
                # punto de anclaje: 1-3 arriba, 4-6 centro, 7-9 abajo; 1/4/7 izq, 2/5/8 centro, 3/6/9 der
                att = e.dxf.get("attachment_point", 1)
                fila, col = (att - 1) // 3, (att - 1) % 3
                n_lineas = len(txt.splitlines())
                p = p - Vec3.from_deg_angle(rot, ancho * col / 2)
                p = p + Vec3.from_deg_angle(rot + 90, alto * n_lineas * (1 - fila / 2) - alto)
                for i, linea in enumerate(txt.splitlines()):
                    if linea.strip():
                        q = p - Vec3.from_deg_angle(rot + 90, i * alto * 1.6)
                        yield (linea.strip(), q.x, q.y, rot, alto, (round(alto, 2), e.dxf.layer, e.dxf.color),
                               len(linea) * alto * 0.75)
        except Exception:
            continue   # una entidad rara no debe parar la lectura


def _pagina_dxf(numero, textos, origen="dxf"):
    """Convierte textos DXF (Y hacia arriba) a una Pagina con Y hacia abajo."""
    if not textos:
        return None
    xs, ys = [], []
    for txt, x, y, rot, alto, est, ancho in textos:
        xs += [x, x + ancho]
        ys += [y, y + alto]
    xmin, xmax, ymin, ymax = min(xs), max(xs), min(ys), max(ys)
    alto_pag = max(ymax - ymin, 1e-6) * 1.02
    esc = ALTO_REF / alto_pag
    pag = Pagina(numero, (xmax - xmin) * esc, ALTO_REF, origen=origen)
    for txt, x, y, rot, alto, est, ancho in textos:
        r = rot % 360
        if abs(r - 90) < 5:
            d = VERTICAL
            # texto girado 90°: crece hacia arriba desde (x, y); el cuerpo queda a la izquierda de x
            x0, x1 = x - alto, x
            yb0, yb1 = y, y + ancho
        else:
            d = HORIZONTAL if (r < 5 or r > 355) else _dir((math.cos(math.radians(r)), -math.sin(math.radians(r))))
            x0, x1 = x, x + ancho
            yb0, yb1 = y, y + alto
        # Y hacia abajo y escala
        X0, X1 = (x0 - xmin) * esc, (x1 - xmin) * esc
        Y0, Y1 = (ymax - yb1) * esc, (ymax - yb0) * esc
        t = Texto(txt, X0, Y0, X1, Y1, d, alto * esc, est)
        pag.trozos.append(t)
        pag.lineas.append(t)
    return pag


def _cajas_geometria(entidades):
    """Rectángulos (xmin, xmax, ymin, ymax) de líneas y polilíneas del dibujo."""
    for e in entidades:
        try:
            t = e.dxftype()
            if t == "LINE":
                pts = [e.dxf.start, e.dxf.end]
            elif t == "LWPOLYLINE":
                pts = [(p[0], p[1]) for p in e.get_points("xy")]
            elif t == "POLYLINE":
                pts = [v.dxf.location for v in e.vertices]
            else:
                continue
            if pts:
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                yield min(xs), max(xs), min(ys), max(ys)
        except Exception:
            continue


def _grupos(intervalos):
    """Une intervalos (a, b) que se solapan o tocan; devuelve los grupos resultantes."""
    grupos = []
    for a, b in sorted(intervalos):
        if grupos and a <= grupos[-1][1]:
            grupos[-1][1] = max(grupos[-1][1], b)
        else:
            grupos.append([a, b])
    return grupos


def _separar_hojas(textos, cajas):
    """Si el espacio modelo tiene varias hojas (una al lado de otra o en cuadrícula),
    las separa por las franjas vacías: ni textos ni líneas (el marco de cada hoja la
    mantiene unida)."""
    def caja_texto(t):
        txt, x, y, rot, alto, est, ancho = t
        if abs(rot % 360 - 90) < 5:
            return x - alto, x, y, y + ancho
        return x, x + ancho, y, y + alto

    elementos = [("t", caja_texto(t), t) for t in textos] + [("g", c, None) for c in cajas]
    margen = 1e-6
    hojas = []
    for gx in _grupos([(c[0] - margen, c[1] + margen) for _, c, _ in elementos]):
        col = [e for e in elementos if gx[0] <= (e[1][0] + e[1][1]) / 2 <= gx[1]]
        for gy in sorted(_grupos([(c[2] - margen, c[3] + margen) for _, c, _ in col]), key=lambda g: -g[1]):
            hoja = [t for k, c, t in col if k == "t" and gy[0] <= (c[2] + c[3]) / 2 <= gy[1]]
            if len(hoja) >= 3:
                hojas.append((gx[0], -gy[1], hoja))
    # de arriba abajo y de izquierda a derecha
    hojas.sort(key=lambda h: (round(h[1]), h[0]))
    return [h[2] for h in hojas] or [textos]


def leer_dxf(ruta, progreso=None, doc=None):
    try:
        import ezdxf
        from ezdxf import recover
    except ImportError:
        raise ErrorLectura("La lectura de DXF no está instalada en este programa. Vuelve a crear el .exe.")
    if doc is None:
        if progreso:
            progreso("Abriendo el DXF…")
        try:
            doc, _ = recover.readfile(ruta)
        except Exception as e:
            raise ErrorLectura(f"No se puede abrir el DXF: {e}")
    paginas = []
    # 1) Presentaciones (cada una es una hoja), si tienen texto propio
    for layout in doc.layouts:
        if layout.name.lower() == "model":
            continue
        textos = list(_textos_dxf(layout))
        if len(textos) >= 5:
            pag = _pagina_dxf(len(paginas) + 1, textos)
            if pag:
                paginas.append(pag)
    # 2) Si no, el espacio modelo (separando hojas si hay varias)
    if not paginas:
        msp = doc.modelspace()
        textos = list(_textos_dxf(msp))
        if not textos:
            raise ErrorLectura("El DXF no tiene textos que leer.")
        for hoja in _separar_hojas(textos, list(_cajas_geometria(msp))):
            pag = _pagina_dxf(len(paginas) + 1, hoja)
            if pag:
                paginas.append(pag)
    return paginas


def _buscar_oda():
    candidatos = []
    for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        candidatos += glob.glob(os.path.join(base, "ODA", "*", "ODAFileConverter.exe"))
    return sorted(candidatos)[-1] if candidatos else None


def leer_dwg(ruta, progreso=None):
    """El DWG es un formato cerrado: se convierte a DXF con ODA File Converter (gratuito)."""
    try:
        import ezdxf
        from ezdxf.addons import odafc
    except ImportError:
        raise ErrorLectura("La lectura de DWG no está instalada en este programa. Vuelve a crear el .exe.")
    exe = _buscar_oda()
    if exe:
        ezdxf.options.set("odafc-addon", "win_exec_path", exe)
    if not odafc.is_installed():
        raise ErrorLectura("Para leer DWG hace falta «ODA File Converter» (gratuito):\n"
                           "https://www.opendesign.com/guestfiles/oda_file_converter\n\n"
                           "Otra opción: en AutoCAD, «Guardar como» DXF y abrir el DXF.")
    if progreso:
        progreso("Convirtiendo el DWG a DXF…")
    try:
        doc = odafc.readfile(ruta)
    except Exception as e:
        raise ErrorLectura(f"No se ha podido convertir el DWG: {e}")
    return leer_dxf(ruta, progreso, doc=doc)


# --------------------------------------------------------------------------
# Tablas (Excel / CSV) — listados exportados del programa de esquemas
# --------------------------------------------------------------------------
def leer_tabla(ruta):
    """Devuelve las filas (listas de textos) de la primera hoja con datos."""
    ext = os.path.splitext(ruta)[1].lower()
    filas = []
    if ext in (".csv", ".txt"):
        import csv
        datos = None
        for cod in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                with open(ruta, encoding=cod, newline="") as f:
                    datos = f.read()
                break
            except UnicodeDecodeError:
                continue
        if datos is None:
            raise ErrorLectura("No se puede leer el CSV (codificación desconocida).")
        try:
            dialecto = csv.Sniffer().sniff(datos[:5000], delimiters=";,\t|")
        except csv.Error:
            dialecto = csv.excel
            dialecto.delimiter = ";" if datos.count(";") > datos.count(",") else ","
        filas = [[c.strip() for c in fila] for fila in csv.reader(datos.splitlines(), dialecto)]
    elif ext == ".xls":
        raise ErrorLectura("El formato .xls antiguo no se puede leer.\nÁbrelo en Excel y guárdalo como .xlsx.")
    else:
        from openpyxl import load_workbook
        try:
            wb = load_workbook(ruta, read_only=True, data_only=True)
        except Exception as e:
            raise ErrorLectura(f"No se puede abrir el Excel: {e}")
        for ws in wb.worksheets:
            filas = [["" if v is None else str(v).strip() for v in fila] for fila in ws.iter_rows(values_only=True)]
            if any(any(f) for f in filas):
                break
        wb.close()
    # Se conservan las filas vacías intermedias para que «fila N» coincida con el archivo
    while filas and not any(filas[-1]):
        filas.pop()
    if not any(any(f) for f in filas):
        raise ErrorLectura("La tabla está vacía.")
    ancho = max(len(f) for f in filas)
    return [f + [""] * (ancho - len(f)) for f in filas]


def tipo_archivo(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext in EXT_PDF:
        return "pdf"
    if ext in EXT_DXF:
        return "dxf"
    if ext in EXT_DWG:
        return "dwg"
    if ext in EXT_IMAGEN:
        return "imagen"
    if ext in EXT_TABLA:
        return "tabla"
    return None


def leer_paginas(ruta, progreso=None):
    """Lee un esquema (PDF, DXF, DWG o imagen) y devuelve sus páginas."""
    tipo = tipo_archivo(ruta)
    if tipo == "pdf":
        return leer_pdf(ruta, progreso)
    if tipo == "dxf":
        return leer_dxf(ruta, progreso)
    if tipo == "dwg":
        return leer_dwg(ruta, progreso)
    if tipo == "imagen":
        return leer_imagen(ruta, progreso)
    raise ErrorLectura("Formato no reconocido.\nExtensiones admitidas: " + " ".join(extensiones_admitidas()))
