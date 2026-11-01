"""
Extracción de etiquetas a partir de las páginas leídas (lectores.py) o de una
tabla (Excel/CSV).

  - Bornas: páginas "Plano de bornes =...-XALI" -> marcas de cada borna.
  - Elementos: designaciones del esquema (Q100, KM100, K20…) montados en la placa.
"""
import re
from collections import OrderedDict, defaultdict

from lectores import HORIZONTAL, VERTICAL

MAX_DIST_DESCRIPCION = 50       # distancia máx. borna <-> texto descriptivo
PIE = 0.93                      # por debajo de esta altura (fracción) está el cajetín

# Título de las páginas de bornas, en varios idiomas:
#   "Plano de bornes =ALM+ET-XALI"  /  "Plan des bornes =ALM+ET-X1"
#   "Terminal diagram =..."  /  "Klemmenplan =..."
TITULO_RE = re.compile(r"(?:\bbornes|\bterminals?(?: \w+)?|klemmen\w*)\s*=\s*\S+?-(\S+?)\s*$", re.IGNORECASE)
# Palabras del título de las páginas de armario (placa / puerta)
ZONAS = (("Placa", ("PLACA", "PLAQUE", "MOUNTING PLATE", "MONTAGEPLATTE")),
         ("Puerta", ("PUERTA", "PORTE", "DOOR", "TÜR", "TUER")))

# Designación de elemento: letras + número (KM100, Q20, H20.1, ATS1.0, S0.1…)
EQUIPO_RE = re.compile(r"[A-Z]{1,4}\d{1,4}(?:\.\d+)?")
NO_EQUIPO_RE = re.compile(r"^(L\d|N|PE)$")          # potenciales, no elementos
TIPOS = {
    "ATS": "Arrancador suave", "KM": "Contactor", "FC": "Final de carrera",
    "K": "Relé", "Q": "Interruptor / guardamotor", "F": "Protección", "H": "Piloto",
    "S": "Pulsador / selector", "M": "Motor", "U": "Fuente / equipo electrónico",
    "X": "Regleta de bornas",
}
DENTRO_ARMARIO = {"Contactor", "Relé", "Interruptor / guardamotor", "Protección", "Arrancador suave"}
FUERA_O_PUERTA = {"Motor", "Piloto", "Pulsador / selector", "Final de carrera"}


def tipo_equipo(designacion):
    """Tipo según el prefijo de letras (el más largo que coincida)."""
    m = re.match(r"[A-Z]+", designacion.upper())
    letras = m.group() if m else ""
    for pref in sorted(TIPOS, key=len, reverse=True):
        if letras.startswith(pref):
            return TIPOS[pref]
    return "Otros"


def orden_natural(d):
    m = re.match(r"[A-Za-z]+", d)
    return (m.group().upper() if m else "", [float(n) for n in re.findall(r"\d+(?:\.\d+)?", d)], d)


def solo_placa(equipos):
    """Solo los elementos cableados en la placa (KM, Q, K, F…).

    Los de la puerta y los de fuera del armario (motores, finales de carrera…)
    no llevan etiqueta. Si no hay plano de placa, se quitan los tipos que van
    en la puerta o fuera (motores, pilotos, pulsadores, finales de carrera).
    """
    if any(e["ubicacion"] == "Placa" for e in equipos):
        return [e for e in equipos if e["ubicacion"] == "Placa"]
    return [e for e in equipos if e["tipo"] not in FUERA_O_PUERTA]


def _regleta_de(pagina):
    for t in pagina.lineas:
        if t.dir == HORIZONTAL:
            m = TITULO_RE.search(t.texto)
            if m:
                return m.group(1)
    return None


# --------------------------------------------------------------------------
# Bornas
# --------------------------------------------------------------------------
def extraer_bornas(paginas):
    """Lista de dicts: regleta, pos, marca, descripcion, pagina."""
    filas = []
    contador = defaultdict(int)
    for pag in paginas:
        regleta = _regleta_de(pag)
        if not regleta:
            continue  # no es un plano de bornes
        verticales = [t for t in pag.lineas if t.dir == VERTICAL and t.y1 < pag.alto * PIE]
        if not verticales:
            continue
        # Las marcas son la fila de textos verticales más alta (justo encima de
        # la regleta dibujada); el resto, más abajo, son descripciones.
        tam = sorted(t.tam for t in verticales)[len(verticales) // 2]
        if pag.origen == "pdf":
            tol = 2
            y_marcas = min(round(t.y1) for t in verticales)
        else:
            # escaneado/DXF: las marcas no están perfectamente alineadas (papel algo
            # torcido); se toma la fila más alta que tenga bastantes textos
            tol = max(2, tam * 1.2)
            vecinos = {round(t.y1): sum(abs(o.y1 - t.y1) <= tol for o in verticales) for t in verticales}
            mejor = max(vecinos.values())
            y_marcas = min(y for y, n in vecinos.items() if n >= max(2, mejor * 0.5))
        marcas = sorted((t for t in verticales if abs(round(t.y1) - y_marcas) <= tol), key=lambda t: t.x0)
        desc = [((t.x0 + t.x1) / 2, t.texto) for t in verticales if round(t.y1) > y_marcas + 50]
        for t in marcas:
            cx = t.x0 + 5
            cerca = [d for d in desc if abs(d[0] - cx) <= MAX_DIST_DESCRIPCION]
            texto = min(cerca, key=lambda d: abs(d[0] - cx))[1] if cerca else ""
            contador[regleta] += 1
            filas.append({"regleta": regleta, "pos": contador[regleta], "marca": t.texto,
                          "descripcion": texto, "pagina": pag.numero})
    return filas


# --------------------------------------------------------------------------
# Elementos
# --------------------------------------------------------------------------
def extraer_elementos(paginas):
    """Designaciones de elementos (KM100, Q100…) de las páginas de esquema.

    En el esquema las designaciones tienen todas el mismo estilo (mismo color y
    tamaño, texto horizontal). Se toma el estilo con más designaciones
    distintas y se quedan solo los textos que lo cumplen; así se descartan
    números de pin (A1, A2…), cables, referencias de fabricante, etc.
    """
    candidatos = []                       # (texto, tam, estilo, horizontal, pagina, es_ocr)
    zonas = {"Placa": "", "Puerta": ""}   # texto de las páginas de armario
    for pag in paginas:
        pie = pag.texto_zona(pag.alto * PIE)
        cuerpo = pag.texto_zona(0, pag.alto * PIE)
        es_armario = False
        for zona, palabras in ZONAS:
            # sin \b: el OCR a veces quita espacios ("PLACA2016VISION")
            if any(re.search(r"(?<![A-Z])" + p + r"(?![A-Z])", pie, re.IGNORECASE) for p in palabras):
                zonas[zona] += " " + cuerpo
                es_armario = True
        if _regleta_de(pag):
            continue  # los planos de bornes no tienen elementos
        if es_armario and pag.origen == "ocr":
            # En un plano de armario escaneado el OCR confunde designaciones con
            # referencias de fabricante (LC1D09BL, GV2ME06…): solo sirve para la ubicación
            continue
        for t in pag.trozos:
            if (t.y0 + t.y1) / 2 > pag.alto * PIE:      # cajetín
                continue
            txt = t.texto.strip()
            # El OCR a veces junta designaciones vecinas: "Q21Q22Q23" -> Q21, Q22, Q23
            # (solo si todas empiezan por la misma letra; "GV2ME06" es una referencia)
            partes = [txt]
            if pag.origen == "ocr":
                trozos = [p for p in re.split(r"(?<=\d)(?=[A-Z])", txt) if EQUIPO_RE.fullmatch(p)]
                if len(trozos) > 1 and len({p[0] for p in trozos}) == 1:
                    partes = trozos
                # "Q100H": el OCR pega una letra del símbolo de al lado
                m = re.fullmatch(r"([A-Z]{1,4}\d{1,4}(?:\.\d+)?)[A-Z]", txt)
                if m:
                    partes = [m.group(1)]
            for parte in partes:
                if EQUIPO_RE.fullmatch(parte) and not NO_EQUIPO_RE.match(parte):
                    candidatos.append((parte, t.tam, t.estilo, t.dir == HORIZONTAL, pag.numero, pag.origen == "ocr"))
    if not candidatos:
        return []

    # Estilo de referencia = el que tiene más designaciones distintas (texto
    # horizontal). Se calcula por separado para páginas con texto y para páginas
    # escaneadas (OCR), donde el tamaño es aproximado y no hay color.
    paginas_de = defaultdict(set)
    for es_ocr in (False, True):
        grupo = [c for c in candidatos if c[5] == es_ocr]
        estilos = defaultdict(set)
        for txt, tam, est, horiz, _, _ in grupo:
            if horiz:
                estilos[("ocr", round(tam / 2)) if es_ocr else (round(tam, 1), _hashable(est))].add(txt)
        if not estilos:
            continue
        ref = max(estilos, key=lambda e: len(estilos[e]))
        if es_ocr:
            # En una página escaneada no hay color ni tamaño fiable: se aceptan
            # todas las designaciones y se quitan cables (U101/V101/W101) y
            # bornes de bobina (A1, A2).
            textos = {c[0] for c in grupo}
            for txt, tam, est, horiz, num, _ in grupo:
                if horiz and not _parece_cable(txt, textos) and not re.fullmatch(r"A[12]", txt):
                    paginas_de[txt].add(num)
            continue
        for txt, tam, est, _, num, _ in grupo:
            if _hashable(est) == ref[1] and abs(tam - ref[0]) <= 0.3:
                paginas_de[txt].add(num)
    if not paginas_de:
        return []

    def esta_en(d, texto):
        # KM100 no debe casar con M100, ni Q1 con Q100 / Q1.1
        return re.search(r"(?<![A-Z])" + re.escape(d) + r"(?![0-9]|\.[0-9])", texto) is not None

    def ubicacion(d):
        if esta_en(d, zonas["Placa"]):
            return "Placa"
        if esta_en(d, zonas["Puerta"]):
            return "Puerta"
        if not (zonas["Placa"] or zonas["Puerta"]):
            return "Sin plano de armario"
        # KM20.1 = segundo contactor del inversor KM20: va donde va KM20
        base = d.split(".")[0]
        if base != d and esta_en(base, zonas["Placa"]):
            return "Placa"
        # Estos elementos siempre van dentro del armario
        if tipo_equipo(d) in DENTRO_ARMARIO:
            return "Placa"
        return "Fuera del armario"

    return [{"designacion": d, "tipo": tipo_equipo(d), "ubicacion": ubicacion(d), "paginas": sorted(paginas_de[d])}
            for d in sorted(paginas_de, key=orden_natural)]


def _parece_cable(txt, textos):
    """U101 junto a V101 y W101 (o PE101) es un número de cable/borna, no un elemento."""
    m = re.fullmatch(r"(U|V|W|PE)(\d+)", txt)
    if not m:
        return False
    n = m.group(2)
    return sum(f"{f}{n}" in textos for f in ("U", "V", "W", "PE")) >= 2


def _hashable(v):
    return tuple(v) if isinstance(v, list) else v


# --------------------------------------------------------------------------
# Tablas (Excel / CSV)
# --------------------------------------------------------------------------
# Palabras de cabecera habituales en los listados de EPLAN, SEE, AutoCAD Electrical…
CLAVES = {
    "regleta": ("regleta", "bornero", "bornier", "terminal strip", "strip", "klemmleiste", "regleta de bornes"),
    "borna": ("borna", "borne", "terminal", "klemme", "marca", "conexión", "conexion", "nº borna", "n° borne"),
    "designacion": ("designación", "designacion", "aparato", "equipo", "elemento", "dispositivo", "device",
                    "tag", "identificador", "bmk", "repère", "repere", "référence d'appareil"),
    "descripcion": ("descripción", "descripcion", "función", "funcion", "description", "fonction",
                    "texto", "comentario", "beschreibung"),
}


def _puntuar(celda, campo):
    c = celda.lower().strip()
    if not c:
        return 0
    for k in CLAVES[campo]:
        if c == k:
            return 3
    return 1 if any(k in c for k in CLAVES[campo]) else 0


def analizar_tabla(filas):
    """Adivina la fila de cabecera, las columnas y si es un listado de bornas o de elementos.

    Devuelve dict: cabecera (índice de fila o None), columnas {campo: índice o None}, modo.
    """
    mejor = (0, None)
    for i, fila in enumerate(filas[:15]):
        p = sum(max((_puntuar(c, campo) for campo in CLAVES), default=0) for c in fila)
        if p > mejor[0]:
            mejor = (p, i)
    cab = mejor[1] if mejor[0] >= 2 else None
    columnas = {}
    if cab is not None:
        usadas = set()
        for campo in ("regleta", "designacion", "borna", "descripcion"):
            punt = [(_puntuar(c, campo), j) for j, c in enumerate(filas[cab]) if j not in usadas]
            p, j = max(punt, default=(0, None))
            columnas[campo] = j if p else None
            if p:
                usadas.add(j)
    else:
        columnas = {k: None for k in CLAVES}
    datos = [f for f in (filas[cab + 1:] if cab is not None else filas) if any(f)]
    # Sin cabecera: buscar una columna con marcas tipo "X1:5" o designaciones tipo KM100
    if columnas.get("borna") is None and columnas.get("designacion") is None and datos:
        n_col = len(datos[0])
        for j in range(n_col):
            vals = [f[j] for f in datos[:200] if f[j]]
            if not vals:
                continue
            if sum(bool(re.fullmatch(r"\S+:\S+", v)) for v in vals) > len(vals) * 0.6:
                columnas["borna"] = j
                break
            if sum(bool(EQUIPO_RE.search(_ultima_parte(v))) for v in vals) > len(vals) * 0.6:
                columnas["designacion"] = j
    modo = "bornas" if columnas.get("borna") is not None and (
        columnas.get("regleta") is not None or columnas.get("designacion") is None) else "elementos"
    return {"cabecera": cab, "columnas": columnas, "modo": modo}


def _ultima_parte(v):
    """'=ALM+ET-KM100' -> 'KM100';  '-Q1' -> 'Q1'."""
    v = v.strip()
    return re.split(r"[-=+]", v)[-1].strip() if re.search(r"[-=+]", v) else v


def filas_desde_tabla(filas, cabecera, columnas, modo):
    """Convierte la tabla en filas de bornas o de elementos."""
    datos = filas[cabecera + 1:] if cabecera is not None else filas

    def col(fila, campo):
        j = columnas.get(campo)
        return fila[j].strip() if j is not None and j < len(fila) else ""

    if modo == "bornas":
        res, contador = [], defaultdict(int)
        for n, f in enumerate(datos, start=(-1 if cabecera is None else cabecera) + 2):
            if not any(f):
                continue
            marca, regleta = col(f, "borna"), col(f, "regleta")
            if not marca:
                continue
            if not regleta and ":" in marca:          # "X1:5" -> regleta X1, borna 5
                regleta, marca = marca.split(":", 1)
            regleta = _ultima_parte(regleta) or "—"
            contador[regleta] += 1
            res.append({"regleta": regleta, "pos": contador[regleta], "marca": marca,
                        "descripcion": col(f, "descripcion"), "pagina": f"fila {n}"})
        return res

    vistos = OrderedDict()
    for n, f in enumerate(datos, start=(-1 if cabecera is None else cabecera) + 2):
        d = _ultima_parte(col(f, "designacion"))
        if d and d not in vistos:
            vistos[d] = n
    return [{"designacion": d, "tipo": tipo_equipo(d), "ubicacion": "Sin plano de armario", "paginas": [f"fila {n}"]}
            for d, n in sorted(vistos.items(), key=lambda kv: orden_natural(kv[0]))]


# --------------------------------------------------------------------------
# Número del esquema (para separar las etiquetas de cada esquema)
# --------------------------------------------------------------------------
ETIQ_NUMERO_RE = re.compile(r"N\s*[º°oO]\.?\s*(?:DE\s+INSTALACI[ÓO]N|D'INSTALLATION|DE\s+PROYECTO|PROYECTO|PROJECT)",
                            re.IGNORECASE)
VALOR_NUMERO_RE = re.compile(r"[A-Z0-9][A-Z0-9_\-./]{3,}", re.IGNORECASE)


def numero_esquema(paginas):
    """Número de instalación/proyecto del cajetín ("Nº DE INSTALACIÓN: CECE00000000001"), o None."""
    for pag in paginas[:5]:
        for t in pag.trozos:
            m = ETIQ_NUMERO_RE.search(t.texto)
            if not m:
                continue
            # el número puede ir en el mismo texto ("...: CECE…") o en el texto de su derecha
            resto = t.texto[m.end():].strip(" :")
            if VALOR_NUMERO_RE.fullmatch(resto or "-"):
                return resto
            alto = max(t.y1 - t.y0, 1)
            derecha = [o for o in pag.trozos if o is not t and o.x0 >= t.x1 - 1
                       and abs((o.y0 + o.y1) / 2 - (t.y0 + t.y1) / 2) < alto * 0.6]
            for o in sorted(derecha, key=lambda o: o.x0):
                if VALOR_NUMERO_RE.fullmatch(o.texto.strip()):
                    return o.texto.strip()
    return None


# --------------------------------------------------------------------------
# Copias por esquema
# --------------------------------------------------------------------------
def multiplicar(filas, copias):
    """Repite las etiquetas de cada esquema tantas veces como diga `copias` {esquema: n},
    seguidas: esquema A copia 1, A copia 2, A copia 3, esquema B copia 1…
    Cada fila resultante lleva "copia" = "k/n"."""
    orden = []
    por_esquema = {}
    for f in filas:
        e = f.get("esquema", "")
        if e not in por_esquema:
            orden.append(e)
            por_esquema[e] = []
        por_esquema[e].append(f)
    salida = []
    for e in orden:
        n = max(1, int(copias.get(e, 1) or 1))
        for k in range(1, n + 1):
            salida += [dict(f, copia=f"{k}/{n}") for f in por_esquema[e]]
    return salida

