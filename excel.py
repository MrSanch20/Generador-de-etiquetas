"""
Guardado de las etiquetas en Excel, listo para importar en GeniusPro (Cembre).
"""
from collections import OrderedDict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

CAB_FONT = Font(bold=True, color="FFFFFF")
CAB_FILL = PatternFill("solid", fgColor="1D4FC4")
BANDA_FILL = PatternFill("solid", fgColor="EEF1F5")     # fondo de un esquema sí, otro no


def _bandas(ws, col_esquema, col_copia):
    """Fondo alterno por bloque (esquema + copia): se ve de un vistazo dónde empieza y
    acaba cada esquema y cada copia. La primera fila de cada bloque va en negrita."""
    anterior, banda = None, False
    for fila in ws.iter_rows(min_row=1):
        bloque = (fila[col_esquema].value, fila[col_copia].value)
        if bloque != anterior:
            banda = not banda if anterior is not None else False
            anterior = bloque
            fila[col_esquema].font = Font(bold=True)
            fila[col_copia].font = Font(bold=True)
        if banda:
            for c in fila:
                c.fill = BANDA_FILL


def _anchos(ws, anchos):
    """Solo ancho de columnas, SIN fila de cabecera: los datos empiezan en la fila 1
    (así GeniusPro importa desde la primera línea sin saltar nada)."""
    from openpyxl.utils import get_column_letter
    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _cabecera(ws, cols, anchos):
    ws.append(cols)
    for c, w in zip(ws[1], anchos):
        c.font, c.fill = CAB_FONT, CAB_FILL
        ws.column_dimensions[c.column_letter].width = w
    ws.freeze_panes = "A2"


def _libro():
    wb = Workbook()
    wb.remove(wb.active)           # sin la hoja vacía por defecto
    return wb


def _hoja(wb, titulo):
    return wb.create_sheet(titulo)


def _hoja_detalle_bornas(wb, filas, titulo="Bornas"):
    ws = _hoja(wb, titulo)
    # columnas: A Esquema | B Copia | C Marca (etiqueta) | D SW-CODE | E Regleta | F Pos. | G Descripción | H Página
    _anchos(ws, [22, 7, 12, 11, 12, 7, 40, 11])
    for f in filas:
        ws.append([f.get("esquema", ""), f.get("copia", "1/1"), f["marca"], f.get("swcode", ""), f["regleta"],
                   f["pos"], f["descripcion"], f["pagina"]])
    _bandas(ws, 0, 1)


def _hoja_detalle_elementos(wb, elementos, titulo="Elementos"):
    ws = _hoja(wb, titulo)
    # columnas: A Esquema | B Copia | C Designación (etiqueta) | D SW-CODE | E Tipo | F Páginas/filas
    _anchos(ws, [22, 7, 14, 11, 28, 18])
    for e in elementos:
        ws.append([e.get("esquema", ""), e.get("copia", "1/1"), e["designacion"], e.get("swcode", ""), e["tipo"],
                   ", ".join(map(str, e["paginas"]))])
    _bandas(ws, 0, 1)


def _hoja_geniuspro(wb, filas, titulo="GeniusPro"):
    """Una etiqueta por fila, lista para importar: texto y SW-CODE del consumible."""
    ws = _hoja(wb, titulo)
    # columnas: A Esquema | B Copia | C Texto (etiqueta) | D SW-CODE
    # (la etiqueta va en la columna C en TODAS las hojas: en GeniusPro, Valor inicial = [C])
    _anchos(ws, [22, 7, 14, 11])
    for texto, swcode, esquema, copia in filas:
        ws.append([esquema, copia, texto, swcode])
    _bandas(ws, 0, 1)


def _hoja_resumen(wb, grupos, titulo_col, titulo_n, swcodes=(), titulo="Resumen"):
    ws = _hoja(wb, titulo)
    _cabecera(ws, [titulo_col, titulo_n], [28, 13])
    cuenta = OrderedDict()
    for g in grupos:
        cuenta[g] = cuenta.get(g, 0) + 1
    for g, n in cuenta.items():
        ws.append([g, n])
    ws.append(["TOTAL", f"=SUM(B2:B{len(cuenta) + 1})"])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    # Etiquetas por consumible (SW-CODE): cuántas hay que imprimir en cada uno
    if swcodes:
        ws.append([])
        ws.append(["SW-CODE", "Nº etiquetas"])
        for c in ws[ws.max_row]:
            c.font, c.fill = CAB_FONT, CAB_FILL
        por_codigo = OrderedDict()
        for c in swcodes:
            por_codigo[c or "(sin SW-CODE)"] = por_codigo.get(c or "(sin SW-CODE)", 0) + 1
        for c, n in por_codigo.items():
            ws.append([c, n])


def guardar_excel_bornas(filas, salida):
    """Solo bornas: detalle, columna para GeniusPro y resumen por regleta."""
    wb = _libro()
    _hoja_geniuspro(wb, [(f["marca"], f.get("swcode", ""), f.get("esquema", ""), f.get("copia", "1/1")) for f in filas])
    _hoja_detalle_bornas(wb, filas)
    _hoja_resumen(wb, [f["regleta"] for f in filas], "Regleta", "Nº bornas", [f.get("swcode", "") for f in filas])
    wb.save(salida)


def guardar_excel_elementos(elementos, salida):
    """Solo elementos (Q, KM, K, F…): detalle, columna para GeniusPro y resumen por tipo."""
    wb = _libro()
    _hoja_geniuspro(wb, [(e["designacion"], e.get("swcode", ""), e.get("esquema", ""), e.get("copia", "1/1")) for e in elementos])
    _hoja_detalle_elementos(wb, elementos)
    _hoja_resumen(wb, [e["tipo"] for e in elementos], "Tipo", "Nº elementos", [e.get("swcode", "") for e in elementos])
    wb.save(salida)


def guardar_excel_todo(filas, elementos, salida):
    """Bornas y elementos en el mismo libro, cada uno con su hoja para GeniusPro."""
    wb = _libro()
    _hoja_geniuspro(wb, [(f["marca"], f.get("swcode", ""), f.get("esquema", ""), f.get("copia", "1/1")) for f in filas], "GeniusPro bornas")
    _hoja_geniuspro(wb, [(e["designacion"], e.get("swcode", ""), e.get("esquema", ""), e.get("copia", "1/1")) for e in elementos], "GeniusPro elementos")
    _hoja_detalle_bornas(wb, filas)
    _hoja_detalle_elementos(wb, elementos)
    _hoja_resumen(wb, ["Bornas"] * len(filas) + ["Elementos"] * len(elementos), "Etiquetas", "Cantidad",
                  [f.get("swcode", "") for f in filas] + [e.get("swcode", "") for e in elementos])
    wb.save(salida)
