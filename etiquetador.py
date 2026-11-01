"""
Etiquetador de cuadros

Lee un esquema eléctrico y saca las etiquetas para la impresora Cembre (GeniusPro):
  - las marcas de las bornas (páginas "Plano de bornes" / "Plan des bornes"…)
  - las designaciones de los elementos cableados en la placa (Q100, KM100, K20…)

Formatos: PDF (con texto o escaneado), DXF, DWG, imágenes escaneadas
(PNG/JPG/TIF) y listados en Excel/CSV exportados del programa de esquemas.

Uso:
  - Doble clic en el .exe y pulsar "Abrir esquema…", o
  - Arrastrar el archivo encima del .exe.
"""
import os
import sys
import threading
import tkinter as tk
from collections import OrderedDict
from tkinter import filedialog, messagebox, ttk

import config
import extraccion as X
import lectores as L
from excel import guardar_excel_bornas, guardar_excel_elementos, guardar_excel_todo

APP_NAME = "Etiquetador de cuadros"


# --------------------------------------------------------------------------
# Ventana
# --------------------------------------------------------------------------
class Lista(ttk.Frame):
    """Pestaña con filtros a la izquierda (casillas) y tabla a la derecha."""

    def __init__(self, master, titulo_filtro, columnas, al_cambiar, al_poner_swcode=None):
        super().__init__(master, padding=(0, 8, 0, 0))
        self.al_cambiar = al_cambiar
        self.al_poner_swcode = al_poner_swcode or (lambda filas: None)
        self.filas = []
        self.clave = None            # función fila -> grupo (regleta / tipo)
        self.valores = None          # función fila -> tupla de columnas
        self.texto = None            # función fila -> texto de la etiqueta
        self.es_pe = lambda f: False
        self.incluir = OrderedDict()

        izq = ttk.LabelFrame(self, text=titulo_filtro, padding=8)
        izq.pack(side="left", fill="y", padx=(0, 10))
        self.frame_filtros = ttk.Frame(izq)
        self.frame_filtros.pack(fill="y", expand=True)
        botones = ttk.Frame(izq)
        botones.pack(fill="x", pady=(8, 0))
        ttk.Button(botones, text="Todas", width=8, command=lambda: self._marcar(True)).pack(side="left")
        ttk.Button(botones, text="Ninguna", width=8, command=lambda: self._marcar(False)).pack(side="left", padx=(6, 0))

        der = ttk.Frame(self)
        der.pack(side="left", fill="both", expand=True)

        # SW-CODE: referencia del consumible donde se imprime la etiqueta
        sw = ttk.Frame(der)
        sw.pack(fill="x", pady=(0, 6))
        ttk.Label(sw, text="SW-CODE:").pack(side="left")
        self.swcode = tk.StringVar()
        entrada = ttk.Entry(sw, textvariable=self.swcode, width=10)
        entrada.pack(side="left", padx=(6, 6))
        entrada.bind("<Return>", lambda _e: self.poner_swcode(solo_seleccion=True))
        ttk.Button(sw, text="Aplicar a seleccionadas",
                   command=lambda: self.poner_swcode(solo_seleccion=True)).pack(side="left")
        ttk.Button(sw, text="Aplicar a todas las visibles",
                   command=lambda: self.poner_swcode(solo_seleccion=False)).pack(side="left", padx=(6, 0))
        ttk.Label(sw, text="(Ctrl+E selecciona todo lo visible)", foreground="#5d6874").pack(side="left", padx=(10, 0))

        self.tabla = ttk.Treeview(der, columns=[c[0] for c in columnas], show="headings", selectmode="extended")
        self.tabla.bind("<Control-e>", lambda _e: self.tabla.selection_set(self.tabla.get_children()))
        self.tabla.bind("<Control-a>", lambda _e: self.tabla.selection_set(self.tabla.get_children()))
        for c, t, w, anchor, stretch in columnas:
            self.tabla.heading(c, text=t)
            self.tabla.column(c, width=w, anchor=anchor, stretch=stretch)
        self.tabla.tag_configure("pe", background="#eef7e6")
        self.tabla.tag_configure("off", foreground="#9aa3ad")
        sb = ttk.Scrollbar(der, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=sb.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def cargar(self, filas, activos=None):
        self.filas = filas
        antes = {g: v.get() for g, v in self.incluir.items()}   # se respetan las casillas ya tocadas
        for w in self.frame_filtros.winfo_children():
            w.destroy()
        self.incluir = OrderedDict()
        cuenta = OrderedDict()
        for f in filas:
            cuenta[self.clave(f)] = cuenta.get(self.clave(f), 0) + 1
        for grupo, n in cuenta.items():
            var = tk.BooleanVar(value=antes.get(grupo, activos is None or grupo in activos))
            var.trace_add("write", lambda *_: self.al_cambiar())
            self.incluir[grupo] = var
            ttk.Checkbutton(self.frame_filtros, text=f"{grupo}  ({n})", variable=var).pack(anchor="w", pady=1)

    def incluidas(self):
        return [f for f in self.filas if self.incluir[self.clave(f)].get()]

    def _marcar(self, valor):
        for v in self.incluir.values():
            v.set(valor)

    def pintar(self, q):
        self.tabla.delete(*self.tabla.get_children())
        for i, f in enumerate(self.filas):
            vals = self.valores(f)
            if q and not any(q in str(v).lower() for v in vals):
                continue
            tags = []
            if self.es_pe(f):
                tags.append("pe")
            if not self.incluir[self.clave(f)].get():
                tags.append("off")
            self.tabla.insert("", "end", iid=str(i), values=vals, tags=tags)

    def poner_swcode(self, solo_seleccion):
        """Pone el SW-CODE escrito en las filas seleccionadas o en todas las visibles."""
        codigo = self.swcode.get().strip()
        if solo_seleccion:
            ids = self.tabla.selection()
            if not ids:
                messagebox.showinfo(APP_NAME, "Selecciona primero las filas en la tabla "
                                              "(Mayús/Ctrl + clic, o Ctrl+E para todas las visibles).")
                return
        else:
            ids = self.tabla.get_children()
        filas = [self.filas[int(i)] for i in ids]
        for f in filas:
            f["swcode"] = codigo
        self.al_poner_swcode(filas)
        self.al_cambiar()
        # mantener la selección tras repintar
        self.tabla.selection_set([i for i in ids if self.tabla.exists(i)])


class DialogoTabla(tk.Toplevel):
    """Pregunta qué contiene un listado Excel/CSV y en qué columnas está cada dato."""

    NINGUNA = "(ninguna)"

    def __init__(self, master, nombre, filas, propuesta):
        super().__init__(master)
        self.title(f"Listado: {nombre}")
        self.transient(master)
        self.resizable(True, True)
        self.filas = filas
        self.resultado = None
        n_col = len(filas[0])
        self.modo = tk.StringVar(value=propuesta["modo"])
        self.fila_cab = tk.IntVar(value=(propuesta["cabecera"] + 1) if propuesta["cabecera"] is not None else 0)

        marco = ttk.Frame(self, padding=14)
        marco.pack(fill="both", expand=True)
        ttk.Label(marco, text="¿Qué contiene este listado?", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Radiobutton(marco, text="Bornas (regleta y borna)", value="bornas", variable=self.modo,
                        command=self._actualizar).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Radiobutton(marco, text="Elementos (KM100, Q100…)", value="elementos", variable=self.modo,
                        command=self._actualizar).grid(row=1, column=2, columnspan=2, sticky="w", pady=(4, 0))

        ttk.Label(marco, text="Fila de cabecera (0 = no tiene):").grid(row=2, column=0, sticky="w", pady=(10, 0))
        sp = ttk.Spinbox(marco, from_=0, to=min(30, len(filas)), width=5, textvariable=self.fila_cab,
                         command=self._actualizar)
        sp.grid(row=2, column=1, sticky="w", pady=(10, 0))
        sp.bind("<KeyRelease>", lambda _e: self._actualizar())

        self.combos = {}
        self.etiquetas = {}
        campos = [("regleta", "Regleta:"), ("borna", "Borna / marca:"), ("designacion", "Designación del elemento:"),
                  ("descripcion", "Descripción (opcional):")]
        for i, (campo, texto) in enumerate(campos):
            self.etiquetas[campo] = ttk.Label(marco, text=texto)
            self.etiquetas[campo].grid(row=3 + i, column=0, sticky="w", pady=2)
            cb = ttk.Combobox(marco, state="readonly", width=34)
            cb.grid(row=3 + i, column=1, columnspan=3, sticky="w", pady=2)
            self.combos[campo] = cb
        self.propuesta_cols = propuesta["columnas"]

        ttk.Label(marco, text="Primeras filas:").grid(row=7, column=0, sticky="w", pady=(10, 2))
        vista = ttk.Frame(marco)
        vista.grid(row=8, column=0, columnspan=4, sticky="nsew")
        marco.rowconfigure(8, weight=1)
        marco.columnconfigure(3, weight=1)
        cols = [self._letra(j) for j in range(n_col)]
        self.tabla = ttk.Treeview(vista, columns=cols, show="headings", height=8)
        for c in cols:
            self.tabla.heading(c, text=c)
            self.tabla.column(c, width=90, stretch=False)
        sx = ttk.Scrollbar(vista, orient="horizontal", command=self.tabla.xview)
        self.tabla.configure(xscrollcommand=sx.set)
        self.tabla.pack(fill="both", expand=True)
        sx.pack(fill="x")
        for f in filas[:12]:
            self.tabla.insert("", "end", values=f)

        self.lbl_aviso = ttk.Label(marco, text="", foreground="#b42318")
        self.lbl_aviso.grid(row=9, column=0, columnspan=4, sticky="w", pady=(8, 0))
        bots = ttk.Frame(marco)
        bots.grid(row=10, column=0, columnspan=4, sticky="e", pady=(10, 0))
        ttk.Button(bots, text="Cancelar", command=self.destroy).pack(side="right")
        ttk.Button(bots, text="Cargar", command=self._aceptar).pack(side="right", padx=(0, 8))

        self._actualizar(primera_vez=True)
        self.geometry("720x520")
        self.grab_set()

    @staticmethod
    def _letra(j):
        s = ""
        j += 1
        while j:
            j, r = divmod(j - 1, 26)
            s = chr(65 + r) + s
        return s

    def _opciones(self):
        try:
            cab = int(self.fila_cab.get())
        except (tk.TclError, ValueError):
            cab = 0
        nombres = self.filas[cab - 1] if 0 < cab <= len(self.filas) else [""] * len(self.filas[0])
        return [self.NINGUNA] + [f"{self._letra(j)}  {n}".rstrip() for j, n in enumerate(nombres)]

    def _actualizar(self, primera_vez=False):
        ops = self._opciones()
        for campo, cb in self.combos.items():
            actual = cb.current() if not primera_vez else (
                (self.propuesta_cols.get(campo) + 1) if self.propuesta_cols.get(campo) is not None else 0)
            cb["values"] = ops
            cb.current(max(0, min(actual, len(ops) - 1)))
        bornas = self.modo.get() == "bornas"
        for campo in ("regleta", "borna"):
            estado = "readonly" if bornas else "disabled"
            self.combos[campo].configure(state=estado)
        self.combos["designacion"].configure(state="disabled" if bornas else "readonly")

    def _aceptar(self):
        try:
            cab = int(self.fila_cab.get())
        except (tk.TclError, ValueError):
            cab = 0
        columnas = {c: (cb.current() - 1 if cb.current() > 0 else None) for c, cb in self.combos.items()}
        modo = self.modo.get()
        if modo == "bornas" and columnas["borna"] is None:
            self.lbl_aviso.configure(text="Elige la columna de la borna / marca.")
            return
        if modo == "elementos" and columnas["designacion"] is None:
            self.lbl_aviso.configure(text="Elige la columna de la designación del elemento.")
            return
        self.resultado = ((cab - 1) if cab > 0 else None, columnas, modo)
        self.destroy()


class App(tk.Tk):
    def __init__(self, pdf_inicial=None):
        super().__init__()
        self.cfg = config.cargar()
        self.title(APP_NAME)
        self.geometry("1040x660")
        self.minsize(780, 480)
        self.pdf_path = None
        self.busqueda = tk.StringVar()
        self.busqueda.trace_add("write", lambda *_: self._pintar())

        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Big.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("Num.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Treeview", rowheight=24)

        self._construir()
        if pdf_inicial:
            self.after(200, lambda: self.abrir(pdf_inicial))

    def _construir(self):
        top = ttk.Frame(self, padding=(14, 12, 14, 6))
        top.pack(fill="x")
        ttk.Label(top, text=APP_NAME, style="Big.TLabel").pack(side="left")
        self.btn_limpiar = ttk.Button(top, text="Limpiar todo", command=self.limpiar_todo)
        self.btn_limpiar.pack(side="right")
        self.btn_anadir = ttk.Button(top, text="Añadir esquema…", command=lambda: self.elegir_archivo(anadir=True))
        self.btn_anadir.pack(side="right", padx=(0, 8))
        self.btn_abrir = ttk.Button(top, text="Abrir esquema…", command=self.elegir_archivo)
        self.btn_abrir.pack(side="right", padx=(0, 8))
        ttk.Button(top, text="Formatos admitidos", command=self.mostrar_formatos).pack(side="right", padx=(0, 8))

        # Extensiones que se pueden abrir, siempre a la vista
        ttk.Label(self, text="Formatos admitidos:  " + "  ".join(L.extensiones_admitidas()),
                  foreground="#5d6874", padding=(14, 0, 14, 4)).pack(fill="x")

        info = ttk.Frame(self, padding=(14, 0, 14, 6))
        info.pack(fill="x")
        self.lbl_pdf = ttk.Label(info, text="Abre uno o varios esquemas, o arrástralos encima del programa.")
        self.lbl_pdf.pack(side="left")
        self.lbl_tags = ttk.Label(info, text="", style="Num.TLabel")
        self.lbl_tags.pack(side="right")

        # Copias por esquema: cuántas veces se imprimen las etiquetas de cada uno
        self.copias = {}                       # esquema -> StringVar con el nº de copias
        self.frame_copias = ttk.Frame(self, padding=(14, 0, 14, 6))
        self.frame_copias.pack(fill="x")

        barra = ttk.Frame(self, padding=(14, 0, 14, 0))
        barra.pack(fill="x")
        ttk.Label(barra, text="Buscar:").pack(side="left")
        ttk.Entry(barra, textvariable=self.busqueda, width=30).pack(side="left", padx=(6, 0))

        self.pestanas = ttk.Notebook(self, padding=(14, 6, 14, 0))
        self.pestanas.pack(fill="both", expand=True)

        self.bornas = Lista(self.pestanas, "Regletas a incluir", [
            ("esquema", "Esquema", 150, "w", False), ("regleta", "Regleta", 80, "w", False), ("pos", "Pos.", 50, "e", False),
            ("marca", "Marca", 100, "w", False), ("descripcion", "Descripción (orientativa)", 320, "w", True),
            ("pagina", "Pág. / fila", 70, "e", False), ("swcode", "SW-CODE", 80, "w", False)],
            self._pintar, self._recordar_swcode_bornas)
        # con varios esquemas, la misma regleta (XALI) de cada esquema se filtra por separado
        self.bornas.clave = lambda f: f'{f["regleta"]} · {f["esquema"]}' if self._varios() else f["regleta"]
        self.bornas.valores = lambda f: (f["esquema"], f["regleta"], f["pos"], f["marca"], f["descripcion"], f["pagina"],
                                         f.get("swcode", ""))
        self.bornas.texto = lambda f: f["marca"]
        self.bornas.es_pe = lambda f: f["marca"].upper().startswith("PE")

        self.equipos = Lista(self.pestanas, "Tipos a incluir", [
            ("esquema", "Esquema", 150, "w", False),
            ("designacion", "Designación", 110, "w", False), ("tipo", "Tipo", 220, "w", False),
            ("paginas", "Páginas / filas", 200, "w", True), ("swcode", "SW-CODE", 80, "w", False)],
            self._pintar, self._recordar_swcode_elementos)
        self.equipos.clave = lambda e: e["tipo"]
        self.equipos.valores = lambda e: (e["esquema"], e["designacion"], e["tipo"], ", ".join(map(str, e["paginas"])),
                                          e.get("swcode", ""))
        self.equipos.texto = lambda e: e["designacion"]

        # Pestaña "Todo": bornas y elementos juntos (usa los filtros de las otras dos)
        self.todo = ttk.Frame(self.pestanas, padding=(0, 8, 0, 0))
        ttk.Label(self.todo, text="Bornas y elementos juntos. Usa los filtros de las pestañas Bornas y Elementos.",
                  foreground="#5d6874").pack(anchor="w", pady=(0, 6))
        tf = ttk.Frame(self.todo)
        tf.pack(fill="both", expand=True)
        cols = [("esquema", "Esquema", 150, "w", False),
                ("clase", "Etiqueta", 90, "w", False), ("texto", "Texto", 110, "w", False),
                ("grupo", "Regleta / tipo", 190, "w", False), ("detalle", "Descripción", 260, "w", True),
                ("pag", "Pág.", 90, "w", False), ("swcode", "SW-CODE", 80, "w", False)]
        self.tabla_todo = ttk.Treeview(tf, columns=[c[0] for c in cols], show="headings")
        for c, t, w, anchor, stretch in cols:
            self.tabla_todo.heading(c, text=t)
            self.tabla_todo.column(c, width=w, anchor=anchor, stretch=stretch)
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.tabla_todo.yview)
        self.tabla_todo.configure(yscrollcommand=sb.set)
        self.tabla_todo.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.pestanas.add(self.bornas, text="Bornas")
        self.pestanas.add(self.equipos, text="Elementos")
        self.pestanas.add(self.todo, text="Todo")
        self.pestanas.bind("<<NotebookTabChanged>>", lambda _e: self._pintar())

        bot = ttk.Frame(self, padding=(14, 10, 14, 12))
        bot.pack(fill="x")
        self.btn_excel = ttk.Button(bot, text="Guardar Excel…", state="disabled", command=self.guardar, width=28)
        self.btn_excel.pack(side="right")
        self.btn_copiar = ttk.Button(bot, text="Copiar esta lista", command=self.copiar, state="disabled")
        self.btn_copiar.pack(side="right", padx=(0, 8))
        self.lbl_estado = ttk.Label(bot, text="Bornas: páginas «Plano de bornes». Elementos: los cableados en la placa (Q, KM, K, F…).",
                                    foreground="#5d6874")
        self.lbl_estado.pack(side="left")

    # ---- acciones ----
    def mostrar_formatos(self):
        texto = "\n\n".join(f"{nombre}  ({'  '.join(exts)})\n{nota}" for nombre, exts, nota in L.FORMATOS)
        texto += ("\n\nOCR en este programa: " +
                  ("SÍ incluido." if L.ocr_disponible() else
                   "NO incluido (no lee planos escaneados ni imágenes). Vuelve a crear el .exe "
                   "con crear_exe.bat usando Python 3.12 de 64 bits."))
        messagebox.showinfo("Formatos admitidos", texto)

    def elegir_archivo(self, anadir=False):
        rutas = filedialog.askopenfilenames(
            title="Añadir esquemas" if anadir else "Elegir esquemas (puedes elegir varios)", filetypes=[
                ("Esquemas y listados", "*.pdf *.dxf *.dwg *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.xlsx *.xlsm *.csv *.txt"),
                ("PDF", "*.pdf"), ("AutoCAD (DXF, DWG)", "*.dxf *.dwg"),
                ("Imagen escaneada", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp"),
                ("Listado Excel / CSV", "*.xlsx *.xlsm *.csv *.txt"), ("Todos", "*.*")])
        if rutas:
            self.abrir(list(rutas), anadir=anadir)

    # ---- carga de uno o varios esquemas ----
    @staticmethod
    def _nombre(ruta):
        return os.path.splitext(os.path.basename(ruta))[0]

    def _esquemas(self):
        vistos = []
        for f in self.bornas.filas + self.equipos.filas:
            if f["esquema"] not in vistos:
                vistos.append(f["esquema"])
        return vistos

    def _varios(self):
        return len(self._esquemas()) > 1

    def abrir(self, rutas, anadir=False):
        """Abre uno o varios archivos. anadir=False empieza de cero; True los suma a los cargados."""
        if isinstance(rutas, str):
            rutas = [rutas]
        malos = [r for r in rutas if L.tipo_archivo(r) is None]
        if malos:
            messagebox.showwarning(APP_NAME, "Formato no reconocido: " + ", ".join(map(os.path.basename, malos))
                                   + "\n\nExtensiones admitidas:\n" + "  ".join(L.extensiones_admitidas())
                                   + "\n\nPulsa «Formatos admitidos» para ver qué hace cada una.")
            rutas = [r for r in rutas if r not in malos]
        if not rutas:
            return
        if not anadir:
            self.bornas.filas, self.equipos.filas = [], []
            self.bornas.incluir.clear()
            self.equipos.incluir.clear()
            self.copias.clear()
        ya = set(self._esquemas())
        repetidos = [r for r in rutas if self._nombre(r) in ya]
        if repetidos:
            messagebox.showinfo(APP_NAME, "Ya estaban cargados (no se añaden otra vez):\n"
                                + "\n".join(map(os.path.basename, repetidos)))
            rutas = [r for r in rutas if r not in repetidos]
        tablas = [r for r in rutas if L.tipo_archivo(r) == "tabla"]
        esquemas = [r for r in rutas if L.tipo_archivo(r) != "tabla"]
        self.pdf_path = rutas[-1] if rutas else self.pdf_path
        if not esquemas:
            self._abrir_tablas(tablas)
            return
        self._bloquear(True)

        def progreso(msg):
            self.after(0, lambda: self._estado(msg))

        def trabajo():
            resultados = []
            for i, ruta in enumerate(esquemas, start=1):
                prefijo = f"Esquema {i} de {len(esquemas)} ({os.path.basename(ruta)}): " if len(esquemas) > 1 else ""
                try:
                    paginas = L.leer_paginas(ruta, lambda m: progreso(prefijo + m))
                    progreso(prefijo + "buscando bornas y elementos…")
                    bornas = X.extraer_bornas(paginas)
                    elementos = X.extraer_elementos(paginas)
                    # páginas leídas con OCR que han aportado bornas o elementos (para avisar)
                    usadas = {b["pagina"] for b in bornas} | {n for e in elementos for n in e["paginas"]}
                    ocr = [p.numero for p in paginas if p.origen == "ocr" and p.numero in usadas]
                    sin_ocr = [p.numero for p in paginas if p.sin_ocr]
                    nombre = X.numero_esquema(paginas) or self._nombre(ruta)
                    resultados.append((ruta, nombre, bornas, elementos, ocr, sin_ocr, None))
                except L.ErrorLectura as e:
                    resultados.append((ruta, self._nombre(ruta), [], [], [], [], str(e)))
                except Exception as e:  # archivo dañado, etc.
                    resultados.append((ruta, self._nombre(ruta), [], [], [], [],
                                       f"No se ha podido leer el archivo:\n{e}"))
            self.after(0, lambda: self._cargados(resultados, tablas))

        threading.Thread(target=trabajo, daemon=True).start()

    def _bloquear(self, si):
        estado = "disabled" if si else "normal"
        self.btn_abrir.configure(state=estado)
        self.btn_anadir.configure(state=estado)
        self.btn_limpiar.configure(state=estado)

    def _cargados(self, resultados, tablas=()):
        self._bloquear(False)
        errores, avisos_ocr, avisos_sin_ocr, resumen = [], [], [], []
        nuevas_b, nuevos_e = [], []
        ya = set(self._esquemas())
        for ruta, nombre, bornas, elementos, ocr, sin_ocr, error in resultados:
            if sin_ocr:
                pags = ", ".join(map(str, sin_ocr[:15])) + ("…" if len(sin_ocr) > 15 else "")
                avisos_sin_ocr.append(f"{os.path.basename(ruta)}: páginas {pags}")
            if error:
                errores.append(f"{os.path.basename(ruta)}:\n{error}")
                continue
            if nombre in ya:     # mismo número de esquema con otro nombre de archivo
                errores.append(f"{os.path.basename(ruta)}:\nel esquema {nombre} ya estaba cargado.")
                continue
            ya.add(nombre)
            elementos = X.solo_placa(elementos)
            if not bornas and not elementos:
                motivo = ("todas las páginas útiles son imagen (escaneadas) y este programa no tiene OCR."
                          if sin_ocr else "no he encontrado bornas ni elementos.")
                errores.append(f"{os.path.basename(ruta)}:\n{motivo}")
                continue
            for f in bornas + elementos:
                f["esquema"] = nombre
            nuevas_b += bornas
            nuevos_e += elementos
            resumen.append(f"{nombre}: {len(bornas)} bornas, {len(elementos)} elementos")
            if ocr:
                pags = ", ".join(map(str, ocr[:12])) + ("…" if len(ocr) > 12 else "")
                avisos_ocr.append(f"{nombre}: páginas {pags}")
        if nuevas_b or nuevos_e:
            self._cargar_bornas(self.bornas.filas + nuevas_b)
            self._cargar_elementos(self.equipos.filas + nuevos_e)
            self.busqueda.set("")
            self._pintar()
            self.pestanas.select(self.bornas if self.bornas.filas else self.equipos)
        self._actualizar_cabecera()
        if resumen:
            nota = (f"  (sin OCR: no se han leído {sum(len(r[5]) for r in resultados)} páginas sin texto)"
                    if avisos_sin_ocr else "")
            self._estado("Listo. " + " | ".join(resumen) + nota)
        if errores:
            self._estado("Algún archivo no se ha podido leer.")
            messagebox.showerror(APP_NAME, "\n\n".join(errores))
        if avisos_sin_ocr and not (nuevas_b or nuevos_e):
            messagebox.showwarning(APP_NAME, "Estas páginas no tienen texto (son imagen o dibujo) y no se han "
                                             "podido leer porque este programa se creó sin el reconocimiento "
                                             "de texto (OCR):\n" + "\n".join(avisos_sin_ocr) + "\n\n"
                                             "El resto de páginas sí se han leído. Si en esas páginas hay "
                                             "bornas o elementos, vuelve a crear el .exe con crear_exe.bat "
                                             "usando Python 3.12 de 64 bits.")
        if avisos_ocr:
            messagebox.showwarning(APP_NAME, "Algunas páginas eran imagen (escaneadas) y se han leído con "
                                             "reconocimiento de texto (OCR):\n" + "\n".join(avisos_ocr) + "\n\n"
                                             "El OCR puede equivocarse (por ejemplo Q100H en vez de Q100). "
                                             "Revisa la lista antes de imprimir.")
        if tablas:
            self._abrir_tablas(tablas)

    def _abrir_tablas(self, tablas):
        for ruta in tablas:
            self._abrir_tabla(ruta)

    def _abrir_tabla(self, ruta):
        try:
            filas = L.leer_tabla(ruta)
        except L.ErrorLectura as e:
            self._error(f"{os.path.basename(ruta)}:\n{e}")
            return
        dlg = DialogoTabla(self, os.path.basename(ruta), filas, X.analizar_tabla(filas))
        self.wait_window(dlg)
        if not dlg.resultado:
            return
        cab, columnas, modo = dlg.resultado
        datos = X.filas_desde_tabla(filas, cab, columnas, modo)
        if not datos:
            self._error("No he encontrado datos en esas columnas. Revisa la columna elegida.")
            return
        for f in datos:
            f["esquema"] = self._nombre(ruta)
        if modo == "bornas":
            self._cargar_bornas(self.bornas.filas + datos)
            self.pestanas.select(self.bornas)
            self._estado(f"Listo: {len(datos)} bornas del listado {os.path.basename(ruta)}.")
        else:
            datos = X.solo_placa(datos)
            self._cargar_elementos(self.equipos.filas + datos)
            self.pestanas.select(self.equipos)
            self._estado(f"Listo: {len(datos)} elementos del listado {os.path.basename(ruta)} "
                         "(sin motores, pilotos ni pulsadores).")
        self._actualizar_cabecera()
        self._pintar()

    def _actualizar_cabecera(self):
        esquemas = self._esquemas()
        if not esquemas:
            texto = "Abre uno o varios esquemas, o arrástralos encima del programa."
        elif len(esquemas) == 1:
            texto = esquemas[0]
        else:
            texto = f"{len(esquemas)} esquemas: " + ", ".join(esquemas)
            if len(texto) > 110:
                texto = texto[:107] + "…"
        self.lbl_pdf.configure(text=texto)
        self._pintar_copias(esquemas)

    def _pintar_copias(self, esquemas):
        """Una casilla «Copias» por esquema (4 por fila). Se conserva lo que ya se había puesto."""
        for w in self.frame_copias.winfo_children():
            w.destroy()
        if not esquemas:
            return
        ttk.Label(self.frame_copias, text="Copias:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(self.frame_copias, text="(✕ quita el esquema)", foreground="#5d6874").grid(
            row=0, column=5, sticky="w")
        for i, esq in enumerate(esquemas):
            if esq not in self.copias:
                var = tk.StringVar(value="1")
                var.trace_add("write", lambda *_: self._pintar())
                self.copias[esq] = var
            fila, col = divmod(i, 4)
            celda = ttk.Frame(self.frame_copias)
            celda.grid(row=fila, column=1 + col, sticky="w", padx=(0, 18), pady=1)
            ttk.Label(celda, text=esq).pack(side="left")
            ttk.Spinbox(celda, from_=1, to=99, width=4, textvariable=self.copias[esq]).pack(side="left", padx=(6, 0))
            ttk.Button(celda, text="✕", width=3, command=lambda e=esq: self.quitar_esquema(e)).pack(side="left", padx=(4, 0))

    # ---- quitar esquemas ----
    def quitar_esquema(self, esquema):
        if not messagebox.askyesno(APP_NAME, f"¿Quitar el esquema {esquema} de la lista?"):
            return
        self._quitar({esquema})
        self._estado(f"Esquema {esquema} quitado.")

    def limpiar_todo(self):
        if not self._esquemas():
            return
        if not messagebox.askyesno(APP_NAME, "¿Quitar todos los esquemas y empezar de cero?"):
            return
        self._quitar(set(self._esquemas()))
        self._estado("Lista vacía. Abre o añade esquemas.")

    def _quitar(self, esquemas):
        """Quita las bornas y elementos de esos esquemas (y sus casillas de copias)."""
        # primero se filtran las dos listas y luego se recargan: el nombre de los filtros
        # de regleta depende de cuántos esquemas quedan (XALI · CECE… o solo XALI)
        self.bornas.filas = [f for f in self.bornas.filas if f["esquema"] not in esquemas]
        self.equipos.filas = [e for e in self.equipos.filas if e["esquema"] not in esquemas]
        self.bornas.incluir.clear()
        self.bornas.cargar(self.bornas.filas)
        self.equipos.cargar(self.equipos.filas)
        for e in esquemas:
            self.copias.pop(e, None)
        self._actualizar_cabecera()
        self._pintar()

    def _num_copias(self):
        """{esquema: nº de copias} (1 si la casilla está vacía o mal escrita)."""
        res = {}
        for esq, var in self.copias.items():
            try:
                res[esq] = max(1, int(var.get()))
            except (ValueError, tk.TclError):
                res[esq] = 1
        return res

    def _con_copias(self, filas):
        return X.multiplicar(filas, self._num_copias())

    def _error(self, msg):
        self._bloquear(False)
        self._estado("No se ha podido leer el esquema.")
        messagebox.showerror(APP_NAME, msg)

    # ---- SW-CODE (referencia del consumible) ----
    def _cargar_bornas(self, filas):
        """Cada borna toma el SW-CODE que se usó la última vez en su regleta."""
        recordados = self.cfg["swcode_regletas"]
        for f in filas:
            f.setdefault("swcode", recordados.get(f["regleta"], ""))
        self.bornas.cargar(filas)

    def _cargar_elementos(self, filas):
        codigo = self.cfg["swcode_elementos"]
        for e in filas:
            e.setdefault("swcode", codigo)
        self.equipos.cargar(filas)
        self.equipos.swcode.set(codigo)

    def _recordar_swcode_bornas(self, cambiadas):
        """Recuerda el SW-CODE de cada regleta (si todas sus bornas tienen el mismo)."""
        for regleta in {f["regleta"] for f in cambiadas}:
            codigos = {f.get("swcode", "") for f in self.bornas.filas if f["regleta"] == regleta}
            if len(codigos) == 1:
                codigo = codigos.pop()
                if codigo:
                    self.cfg["swcode_regletas"][regleta] = codigo
                else:
                    self.cfg["swcode_regletas"].pop(regleta, None)
        config.guardar(self.cfg)

    def _recordar_swcode_elementos(self, cambiadas):
        """El SW-CODE más usado en los elementos pasa a ser el de por defecto."""
        from collections import Counter
        mas_usado = Counter(e.get("swcode", "") for e in self.equipos.filas).most_common(1)
        if mas_usado and mas_usado[0][0]:
            self.cfg["swcode_elementos"] = mas_usado[0][0]
            config.guardar(self.cfg)

    def _pintar(self):
        q = self.busqueda.get().strip().lower()
        self.bornas.pintar(q)
        self.equipos.pintar(q)
        # pestaña Todo
        self.tabla_todo.delete(*self.tabla_todo.get_children())
        for f in self.bornas.incluidas():
            vals = (f["esquema"], "Borna", f["marca"], f["regleta"], f["descripcion"], f["pagina"], f.get("swcode", ""))
            if not q or any(q in str(v).lower() for v in vals):
                self.tabla_todo.insert("", "end", values=vals)
        for e in self.equipos.incluidas():
            vals = (e["esquema"], "Elemento", e["designacion"], e["tipo"], "", ", ".join(map(str, e["paginas"])), e.get("swcode", ""))
            if not q or any(q in str(v).lower() for v in vals):
                self.tabla_todo.insert("", "end", values=vals)

        nb, ne = len(self.bornas.incluidas()), len(self.equipos.incluidas())
        texto = f"{nb} bornas · {ne} elementos"
        if self._esquemas() and any(n > 1 for n in self._num_copias().values()):
            nb2 = len(self._con_copias(self.bornas.incluidas()))
            ne2 = len(self._con_copias(self.equipos.incluidas()))
            texto += f"  →  con copias: {nb2} + {ne2}"
        self.lbl_tags.configure(text=texto if self._esquemas() else "")
        que = self._que()
        n = {"bornas": nb, "elementos": ne, "todo": nb + ne}[que]
        self.btn_excel.configure(text=f"Guardar Excel de {que}…", state="normal" if n else "disabled")
        self.btn_copiar.configure(state="normal" if n else "disabled")

    def _que(self):
        """Qué se guarda/copia según la pestaña abierta: bornas, elementos o todo."""
        sel = self.pestanas.select()
        if sel == str(self.bornas):
            return "bornas"
        if sel == str(self.equipos):
            return "elementos"
        return "todo"

    def _textos(self, que):
        bornas = [f["marca"] for f in self._con_copias(self.bornas.incluidas())]
        elementos = [e["designacion"] for e in self._con_copias(self.equipos.incluidas())]
        return {"bornas": bornas, "elementos": elementos, "todo": bornas + elementos}[que]

    def guardar(self):
        """Guarda lo de la pestaña abierta: solo bornas, solo elementos o todo."""
        que = self._que()
        # cada esquema repetido tantas veces como diga su casilla «Copias», seguido
        bornas = self._con_copias(self.bornas.incluidas())
        elementos = self._con_copias(self.equipos.incluidas())
        esquemas = self._esquemas()
        base = (esquemas[0] if len(esquemas) == 1 else f"{len(esquemas)}_esquemas") + f"_{que}.xlsx"
        ruta = filedialog.asksaveasfilename(title=f"Guardar Excel de {que}", defaultextension=".xlsx",
                                            initialdir=os.path.dirname(self.pdf_path), initialfile=base,
                                            filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            if que == "bornas":
                guardar_excel_bornas(bornas, ruta)
            elif que == "elementos":
                guardar_excel_elementos(elementos, ruta)
            else:
                guardar_excel_todo(bornas, elementos, ruta)
        except PermissionError:
            messagebox.showerror(APP_NAME, "No se puede guardar: ese Excel está abierto. Ciérralo y vuelve a intentarlo.")
            return
        n = len(self._textos(que))
        self._estado(f"Excel de {que} guardado: {os.path.basename(ruta)} ({n} etiquetas)")
        if sys.platform == "win32" and messagebox.askyesno(APP_NAME, f"Excel de {que} guardado.\n¿Quieres abrirlo ahora?"):
            os.startfile(ruta)

    def copiar(self):
        que = self._que()
        textos = self._textos(que)
        self.clipboard_clear()
        self.clipboard_append("\n".join(textos))
        self._estado(f"Copiadas {len(textos)} etiquetas ({que}), una por línea. Pégalas en Excel o en GeniusPro.")

    def _estado(self, txt):
        self.lbl_estado.configure(text=txt)


def prueba_sin_ventana(ruta):
    """EtiquetadorCuadros.exe --prueba ARCHIVO: lee el archivo y escribe el resumen en
    ARCHIVO.prueba.txt (sirve para comprobar el programa sin abrir la ventana)."""
    lineas = []
    try:
        paginas = L.leer_paginas(ruta, lambda m: None)
        bornas = X.extraer_bornas(paginas)
        elementos = X.solo_placa(X.extraer_elementos(paginas))
        lineas += [f"paginas={len(paginas)}", f"ocr={[p.numero for p in paginas if p.origen == 'ocr']}",
                   f"bornas={len(bornas)}", f"elementos_placa={len(elementos)}",
                   "primeras_bornas=" + ",".join(b["marca"] for b in bornas[:10]),
                   "primeros_elementos=" + ",".join(e["designacion"] for e in elementos[:10])]
    except Exception as e:
        lineas.append(f"ERROR: {e}")
    with open(ruta + ".prueba.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(lineas) + "\n")


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--prueba":
        prueba_sin_ventana(sys.argv[2])
        sys.exit(0)
    archivos = sys.argv[1:] or None  # archivos arrastrados sobre el .exe (uno o varios)
    App(archivos).mainloop()
