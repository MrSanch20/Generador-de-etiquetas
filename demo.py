"""
Datos de ejemplo INVENTADOS para enseñar el programa sin usar esquemas reales.

    python etiquetador.py --demo        (o EtiquetadorCuadros.exe --demo)
"""
import extraccion as X


def _fuerza(n0, motores):
    out = []
    for k, nombre in enumerate(motores):
        n = n0 + k
        out += [(f"U{n}", nombre), (f"V{n}", nombre), (f"W{n}", nombre), (f"PE{n}", nombre)]
    return out


def _bornas(esquema, regletas, swcode_senal=("XS", "XE", "X24")):
    filas = []
    for pag, (reg, marcas) in enumerate(regletas, start=2):
        for i, (marca, desc) in enumerate(marcas, start=1):
            filas.append({"esquema": esquema, "regleta": reg, "pos": i, "marca": marca, "descripcion": desc,
                          "pagina": pag, "swcode": "212" if reg in swcode_senal else "205"})
    return filas


def _elementos(esquema, lista):
    return [{"esquema": esquema, "designacion": d, "tipo": X.tipo_equipo(d), "ubicacion": "Placa",
             "paginas": [p], "swcode": "181"} for d, p in lista]


def datos():
    acometida = [("PE", ""), ("N", "ACOMETIDA"), ("L1", "ACOMETIDA"), ("L2", "ACOMETIDA"), ("L3", "ACOMETIDA")]
    bornas = _bornas("CECE00000000001", [
        ("XACO", acometida),
        ("XM", _fuerza(100, ["BOMBA 1", "BOMBA 2", "VENTILADOR 1", "VENTILADOR 2", "CINTA 1", "CINTA 2"])),
        ("XS", [(f"{a}S{i}", f"SENSOR {i}") for i in range(1, 5) for a in (1, 2)] + [("PANT", "PANTALLA")]),
        ("X24", [("24V", "MANDO"), ("0V", "MANDO")]),
    ]) + _bornas("CECE00000000002", [
        ("XACO", acometida),
        ("XM", _fuerza(200, ["COMPRESOR", "EXTRACTOR 1", "EXTRACTOR 2"])),
        ("XE", [(str(i), "ENTRADA DIGITAL") for i in range(1, 9)]),
    ])
    elementos = _elementos("CECE00000000001", [("Q0", 3), ("F1", 3), ("F2", 3), ("U1", 4)]
                           + [(f"Q{100 + i}", 5) for i in range(6)] + [(f"KM{100 + i}", 5) for i in range(6)]
                           + [("K1", 6), ("K2", 6), ("K3", 6), ("KA1", 6)]) \
        + _elementos("CECE00000000002", [("Q0", 3), ("F1", 3), ("U1", 4)]
                     + [(f"Q{200 + i}", 5) for i in range(3)] + [(f"KM{200 + i}", 5) for i in range(3)]
                     + [("K10", 6), ("K11", 6)])
    return bornas, elementos
