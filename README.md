# Etiquetador de cuadros

Aplicación de escritorio para Windows que lee esquemas eléctricos y genera automáticamente las etiquetas de un cuadro (bornas y elementos de la placa) en un Excel. Sirve para cualquier software de impresión de etiquetas que importe datos desde Excel. Está probado con **Cembre GeniusPro**.

Evita copiar a mano cientos de marcas desde los planos: abres el esquema, revisas la lista y exportas.

## Qué hace

- **Bornas:** lee las páginas «Plano de bornes» (también en francés, inglés y alemán) y saca la marca de cada borna, su regleta, posición y una descripción orientativa.
- **Elementos de la placa:** detecta las designaciones de los aparatos cableados en la placa (Q, KM, K, F, ATS, U…). Descarta los de la puerta y los de fuera del armario (motores, pilotos, pulsadores, finales de carrera).
- **Varios esquemas a la vez**, identificados por el número de instalación del cajetín. Se pueden añadir, quitar o limpiar sin cerrar el programa.
- **Copias por esquema:** repite las etiquetas de un esquema las veces que haga falta, en bloques seguidos.
- **SW-CODE:** asigna a cada etiqueta la referencia del consumible donde se imprime. Recuerda el valor de cada regleta para los siguientes esquemas.
- **Exportación a Excel** por pestaña (solo bornas, solo elementos o todo), sin cabeceras y con la etiqueta siempre en la columna C.

## Formatos admitidos

| Formato | Cómo se lee |
|---|---|
| PDF exportado del programa de esquemas | Texto del PDF (el más fiable) |
| PDF escaneado, PNG, JPG, TIF, BMP | Reconocimiento de texto (OCR) |
| DXF (AutoCAD) | Textos, MTEXT y atributos de bloques |
| DWG (AutoCAD) | Conversión previa con ODA File Converter |
| Excel / CSV | Listados de bornas o aparatos exportados (EPLAN, SEE Electrical, AutoCAD Electrical…) |

## Formato del Excel

Todas las hojas empiezan igual, sin fila de cabecera:

| A | B | C | D |
|---|---|---|---|
| Esquema | Copia | **Etiqueta** | SW-CODE |

Se puede importar en cualquier programa de etiquetas que lea Excel. Ejemplo probado con Cembre GeniusPro → *Importación → De Excel*: **Línea de inicio = 1**, **Valor inicial = [C]**, **Esquema = [A]**.

## Crear el ejecutable

Requisitos solo en el PC donde se compila:

- Windows con **Python 3.12 de 64 bits** (el OCR no funciona todavía con versiones más nuevas).
- Conexión a internet.

Después, doble clic en `crear_exe.bat`. El programa queda en `dist/EtiquetadorCuadros.exe`. Ese `.exe` funciona en cualquier PC con Windows sin instalar Python ni tener internet.

Para probarlo sin ventana:

```
EtiquetadorCuadros.exe --prueba ESQUEMA.pdf
```

Deja un resumen en `ESQUEMA.pdf.prueba.txt`.

## Estructura del código

| Archivo | Función |
|---|---|
| `etiquetador.py` | Ventana (Tkinter): pestañas, filtros, copias, SW-CODE |
| `lectores.py` | Un lector por formato → páginas con textos y posiciones |
| `extraccion.py` | Detección de bornas, elementos, número de esquema y copias |
| `excel.py` | Generación de los Excel para GeniusPro |
| `config.py` | Ajustes recordados (SW-CODE por regleta) |

## Tecnologías

Python · Tkinter · PyMuPDF · RapidOCR (ONNX Runtime) · ezdxf · openpyxl · PyInstaller

## Limitaciones

- La descripción de cada borna es orientativa (el texto más cercano bajo la borna).
- El OCR puede equivocarse: revisa la lista antes de imprimir.
- La ubicación placa/puerta se basa en las páginas del armario y en el tipo de aparato; conviene revisar la lista de elementos.

## Autor

Daniel Sancho Goñi
