"""
Búsqueda de información en una ficha de producción — versión DIGITAL interactiva.

Usa EXACTAMENTE los mismos 36 carteles que genera generador_carteles.py (misma semilla),
así que los resultados digitales y los impresos son comparables.

Flujo por participante (un participante = un cartel):
  1. Formulario (lo llena el grupo): edad, lentes, horas de sueño y distancia a la pantalla.
  2. Pantalla de instrucción: se muestra QUÉ buscar ("N.º DE LOTE") ANTES de empezar,
     para que el tiempo mida solo la búsqueda y no la lectura de la consigna.
  3. Ficha: el cronómetro arranca cuando aparece. Se termina al hacer clic sobre el valor
     de N.º DE LOTE. Los clics en otras celdas cuentan como errores y no detienen la prueba.
  4. Se guarda UNA fila en resultados_digital.csv y se pasa al siguiente participante.

Factores MANIPULADOS: cantidad de elementos (x1) y distancia del objetivo al centro (x2).
Todo lo demás es constante: tipo y tamaño de letra, colores, fondo y formato del objetivo.

Teclas:  Esc = salir de pantalla completa   |   F11 = volver a pantalla completa
         Ctrl+Q = cerrar el programa

Uso:  python busqueda_cartel_digital.py        (requiere:  pip install pillow)
"""
import csv
import math
import os
import random
import time
import tkinter as tk
from datetime import datetime

from PIL import Image, ImageDraw, ImageFont, ImageTk

# ======================================================================
# CONFIGURACIÓN (debe coincidir con generador_carteles.py para obtener los mismos carteles)
# ======================================================================
SEMILLA = 2026
DPI = 100
ANCHO_MM, ALTO_MM = 420, 297

NIVELES_N = {16: (4, 4), 36: (6, 6), 64: (8, 8), 100: (10, 10)}
NIVELES_POSICION = {"centro": 0.0, "media": 0.5, "periferia": 1.0}
REPLICAS = 3
PARTICIPANTES = 100

# MODO = "piloto": premuestra de 15 personas (cubre las 12 combinaciones de cartel)
# MODO = "estudio": muestra completa (PARTICIPANTES)
MODO = "piloto"
PARTICIPANTES_PILOTO = 15
ARCHIVO_RESULTADOS = "resultados_piloto.csv" if MODO == "piloto" else "resultados_digital.csv"

CELDA_W, CELDA_H = 160, 100
FUENTE_ETIQUETA = 16
FUENTE_VALOR = 30
ALTO_ENCABEZADO = 110
MARGEN_CELDA = 8             # px, margen izquierdo del texto dentro de la celda
Y_ETIQUETA = 10               # px, posición vertical de la etiqueta dentro de la celda
Y_VALOR = 42                  # px, posición vertical del valor dentro de la celda
ETIQUETA_OBJETIVO = "N.º DE LOTE"

COLOR_FONDO = (255, 255, 255)
COLOR_BORDE = (190, 190, 190)
COLOR_ETIQUETA = (70, 70, 70)
COLOR_VALOR = (0, 0, 0)
COLOR_ENCABEZADO = (20, 50, 100)

# Opciones de la versión digital
TIEMPO_LIMITE_S = 90        # segundos; None = sin límite. Si se agota, se guarda encontrado = 0
PANTALLA_COMPLETA = True
FONDO_VENTANA = "#1e1e1e"


# ======================================================================
# CAMPOS DE LA FICHA (idénticos a generador_carteles.py; NO cambiar el orden)
# ======================================================================
def _n(a, b):
    return lambda r: str(r.randint(a, b))


def _u(a, b, u):
    return lambda r: f"{r.randint(a, b)} {u}"


def _p(prefijo, a, b):
    return lambda r: f"{prefijo}{r.randint(a, b)}"


CAMPOS = [
    ("ORDEN PROD.", _p("OP-", 1000, 9999)),
    ("CÓD. PRODUCTO", _n(100000, 999999)),
    ("PESO NETO", _u(100, 2500, "g")),
    ("VENCIMIENTO", lambda r: f"{r.randint(1, 28):02d}/{r.randint(1, 12):02d}/27"),
    ("LÍNEA", lambda r: f"Línea {r.randint(1, 9)}"),
    ("TURNO", lambda r: r.choice(["A", "B", "C"])),
    ("CAJAS/PALET", _n(10, 99)),
    ("TEMP. MÁX.", _u(4, 35, "°C")),
    ("HUMEDAD", _u(30, 80, "%")),
    ("PRESIÓN", lambda r: f"{r.randint(1, 9)}.{r.randint(0, 9)} bar"),
    ("VELOCIDAD", _u(20, 120, "u/min")),
    ("VOLTAJE", lambda r: f"{r.choice([110, 220, 440])} V"),
    ("N.º SERIE", _p("SN", 10000, 99999)),
    ("OPERARIO", _p("ID ", 100, 999)),
    ("INSPECTOR", _p("ID ", 100, 999)),
    ("CANT. PEDIDA", _u(100, 9999, "u")),
    ("LARGO", _u(50, 990, "mm")),
    ("ANCHO", _u(50, 990, "mm")),
    ("ESPESOR", _u(1, 30, "mm")),
    ("CICLO", _u(5, 90, "s")),
    ("CURADO", _u(5, 90, "min")),
    ("TOLERANCIA", lambda r: f"±{r.randint(1, 9)}.{r.randint(0, 9)} mm"),
    ("BODEGA", _p("B-", 10, 99)),
    ("ESTANTE", _p("E-", 100, 999)),
    ("PEDIDO", _p("PD-", 1000, 9999)),
    ("RUTA", _p("R-", 10, 99)),
    ("CLIENTE", _p("C-", 1000, 9999)),
    ("MÁQUINA", _p("M-", 10, 99)),
    ("MOLDE", _p("MD-", 100, 999)),
    ("MERMA", lambda r: f"{r.randint(0, 9)}.{r.randint(0, 9)} %"),
    ("TEMP. PROCESO", _u(80, 250, "°C")),
    ("AMPERAJE", _u(5, 60, "A")),
    ("POTENCIA", _u(1, 90, "kW")),
    ("GIRO", _u(500, 3500, "rpm")),
    ("CAUDAL", _u(5, 99, "L/min")),
    ("ALTURA", _u(100, 990, "mm")),
    ("ZONA", _p("Z-", 10, 99)),
    ("REVISIÓN", _p("REV ", 1, 15)),
    ("SECCIÓN", _p("S-", 1, 20)),
    # --- campos adicionales para las fichas de 64 y 100 elementos ---
    ("DIÁMETRO", _u(5, 500, "mm")),
    ("PROFUNDIDAD", _u(1, 200, "mm")),
    ("DUREZA", _p("HRC ", 20, 65)),
    ("TORQUE", _u(5, 450, "N·m")),
    ("FUERZA", _u(1, 90, "kN")),
    ("VISCOSIDAD", _u(10, 990, "cP")),
    ("PH", lambda r: f"{r.randint(4, 9)}.{r.randint(0, 9)}"),
    ("CONDUCT.", _u(50, 990, "µS")),
    ("NIVEL TANQUE", _u(5, 99, "%")),
    ("TANQUE", _p("T-", 10, 99)),
    ("SILO", _p("SL-", 10, 99)),
    ("BANDA", _p("BT-", 10, 99)),
    ("BOBINA", _p("BB-", 100, 999)),
    ("TARIMA", _p("TR-", 1000, 9999)),
    ("CONTENEDOR", _p("CT-", 1000, 9999)),
    ("SELLO", _p("SE-", 1000, 9999)),
    ("PRECINTO", _p("PR-", 10000, 99999)),
    ("FACTURA", _p("FA-", 10000, 99999)),
    ("GUÍA", _p("GU-", 10000, 99999)),
    ("DESPACHO", _p("DS-", 1000, 9999)),
    ("MUELLE", _p("ML-", 1, 20)),
    ("CHOFER", _p("ID ", 100, 999)),
    ("TRANSPORTE", _p("TP-", 100, 999)),
    ("PROVEEDOR", _p("PV-", 100, 999)),
    ("FAMILIA", _p("FM-", 10, 99)),
    ("MODELO", _p("MO-", 100, 999)),
    ("COLOR", _p("RAL ", 1000, 9010)),
    ("TALLA", lambda r: r.choice(["S", "M", "L", "XL"])),
    ("GRADO", lambda r: r.choice(["A", "B", "C"])),
    ("CALIBRE", _u(1, 40, "mm")),
    ("SEMANA", _p("SEM ", 1, 52)),
    ("REF. INTERNA", _p("RI-", 1000, 9999)),
    ("CERTIFICADO", _p("CE-", 1000, 9999)),
    ("AUDITORÍA", _p("AU-", 100, 999)),
    ("HORNO", _p("H-", 1, 20)),
    ("EXTRUSORA", _p("EX-", 1, 20)),
    ("PRENSA", _p("PN-", 1, 20)),
    ("TORNO", _p("TN-", 1, 20)),
    ("ROBOT", _p("RB-", 1, 20)),
    ("CELDA", _p("CL-", 1, 20)),
    ("ESTACIÓN", _p("ST-", 1, 30)),
    ("PUESTO", _p("PU-", 1, 30)),
    ("EMPAQUE", _p("EM-", 100, 999)),
    ("ETIQUETA", _p("ET-", 1000, 9999)),
    ("CÓD. BARRAS", _n(1000000, 9999999)),
    ("COSTO", lambda r: f"${r.randint(1, 99)}.{r.randint(0, 9)}0"),
    ("PRECIO", lambda r: f"${r.randint(1, 99)}.{r.randint(0, 9)}5"),
    ("IMPUESTO", _u(1, 15, "%")),
    ("DESCUENTO", _u(1, 40, "%")),
    ("UNID./CAJA", _n(2, 48)),
    ("PESO BRUTO", _u(100, 2900, "g")),
    ("TARA", _u(10, 400, "g")),
    ("VOLUMEN", _u(1, 990, "L")),
    ("SUPERFICIE", _u(1, 99, "m²")),
    ("DENSIDAD", lambda r: f"{r.randint(8, 19) / 10:.1f} g/cm³"),
    ("FRECUENCIA", _u(30, 90, "Hz")),
    ("CARGA", _u(10, 990, "kg")),
    ("CAPACIDAD", _u(100, 5000, "kg")),
    ("RECHAZOS", _u(0, 99, "u")),
    ("DEFECTOS", lambda r: f"{r.randint(0, 9)}.{r.randint(0, 9)} %"),
    ("REPROCESO", lambda r: f"{r.randint(0, 9)}.{r.randint(0, 9)} %"),
    ("EFICIENCIA", _u(60, 99, "%")),
    ("PARO", _u(1, 90, "min")),
    ("ARRANQUE", lambda r: f"{r.randint(5, 9):02d}:{r.choice([0, 15, 30, 45]):02d}"),
    ("CIERRE", lambda r: f"{r.randint(14, 22):02d}:{r.choice([0, 15, 30, 45]):02d}"),
    ("ENTREGA", lambda r: f"{r.randint(1, 28):02d}/{r.randint(1, 12):02d}"),
    ("FABRICACIÓN", lambda r: f"{r.randint(1, 28):02d}/{r.randint(1, 12):02d}/26"),
    ("ENSAYO", _p("EN-", 1000, 9999)),
]


# ======================================================================
# FUENTES Y GEOMETRÍA
# ======================================================================
def cargar_fuente(tam):
    candidatas = [
        "arialbd.ttf",
        "Arial Bold.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "C:/Windows/Fonts/arialbd.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]
    for ruta in candidatas:
        try:
            return ImageFont.truetype(ruta, tam)
        except OSError:
            continue
    raise RuntimeError("No se encontró una fuente en negrita (Arial, Liberation Sans o DejaVu).")


F_ETIQ = cargar_fuente(FUENTE_ETIQUETA)
F_VALOR = cargar_fuente(FUENTE_VALOR)
F_TITULO = cargar_fuente(44)
F_ID = cargar_fuente(28)

PX_A_MM = 25.4 / DPI
W = round(ANCHO_MM / 25.4 * DPI)
H = round(ALTO_MM / 25.4 * DPI)

assert len(CAMPOS) >= max(NIVELES_N) - 1, "Faltan campos para el nivel más grande"
assert len({e for e, _ in CAMPOS}) == len(CAMPOS), "Hay etiquetas repetidas en CAMPOS"
_c, _f = NIVELES_N[max(NIVELES_N)]
assert _c * CELDA_W <= W and ALTO_ENCABEZADO + _f * CELDA_H <= H, "El nivel más grande no cabe en la hoja"


def geometria(cols, filas):
    bloque_w, bloque_h = cols * CELDA_W, filas * CELDA_H
    x0 = (W - bloque_w) // 2
    y0 = ALTO_ENCABEZADO + (H - ALTO_ENCABEZADO - bloque_h) // 2
    return x0, y0, x0 + bloque_w / 2, y0 + bloque_h / 2


def elegir_celda_objetivo(cols, filas, fraccion, rng):
    x0, y0, cx, cy = geometria(cols, filas)
    celdas = []
    for f in range(filas):
        for c in range(cols):
            ccx = x0 + c * CELDA_W + CELDA_W / 2
            ccy = y0 + f * CELDA_H + CELDA_H / 2
            celdas.append((f, c, math.hypot(ccx - cx, ccy - cy)))
    d_max = max(d for _, _, d in celdas)
    objetivo = fraccion * d_max
    mejor = min(abs(d - objetivo) for _, _, d in celdas)
    candidatas = [(f, c, d) for f, c, d in celdas if abs(abs(d - objetivo) - mejor) < 1e-6]
    return rng.choice(candidatas)


# ======================================================================
# DISEÑO Y DIBUJO DE CARTELES (el diseño usa rng; el dibujo no)
# ======================================================================
def disenar_cartel(cartel_id, n, nivel, fraccion, rng):
    cols, filas = NIVELES_N[n]
    f_obj, c_obj, d_obj = elegir_celda_objetivo(cols, filas, fraccion, rng)
    valor_obj = f"L-{rng.randint(10000, 99999)}"
    otros = rng.sample(CAMPOS, n - 1)
    contenido = [(et, gen(rng)) for et, gen in otros]
    rng.shuffle(contenido)
    contenido.insert(f_obj * cols + c_obj, (ETIQUETA_OBJETIVO, valor_obj))
    return {
        "cartel_id": cartel_id,
        "n_elementos": n,
        "columnas": cols,
        "filas": filas,
        "nivel_posicion": nivel,
        "fila_obj": f_obj,
        "col_obj": c_obj,
        "dist_centro_cm": round(d_obj * PX_A_MM / 10, 1),
        "valor_objetivo": valor_obj,
        "contenido": contenido,
    }


def dibujar_cartel(d):
    cols, filas = d["columnas"], d["filas"]
    x0, y0, _, _ = geometria(cols, filas)
    img = Image.new("RGB", (W, H), COLOR_FONDO)
    draw = ImageDraw.Draw(img)

    draw.rectangle([(0, 0), (W, ALTO_ENCABEZADO)], fill=COLOR_ENCABEZADO)
    draw.text((40, 28), "FICHA DE CONTROL DE PRODUCCIÓN", fill="white", font=F_TITULO)
    texto_id = f"ID {d['cartel_id']}"
    ancho_id = draw.textlength(texto_id, font=F_ID)
    draw.text((W - ancho_id - 40, 36), texto_id, fill=(200, 210, 230), font=F_ID)

    margen = MARGEN_CELDA
    for i, (etiqueta, valor) in enumerate(d["contenido"]):
        f, c = divmod(i, cols)
        x = x0 + c * CELDA_W
        y = y0 + f * CELDA_H
        draw.rectangle([(x, y), (x + CELDA_W, y + CELDA_H)], outline=COLOR_BORDE, width=2)
        draw.text((x + margen, y + Y_ETIQUETA), etiqueta, fill=COLOR_ETIQUETA, font=F_ETIQ)
        draw.text((x + margen, y + Y_VALOR), valor, fill=COLOR_VALOR, font=F_VALOR)
    return img


def secuencia_piloto(disenos, cantidad=None):
    """Premuestra: cubre las 12 combinaciones (cantidad x posición) con carteles distintos
    y reparte el resto entre las combinaciones menos usadas. Orden aleatorio."""
    cantidad = min(cantidad or PARTICIPANTES_PILOTO, len(disenos))
    rng = random.Random(SEMILLA + 7)
    combos = {}
    for d in disenos:
        combos.setdefault((d["n_elementos"], d["nivel_posicion"]), []).append(d["cartel_id"])
    elegidos = [rng.choice(ids) for ids in combos.values()]
    conteo = {k: 1 for k in combos}
    while len(elegidos) < cantidad:
        minimo = min(c for k, c in conteo.items() if any(i not in elegidos for i in combos[k]))
        candidatas = [k for k, c in conteo.items()
                      if c == minimo and any(i not in elegidos for i in combos[k])]
        k = rng.choice(candidatas)
        elegidos.append(rng.choice([i for i in combos[k] if i not in elegidos]))
        conteo[k] += 1
    rng.shuffle(elegidos)
    return elegidos[:cantidad]


def construir_diseno():
    """Reproduce el mismo orden de generación que generador_carteles.py."""
    rng = random.Random(SEMILLA)
    disenos, contador = [], 0
    for n in NIVELES_N:
        for nivel, fraccion in NIVELES_POSICION.items():
            for _ in range(REPLICAS):
                contador += 1
                disenos.append(disenar_cartel(f"C{contador:02d}", n, nivel, fraccion, rng))

    ids = [d["cartel_id"] for d in disenos]
    secuencia = []
    while len(secuencia) < PARTICIPANTES:
        bloque = ids[:]
        rng.shuffle(bloque)
        secuencia.extend(bloque)
    if MODO == "piloto":
        return disenos, secuencia_piloto(disenos)
    return disenos, secuencia[:PARTICIPANTES]


# ======================================================================
# RESULTADOS
# ======================================================================
COLUMNAS = [
    "id_participante", "cartel_id", "n_elementos", "nivel_posicion", "dist_centro_cm",
    "fila_obj", "col_obj", "valor_objetivo",
    "tiempo_s", "errores", "encontrado",
    "edad", "usa_lentes_0_1", "horas_sueno", "distancia_persona_m",
    "hora", "fondo", "color_texto", "tam_letra_valor_px", "escala_pantalla", "pantalla_px",
]


def siguiente_id():
    if not os.path.exists(ARCHIVO_RESULTADOS):
        return 1
    with open(ARCHIVO_RESULTADOS, newline="", encoding="utf-8-sig") as f:
        return max(sum(1 for _ in csv.reader(f)) - 1, 0) + 1


def guardar_fila(fila):
    nuevo = not os.path.exists(ARCHIVO_RESULTADOS)
    with open(ARCHIVO_RESULTADOS, "a", newline="", encoding="utf-8-sig" if nuevo else "utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNAS)
        if nuevo:
            w.writeheader()
        w.writerow(fila)


# ======================================================================
# INTERFAZ
# ======================================================================
class App:
    def __init__(self, disenos, secuencia):
        self.por_id = {d["cartel_id"]: d for d in disenos}
        self.ids = [d["cartel_id"] for d in disenos]
        self.secuencia = list(secuencia)
        self.extra_rng = random.Random(SEMILLA + 1)

        self.root = tk.Tk()
        self.root.title("Búsqueda de información en ficha de producción")
        self.root.configure(bg=FONDO_VENTANA)
        if PANTALLA_COMPLETA:
            self.root.attributes("-fullscreen", True)
        else:
            self.root.geometry("1280x800")
        self.root.bind("<Escape>", lambda e: self.root.attributes("-fullscreen", False))
        self.root.bind("<F11>", lambda e: self.root.attributes("-fullscreen", True))
        self.root.bind("<Control-q>", lambda e: self.root.destroy())

        self.ultimo = {"distancia": ""}
        self.datos = {}
        self.foto = None
        self.timer_id = None
        self.t0 = None
        self.activo = False
        self.errores = 0
        self.escala = 1.0
        self.panel = None
        self.mostrar_formulario()

    # ---------- utilidades ----------
    def limpiar(self):
        if self.timer_id is not None:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None
        if self.panel is not None:
            self.panel.destroy()
        self.panel = tk.Frame(self.root, bg=FONDO_VENTANA)
        self.panel.pack(fill="both", expand=True)
        self.root.unbind("<space>")
        self.root.unbind("<Return>")
        return self.panel

    def cartel_para(self, id_participante):
        while len(self.secuencia) < id_participante:
            bloque = self.ids[:]
            self.extra_rng.shuffle(bloque)
            self.secuencia.extend(bloque)
        return self.por_id[self.secuencia[id_participante - 1]]

    def etiqueta(self, parent, texto, tam=14, color="white", negrita=False, **kw):
        return tk.Label(parent, text=texto, bg=FONDO_VENTANA, fg=color,
                        font=("Arial", tam, "bold" if negrita else "normal"), **kw)

    # ---------- 1. formulario ----------
    def mostrar_formulario(self):
        p = self.limpiar()
        self.id_part = siguiente_id()
        if MODO == "piloto" and self.id_part > PARTICIPANTES_PILOTO:
            return self.mostrar_completo()
        self.diseno = self.cartel_para(self.id_part)

        centro = tk.Frame(p, bg=FONDO_VENTANA)
        centro.place(relx=0.5, rely=0.5, anchor="center")

        total = f" de {PARTICIPANTES_PILOTO}" if MODO == "piloto" else ""
        self.etiqueta(centro, f"Participante {self.id_part}{total}  |  Cartel {self.diseno['cartel_id']}",
                      22, "#7fb2ff", True).grid(row=0, column=0, columnspan=2, pady=(0, 18))
        self.etiqueta(centro, "(Lo llena el grupo antes de entregar la pantalla)", 11, "#aaaaaa"
                      ).grid(row=1, column=0, columnspan=2, pady=(0, 14))

        campos = [
            ("edad", "Edad (años)", ""),
            ("sueno", "Horas de sueño de anoche", ""),
            ("distancia", "Distancia de la persona a la pantalla (m)", self.ultimo["distancia"]),
        ]
        self.entradas = {}
        for i, (clave, texto, valor) in enumerate(campos, start=2):
            self.etiqueta(centro, texto, 14, anchor="e").grid(row=i, column=0, sticky="e", padx=10, pady=6)
            e = tk.Entry(centro, font=("Arial", 14), width=12)
            e.insert(0, valor)
            e.grid(row=i, column=1, sticky="w", pady=6)
            self.entradas[clave] = e

        self.var_lentes = tk.IntVar(value=0)
        tk.Checkbutton(centro, text="Usa lentes (gafas o contacto)", variable=self.var_lentes,
                       font=("Arial", 14), bg=FONDO_VENTANA, fg="white", selectcolor="#333333",
                       activebackground=FONDO_VENTANA, activeforeground="white"
                       ).grid(row=len(campos) + 2, column=0, columnspan=2, pady=10)

        self.lbl_error = self.etiqueta(centro, "", 12, "#ff7b7b")
        self.lbl_error.grid(row=len(campos) + 3, column=0, columnspan=2)
        tk.Button(centro, text="Continuar", font=("Arial", 14, "bold"), width=16,
                  command=self.continuar).grid(row=len(campos) + 4, column=0, columnspan=2, pady=14)
        self.entradas["edad"].focus_set()

    def _numero(self, clave, minimo, maximo, entero=False, opcional=False):
        txt = self.entradas[clave].get().strip().replace(",", ".")
        if txt == "":
            if opcional:
                return ""
            raise ValueError(f"Falta: {clave}")
        valor = float(txt)
        if not (minimo <= valor <= maximo):
            raise ValueError(f"Valor fuera de rango en: {clave}")
        return int(valor) if entero else valor

    def continuar(self):
        try:
            self.datos = {
                "edad": self._numero("edad", 5, 110, entero=True),
                "horas_sueno": self._numero("sueno", 0, 24),
                "distancia_persona_m": self._numero("distancia", 0, 20),
                "usa_lentes_0_1": self.var_lentes.get(),
            }
        except ValueError as err:
            self.lbl_error.config(text=f"Revise los datos ({err})")
            return
        self.ultimo = {"distancia": self.entradas["distancia"].get()}
        self.mostrar_instruccion()

    # ---------- 2. instrucción ----------
    def mostrar_instruccion(self):
        p = self.limpiar()
        centro = tk.Frame(p, bg=FONDO_VENTANA)
        centro.place(relx=0.5, rely=0.5, anchor="center")
        self.etiqueta(centro, "En la ficha que va a ver, encuentre el valor de", 24).pack(pady=10)
        self.etiqueta(centro, ETIQUETA_OBJETIVO, 64, "#ffd34d", True).pack(pady=14)
        self.etiqueta(centro, "y haga clic sobre ese valor lo más rápido que pueda.", 24).pack(pady=10)
        self.etiqueta(centro, "El tiempo empieza cuando aparece la ficha.", 16, "#aaaaaa").pack(pady=(30, 6))
        tk.Button(centro, text="Estoy listo (ESPACIO)", font=("Arial", 18, "bold"),
                  command=self.mostrar_cartel).pack(pady=16)
        self.root.bind("<space>", lambda e: self.mostrar_cartel())

    # ---------- 3. cartel ----------
    def mostrar_cartel(self):
        p = self.limpiar()
        self.root.update_idletasks()
        cw = max(self.root.winfo_width(), 400)
        ch = max(self.root.winfo_height(), 300)
        self.escala = min(cw / W, ch / H)
        self.tam_pantalla = (cw, ch)
        ancho, alto = int(W * self.escala), int(H * self.escala)
        self.off_x, self.off_y = (cw - ancho) // 2, (ch - alto) // 2

        imagen = dibujar_cartel(self.diseno).resize((ancho, alto), Image.LANCZOS)
        self.foto = ImageTk.PhotoImage(imagen)
        self.canvas = tk.Canvas(p, bg=FONDO_VENTANA, highlightthickness=0, cursor="hand2")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_image(self.off_x, self.off_y, anchor="nw", image=self.foto)
        self.canvas.bind("<Button-1>", self.clic)

        self.errores = 0
        self.activo = True
        self.root.update()
        self.t0 = time.perf_counter()  # el cronómetro arranca cuando la ficha ya está en pantalla
        if TIEMPO_LIMITE_S:
            self.timer_id = self.root.after(int(TIEMPO_LIMITE_S * 1000), self.tiempo_agotado)

    def clic(self, event):
        if not self.activo:
            return
        t = time.perf_counter() - self.t0
        d = self.diseno
        x0, y0, _, _ = geometria(d["columnas"], d["filas"])
        px = (event.x - self.off_x) / self.escala
        py = (event.y - self.off_y) / self.escala
        col = math.floor((px - x0) / CELDA_W)
        fila = math.floor((py - y0) / CELDA_H)
        if not (0 <= col < d["columnas"] and 0 <= fila < d["filas"]):
            return  # clic fuera de las celdas: se ignora
        if (fila, col) == (d["fila_obj"], d["col_obj"]):
            self.terminar(t, encontrado=1)
        else:
            self.errores += 1
            self.marcar_error(fila, col, x0, y0)

    def marcar_error(self, fila, col, x0, y0):
        s = self.escala
        x1 = self.off_x + (x0 + col * CELDA_W) * s
        y1 = self.off_y + (y0 + fila * CELDA_H) * s
        marca = self.canvas.create_rectangle(x1, y1, x1 + CELDA_W * s, y1 + CELDA_H * s,
                                             outline="#e53935", width=4)
        canvas = self.canvas

        def borrar():
            try:
                canvas.delete(marca)
            except tk.TclError:
                pass  # el cartel ya se cerró (acertó justo después del error)

        self.root.after(350, borrar)

    def tiempo_agotado(self):
        if self.activo:
            self.terminar(float(TIEMPO_LIMITE_S), encontrado=0)

    # ---------- 4. fin y guardado ----------
    def terminar(self, tiempo, encontrado):
        self.activo = False
        d = self.diseno
        fila = {
            "id_participante": self.id_part,
            "cartel_id": d["cartel_id"],
            "n_elementos": d["n_elementos"],
            "nivel_posicion": d["nivel_posicion"],
            "dist_centro_cm": d["dist_centro_cm"],
            "fila_obj": d["fila_obj"] + 1,
            "col_obj": d["col_obj"] + 1,
            "valor_objetivo": d["valor_objetivo"],
            "tiempo_s": round(tiempo, 3),
            "errores": self.errores,
            "encontrado": encontrado,
            "hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "fondo": "blanco",
            "color_texto": "negro",
            "tam_letra_valor_px": FUENTE_VALOR,
            "escala_pantalla": round(self.escala, 3),
            "pantalla_px": f"{self.tam_pantalla[0]}x{self.tam_pantalla[1]}",
        }
        fila.update(self.datos)
        guardar_fila(fila)
        self.mostrar_fin()

    def mostrar_fin(self):
        p = self.limpiar()
        centro = tk.Frame(p, bg=FONDO_VENTANA)
        centro.place(relx=0.5, rely=0.5, anchor="center")
        self.etiqueta(centro, "¡Gracias por participar!", 40, "#7fe08a", True).pack(pady=20)
        self.etiqueta(centro, "(Entregue la pantalla al grupo)", 14, "#aaaaaa").pack(pady=10)
        ultimo = MODO == "piloto" and self.id_part >= PARTICIPANTES_PILOTO
        tk.Button(centro, text="Finalizar premuestra" if ultimo else "Siguiente participante",
                  font=("Arial", 16, "bold"), command=self.mostrar_formulario).pack(pady=20)

    def mostrar_completo(self):
        p = self.limpiar()
        centro = tk.Frame(p, bg=FONDO_VENTANA)
        centro.place(relx=0.5, rely=0.5, anchor="center")
        self.etiqueta(centro, f"Premuestra completa ({PARTICIPANTES_PILOTO} de {PARTICIPANTES_PILOTO})",
                      32, "#7fe08a", True).pack(pady=20)
        self.etiqueta(centro, f"Los datos están en {ARCHIVO_RESULTADOS}", 16).pack(pady=10)
        self.etiqueta(centro, "Para la muestra completa cambie MODO a \"estudio\" en el código.",
                      13, "#aaaaaa").pack(pady=6)
        tk.Button(centro, text="Cerrar", font=("Arial", 16, "bold"),
                  command=self.root.destroy).pack(pady=20)

    def ejecutar(self):
        self.root.mainloop()


def validar_textos(disenos):
    """Comprueba al iniciar que ningún texto se sale de su celda (evita sorpresas durante la prueba)."""
    medidor = ImageDraw.Draw(Image.new("RGB", (10, 10)))
    maximo = CELDA_W - 2 * MARGEN_CELDA
    for d in disenos:
        for etiqueta, valor in d["contenido"]:
            if medidor.textlength(etiqueta, font=F_ETIQ) > maximo or medidor.textlength(valor, font=F_VALOR) > maximo:
                raise ValueError(f"En {d['cartel_id']}, '{etiqueta}: {valor}' no cabe en la celda")


def main():
    disenos, secuencia = construir_diseno()
    validar_textos(disenos)
    App(disenos, secuencia).ejecutar()


if __name__ == "__main__":
    main()