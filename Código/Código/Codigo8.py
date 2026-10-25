"""
Búsqueda de información en una ficha de producción — Versión Industrial Definitiva
Incluye:
- Factores manipulados: Densidad de elementos (N), Posición, Tamaño de fuente variable y Nivel de Contraste (alto, medio, bajo).
- Sin formularios de usuario intermedios (ejecución continua).
- Registro automático en CSV para análisis de Regresión Lineal Múltiple (RLM).
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
# CONFIGURACIÓN GENERAL
# ======================================================================
SEMILLA = 2026
DPI = 100
ANCHO_MM, ALTO_MM = 420, 297

NIVELES_N = {16: (4, 4), 36: (6, 6), 64: (8, 8), 100: (10, 10)}
NIVELES_POSICION = {"centro": 0.0, "media": 0.5, "periferia": 1.0}
REPLICAS = 3

# PREDICTOR CRÍTICO: Niveles de Contraste (Texto vs Fondo)
# Definimos colores de fondo base y sus combinaciones de contraste
CONFIG_CONTRASTES = {
    "alto": {"fondo": (255, 255, 255), "texto": (0, 0, 0)},             # Negro sobre blanco (máximo contraste)
    "medio": {"fondo": (230, 235, 245), "texto": (70, 90, 120)},       # Gris azulado sobre fondo claro
    "bajo": {"fondo": (210, 215, 220), "texto": (160, 165, 175)}       # Gris suave sobre gris (bajo contraste)
}

# Rango para el tamaño de fuente variable del texto (predictor continuo)
MIN_TAMANO_VALOR = 18
MAX_TAMANO_VALOR = 38

TOTAL_PRUEBAS = 40  
ARCHIVO_RESULTADOS = "resultados_estudio_industrial.csv"

CELDA_W, CELDA_H = 160, 100
FUENTE_ETIQUETA = 12
ALTO_ENCABEZADO = 110
MARGEN_CELDA = 8
Y_ETIQUETA = 8
Y_VALOR = 36
ETIQUETA_OBJETIVO = "N.º DE LOTE"

COLOR_BORDE = (180, 180, 180)
COLOR_ETIQUETA = (60, 60, 60)
COLOR_ENCABEZADO = (20, 50, 100)

TIEMPO_LIMITE_S = 90
PANTALLA_COMPLETA = True
FONDO_VENTANA = "#1e1e1e"


# ======================================================================
# CAMPOS DE LA FICHA (Catálogo completo de producción industrial)
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
# UTILIDADES DE DISEÑO Y GEOMETRÍA
# ======================================================================
def cargar_fuente(tam):
    candidatas = [
        "arialbd.ttf", "Arial Bold.ttf", "C:/Windows/Fonts/arialbd.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
    ]
    for ruta in candidatas:
        try:
            return ImageFont.truetype(ruta, tam)
        except OSError:
            continue
    raise RuntimeError("No se encontró fuente en negrita.")

F_ETIQ = cargar_fuente(FUENTE_ETIQUETA)
F_TITULO = cargar_fuente(40)
F_ID = cargar_fuente(26)

PX_A_MM = 25.4 / DPI
W = round(ANCHO_MM / 25.4 * DPI)
H = round(ALTO_MM / 25.4 * DPI)

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
# GENERACIÓN DE DISEÑOS CON CONTRASTE Y TAMAÑO DE FUENTE
# ======================================================================
def disenar_cartel(cartel_id, n, nivel, fraccion, nivel_contraste, tam_fuente_objetivo, rng):
    cols, filas = NIVELES_N[n]
    f_obj, c_obj, d_obj = elegir_celda_objetivo(cols, filas, fraccion, rng)
    valor_obj = f"L-{rng.randint(10000, 99999)}"
    
    otros = rng.sample(CAMPOS, n - 1)
    
    # Extraer colores según el nivel de contraste seleccionado
    colores_cfg = CONFIG_CONTRASTES[nivel_contraste]
    color_fondo_rgb = colores_cfg["fondo"]
    color_texto_base = colores_cfg["texto"]

    contenido = []
    idx_obj = f_obj * cols + c_obj
    
    for i in range(n):
        if i == idx_obj:
            etiqueta = ETIQUETA_OBJETIVO
            valor = valor_obj
            tam = tam_fuente_objetivo
            col_txt = color_texto_base
        else:
            et, gen = otros.pop(0)
            etiqueta = et
            valor = gen(rng)
            tam = rng.randint(MIN_TAMANO_VALOR, MAX_TAMANO_VALOR)
            col_txt = color_texto_base
        
        contenido.append({
            "etiqueta": etiqueta,
            "valor": valor,
            "tamano_fuente": tam,
            "color_texto": col_txt
        })

    return {
        "cartel_id": cartel_id,
        "n_elementos": n,
        "columnas": cols,
        "filas": filas,
        "nivel_posicion": nivel,
        "fila_obj": f_obj,
        "col_obj": c_obj,
        "dist_centro_cm": round(d_obj * PX_A_MM / 10, 1),
        "nivel_contraste": nivel_contraste,
        "color_fondo_rgb": color_fondo_rgb,
        "tam_fuente_obj": tam_fuente_objetivo,
        "valor_objetivo": valor_obj,
        "contenido": contenido,
    }


def dibujar_cartel(d):
    cols, filas = d["columnas"], d["filas"]
    x0, y0, _, _ = geometria(cols, filas)
    img = Image.new("RGB", (W, H), d["color_fondo_rgb"])
    draw = ImageDraw.Draw(img)

    draw.rectangle([(0, 0), (W, ALTO_ENCABEZADO)], fill=COLOR_ENCABEZADO)
    draw.text((40, 28), "ORDEN DE PRODUCCIÓN - CONTROL INDUSTRIAL", fill="white", font=F_TITULO)
    texto_id = f"ID {d['cartel_id']}"
    ancho_id = draw.textlength(texto_id, font=F_ID)
    draw.text((W - ancho_id - 40, 36), texto_id, fill=(200, 210, 230), font=F_ID)

    margen = MARGEN_CELDA
    for i, celda in enumerate(d["contenido"]):
        f, c = divmod(i, cols)
        x = x0 + c * CELDA_W
        y = y0 + f * CELDA_H
        draw.rectangle([(x, y), (x + CELDA_W, y + CELDA_H)], outline=COLOR_BORDE, width=2)
        
        try:
            f_val_dinamica = ImageFont.truetype("arial.ttf", celda["tamano_fuente"])
        except OSError:
            f_val_dinamica = ImageFont.load_default()

        draw.text((x + margen, y + Y_ETIQUETA), celda["etiqueta"], fill=COLOR_ETIQUETA, font=F_ETIQ)
        draw.text((x + margen, y + Y_VALOR), celda["valor"], fill=celda["color_texto"], font=f_val_dinamica)
    return img


def construir_diseno():
    rng = random.Random(SEMILLA)
    disenos, contador = [], 0
    for n in NIVELES_N:
        for nivel, fraccion in NIVELES_POSICION.items():
            for nivel_c in CONFIG_CONTRASTES.keys():
                for tam_f in [22, 32]:
                    contador += 1
                    disenos.append(disenar_cartel(f"C{contador:03d}", n, nivel, fraccion, nivel_c, tam_f, rng))
    
    secuencia = [d["cartel_id"] for d in disenos]
    rng.shuffle(secuencia)
    return disenos, secuencia[:TOTAL_PRUEBAS]


# ======================================================================
# GESTIÓN DE RESULTADOS (CSV)
# ======================================================================
COLUMNAS = [
    "id_prueba", "cartel_id", "n_elementos", "nivel_posicion", "dist_centro_cm", 
    "nivel_contraste", "tamano_fuente_obj", "tiempo_s", "errores", "encontrado", "hora"
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
# APLICACIÓN INTERACTIVA TKINTER (FLUJO CONTINUO SIN FORMULARIO DE USUARIO)
# ======================================================================
class App:
    def __init__(self, disenos, secuencia):
        self.por_id = {d["cartel_id"]: d for d in disenos}
        self.secuencia = list(secuencia)
        self.root = tk.Tk()
        self.root.title("Estudio Cuasi-Experimental de Búsqueda Visual")
        self.root.configure(bg=FONDO_VENTANA)
        if PANTALLA_COMPLETA:
            self.root.attributes("-fullscreen", True)
        self.root.bind("<Escape>", lambda e: self.root.attributes("-fullscreen", False))
        self.root.bind("<Control-q>", lambda e: self.root.destroy())

        self.panel = None
        self.mostrar_instruccion()

    def limpiar(self):
        if self.panel is not None:
            self.panel.destroy()
        self.panel = tk.Frame(self.root, bg=FONDO_VENTANA)
        self.panel.pack(fill="both", expand=True)
        return self.panel

    def mostrar_instruccion(self):
        p = self.limpiar()
        self.id_prueba = siguiente_id()
        if self.id_prueba > len(self.secuencia):
            self.mostrar_fin_total()
            return
        
        self.diseno = self.por_id[self.secuencia[self.id_prueba - 1]]

        centro = tk.Frame(p, bg=FONDO_VENTANA)
        centro.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(centro, text=f"Prueba {self.id_prueba} de {len(self.secuencia)}", font=("Arial", 22, "bold"), bg=FONDO_VENTANA, fg="#7fb2ff").pack(pady=(0, 10))
        tk.Label(centro, text="En la orden de producción que verá a continuación, busque:", font=("Arial", 18), bg=FONDO_VENTANA, fg="white").pack(pady=10)
        tk.Label(centro, text=ETIQUETA_OBJETIVO, font=("Arial", 44, "bold"), bg=FONDO_VENTANA, fg="#ffd34d").pack(pady=15)
        tk.Label(centro, text="Haga clic sobre su valor correspondiente lo más rápido posible.", font=("Arial", 16), bg=FONDO_VENTANA, fg="#cccccc").pack(pady=10)
        
        tk.Button(centro, text="Comenzar Prueba (ESPACIO)", font=("Arial", 16, "bold"), command=self.mostrar_cartel).pack(pady=30)
        self.root.bind("<space>", lambda e: self.mostrar_cartel())

    def mostrar_cartel(self):
        self.root.unbind("<space>")
        p = self.limpiar()
        self.root.update_idletasks()
        cw, ch = max(self.root.winfo_width(), 400), max(self.root.winfo_height(), 300)
        self.escala = min(cw / W, ch / H)
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
        self.t0 = time.perf_counter()

    def clic(self, event):
        if not self.activo: return
        t = time.perf_counter() - self.t0
        d = self.diseno
        x0, y0, _, _ = geometria(d["columnas"], d["filas"])
        px = (event.x - self.off_x) / self.escala
        py = (event.y - self.off_y) / self.escala
        col = math.floor((px - x0) / CELDA_W)
        fila = math.floor((py - y0) / CELDA_H)
        
        if not (0 <= col < d["columnas"] and 0 <= fila < d["filas"]): return
        if (fila, col) == (d["fila_obj"], d["col_obj"]):
            self.terminar(t, encontrado=1)
        else:
            self.errores += 1

    def terminar(self, tiempo, encontrado):
        self.activo = False
        d = self.diseno
        fila = {
            "id_prueba": self.id_prueba,
            "cartel_id": d["cartel_id"],
            "n_elementos": d["n_elementos"],
            "nivel_posicion": d["nivel_posicion"],
            "dist_centro_cm": d["dist_centro_cm"],
            "nivel_contraste": d["nivel_contraste"],
            "tamano_fuente_obj": d["tam_fuente_obj"],
            "tiempo_s": round(tiempo, 3),
            "errores": self.errores,
            "encontrado": encontrado,
            "hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        guardar_fila(fila)
        self.mostrar_instruccion()

    def mostrar_fin_total(self):
        p = self.limpiar()
        centro = tk.Frame(p, bg=FONDO_VENTANA)
        centro.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(centro, text="¡Estudio Cuasi-Experimental Finalizado!", font=("Arial", 32, "bold"), bg=FONDO_VENTANA, fg="#7fe08a").pack(pady=20)
        tk.Label(centro, text=f"Los datos recolectados se guardaron en {ARCHIVO_RESULTADOS}", font=("Arial", 16), bg=FONDO_VENTANA, fg="white").pack(pady=10)
        tk.Button(centro, text="Cerrar Aplicación", font=("Arial", 16, "bold"), command=self.root.destroy).pack(pady=20)

    def ejecutar(self):
        self.root.mainloop()


def main():
    disenos, secuencia = construir_diseno()
    App(disenos, secuencia).ejecutar()


if __name__ == "__main__":
    main()