"""
Búsqueda de información — Formato de Factura Nítido (Clics Sincronizados y Pantalla Completa Robusta)
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

NIVELES_N = {12: 4, 24: 8, 36: 12}  # Número de líneas en la factura
NIVELES_POSICION = {"centro": 0.0, "media": 0.5, "periferia": 1.0}
REPLICAS = 3

# PREDICTOR: Niveles de Contraste Ergonómico (Escala de grises formal)
CONFIG_CONTRASTES = {
    "alto": {"fondo": (255, 255, 255), "texto": (0, 0, 0)},             # Negro sobre Blanco puro
    "medio": {"fondo": (235, 235, 235), "texto": (70, 70, 70)},        # Gris oscuro sobre Gris claro
    "bajo": {"fondo": (205, 205, 205), "texto": (110, 110, 110)}       # Gris medio sobre Gris medio-claro
}

# Tamaños de fuente que varían de un documento a otro (constantes dentro de la misma hoja)
NIVELES_TAMAÑO_FUENTE = [13, 17, 21, 25]

TOTAL_PRUEBAS = 40  
ARCHIVO_RESULTADOS = "resultados_factura_balanceada.csv"

# Dimensiones fijas de alta nitidez
W, H = 1200, 800

TIEMPO_LIMITE_S = 90
FONDO_VENTANA = "#1e1e1e"


# ======================================================================
# BANCO DE DATOS PARA LAS LÍNEAS DE LA FACTURA
# ======================================================================
DESCRIPCIONES_BASE = [
    "Talla pequeña traje de luces en rojo", "Mui grande churrolito", "Equipaje de Fútbol",
    "Servicio de consultoría técnica", "Mantenimiento preventivo de línea", "Suministro de válvulas de presión",
    "Kit de juntas tóricas de repuesto", "Lubricante sintético industrial", "Filtro de aire de alta eficiencia",
    "Panel de control digital v2", "Cable apantallado de cobre", "Sensor fotoeléctrico de proximidad",
    "Rodamiento de bolas sellado", "Correa de transmisión dentada", "Módulo de relés de seguridad"
]


# ======================================================================
# CARGA DE FUENTES SANS-SERIF (ARIAL NÍTIDA)
# ======================================================================
def cargar_fuente_sans(tam):
    candidatas = [
        "arialbd.ttf", "Arial Bold.ttf", "C:/Windows/Fonts/arialbd.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
    ]
    for ruta in candidatas:
        try:
            return ImageFont.truetype(ruta, int(tam))
        except OSError:
            continue
    raise RuntimeError("No se encontró fuente sans-serif (Arial).")

F_TITULO = cargar_fuente_sans(28)
F_SUB = cargar_fuente_sans(13)
F_TEXTO_BOLD = cargar_fuente_sans(11)


# ======================================================================
# GENERACIÓN DE DISEÑOS
# ======================================================================
def disenar_documento(doc_id, num_filas, nivel_posicion, fraccion_pos, nivel_contraste, tam_fuente_doc, rng):
    colores_cfg = CONFIG_CONTRASTES[nivel_contraste]
    
    desc_seleccionadas = rng.choices(DESCRIPCIONES_BASE, k=num_filas)
    
    fila_obj = rng.randint(0, num_filas - 1)
    valor_objetivo = f"REF-{rng.randint(1000, 9999)}"
    
    lineas_datos = []
    for idx, desc in enumerate(desc_seleccionadas):
        if idx == fila_obj:
            codigo = valor_objetivo
        else:
            while True:
                distractor = f"REF-{rng.randint(1000, 9999)}"
                if distractor != valor_objetivo:
                    codigo = distractor
                    break
                    
        cant = f"{rng.randint(1, 10)},00"
        precio = f"{rng.randint(10, 500)},00 €"
            
        lineas_datos.append({
            "codigo": codigo,
            "descripcion": desc,
            "unidades": cant,
            "precio": precio,
            "es_objetivo": (idx == fila_obj)
        })

    return {
        "doc_id": doc_id,
        "num_filas": num_filas,
        "nivel_posicion": nivel_posicion,
        "fila_obj": fila_obj,
        "nivel_contraste": nivel_contraste,
        "color_fondo_rgb": colores_cfg["fondo"],
        "color_texto_rgb": colores_cfg["texto"],
        "tamano_fuente_doc": tam_fuente_doc,
        "valor_objetivo": valor_objetivo,
        "lineas_datos": lineas_datos
    }


def dibujar_documento(d):
    img = Image.new("RGB", (W, H), d["color_fondo_rgb"])
    draw = ImageDraw.Draw(img)
    col_txt = d["color_texto_rgb"]

    # --- 1. ENCABEZADO ESTILO FACTURA ---
    draw.text((80, 40), "FACTURA", fill=col_txt, font=F_TITULO)
    draw.text((80, 80), "Fecha: 10/03/2026", fill=col_txt, font=F_SUB)
    draw.text((80, 100), f"Nº factura: {d['doc_id']}", fill=col_txt, font=F_SUB)

    draw.text((80, 140), "DATOS EMISOR", fill=col_txt, font=F_TEXTO_BOLD)
    draw.text((80, 160), "• Industrial S.A. - NIF: A0000000T", fill=col_txt, font=F_SUB)
    draw.text((80, 178), "• Polígono Industrial Norte, 25", fill=col_txt, font=F_SUB)

    draw.text((500, 140), "DATOS CLIENTE", fill=col_txt, font=F_TEXTO_BOLD)
    draw.text((500, 160), "• Cliente Comercial S.L.", fill=col_txt, font=F_SUB)
    draw.text((500, 178), "• Dirección de entrega sector B", fill=col_txt, font=F_SUB)

    # --- 2. TABLA CON COORDENADAS EXACTAS ---
    y_tabla = 220
    draw.line([(80, y_tabla), (W - 80, y_tabla)], fill=(150, 150, 150), width=2)
    
    draw.text((100, y_tabla + 8), "Código", fill=col_txt, font=F_TEXTO_BOLD)
    draw.text((260, y_tabla + 8), "Descripción", fill=col_txt, font=F_TEXTO_BOLD)
    draw.text((700, y_tabla + 8), "Unidades", fill=col_txt, font=F_TEXTO_BOLD)
    draw.text((850, y_tabla + 8), "Precio", fill=col_txt, font=F_TEXTO_BOLD)
    
    y_tabla += 30  # Fin de cabecera de tabla en 250
    draw.line([(80, y_tabla), (W - 80, y_tabla)], fill=(180, 180, 180), width=1)

    f_celda = cargar_fuente_sans(d["tamano_fuente_doc"])

    y_fila = 253
    for linea in d["lineas_datos"]:
        draw.text((100, y_fila), linea["codigo"], fill=col_txt, font=f_celda)
        draw.text((260, y_fila), linea["descripcion"], fill=col_txt, font=f_celda)
        draw.text((700, y_fila), linea["unidades"], fill=col_txt, font=f_celda)
        draw.text((850, y_fila), linea["precio"], fill=col_txt, font=f_celda)
        
        y_fila += 30  # Altura de cada fila exactamente 30 píxeles
        draw.line([(80, y_fila - 2), (W - 80, y_fila - 2)], fill=(210, 210, 210), width=1)

    return img


def construir_diseno():
    rng = random.Random(SEMILLA)
    disenos = []
    contador = 0
    
    contrastes_keys = list(CONFIG_CONTRASTES.keys())
    
    for num_f, cantidad in NIVELES_N.items():
        for nivel, fraccion in NIVELES_POSICION.items():
            for nivel_c in contrastes_keys:
                for tam_f in NIVELES_TAMAÑO_FUENTE:
                    contador += 1
                    disenos.append(disenar_documento(f"FAC-2026/{contador:03d}", cantidad, nivel, fraccion, nivel_c, tam_f, rng))
    
    rng.shuffle(disenos)
    return disenos, disenos[:TOTAL_PRUEBAS]


# ======================================================================
# GESTIÓN DE RESULTADOS (CSV)
# ======================================================================
COLUMNAS = [
    "id_prueba", "doc_id", "num_filas", "nivel_posicion", 
    "nivel_contraste", "tamano_fuente_doc", "distancia_usuario_m", "tiempo_s", "errores", "encontrado", "hora"
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
# APLICACIÓN INTERACTIVA TKINTER
# ======================================================================
class App:
    def __init__(self, disenos, secuencia):
        self.por_id = {d["doc_id"]: d for d in disenos}
        self.secuencia = [d["doc_id"] for d in secuencia]
        self.root = tk.Tk()
        self.root.title("Estudio Cuasi-Experimental - Factura Nitida")
        self.root.configure(bg=FONDO_VENTANA)
        
        # Pantalla completa robusta para Windows y otros sistemas
        self.root.state('zoomed')
        self.root.attributes("-fullscreen", True)
        
        self.root.bind("<Escape>", lambda e: self.root.attributes("-fullscreen", False))
        self.root.bind("<Control-q>", lambda e: self.root.destroy())

        self.ultimo_distancia = "1.0"
        self.panel = None
        self.mostrar_formulario_distancia()

    def limpiar(self):
        if self.panel is not None:
            self.panel.destroy()
        self.panel = tk.Frame(self.root, bg=FONDO_VENTANA)
        self.panel.pack(fill="both", expand=True)
        return self.panel

    def mostrar_formulario_distancia(self):
        p = self.limpiar()
        self.id_prueba = siguiente_id()
        if self.id_prueba > len(self.secuencia):
            self.mostrar_fin_total()
            return
        
        self.diseno = self.por_id[self.secuencia[self.id_prueba - 1]]

        centro = tk.Frame(p, bg=FONDO_VENTANA)
        centro.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(centro, text=f"Prueba {self.id_prueba} de {len(self.secuencia)}", font=("Arial", 22, "bold"), bg=FONDO_VENTANA, fg="#7fb2ff").pack(pady=(0, 10))
        tk.Label(centro, text="Ingrese la distancia del participante a la pantalla (metros):", font=("Arial", 16), bg=FONDO_VENTANA, fg="white").pack(pady=(0, 20))

        f_inputs = tk.Frame(centro, bg=FONDO_VENTANA)
        f_inputs.pack(pady=10)

        tk.Label(f_inputs, text="Distancia (m):", font=("Arial", 14), bg=FONDO_VENTANA, fg="white").grid(row=0, column=0, sticky="e", padx=10, pady=6)
        self.e_dist = tk.Entry(f_inputs, font=("Arial", 14), width=10)
        self.e_dist.insert(0, self.ultimo_distancia)
        self.e_dist.grid(row=0, column=1, sticky="w", pady=6)

        self.lbl_error = tk.Label(centro, text="", font=("Arial", 12), bg=FONDO_VENTANA, fg="#ff7b7b")
        self.lbl_error.pack(pady=10)

        tk.Button(centro, text="Continuar a Instrucción", font=("Arial", 14, "bold"), command=self.validar_y_continuar).pack(pady=20)
        self.e_dist.focus_set()

    def validar_y_continuar(self):
        try:
            dist = float(self.e_dist.get().strip().replace(",", "."))
            if not (0.2 <= dist <= 5.0): raise ValueError("Distancia fuera de rango (0.2m - 5m)")
        except ValueError as err:
            self.lbl_error.config(text=f"Error: {err}")
            return
        
        self.ultimo_distancia = self.e_dist.get()
        self.distancia_usuario_m = dist
        self.mostrar_instruccion()

    def mostrar_instruccion(self):
        p = self.limpiar()
        centro = tk.Frame(p, bg=FONDO_VENTANA)
        centro.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(centro, text="En la factura, busque el siguiente código:", font=("Arial", 18), bg=FONDO_VENTANA, fg="white").pack(pady=10)
        tk.Label(centro, text=self.diseno["valor_objetivo"], font=("Arial", 44, "bold"), bg=FONDO_VENTANA, fg="#ffd34d").pack(pady=15)
        tk.Label(centro, text="Haga clic sobre su fila correspondiente lo más rápido posible.", font=("Arial", 16), bg=FONDO_VENTANA, fg="#cccccc").pack(pady=10)
        
        tk.Button(centro, text="Comenzar Prueba (ESPACIO)", font=("Arial", 16, "bold"), command=self.mostrar_documento).pack(pady=30)
        self.root.bind("<space>", lambda e: self.mostrar_documento())

    def mostrar_documento(self):
        self.root.unbind("<space>")
        p = self.limpiar()
        
        imagen = dibujar_documento(self.diseno)
        self.foto = ImageTk.PhotoImage(imagen)
        
        self.canvas = tk.Canvas(p, bg=FONDO_VENTANA, highlightthickness=0, cursor="hand2")
        self.canvas.pack(fill="both", expand=True)
        
        # Actualizamos la geometría para obtener dimensiones reales del canvas expandido
        self.root.update_idletasks()
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        
        # Calculamos el offset exacto para centrar la factura de 1200x800
        self.offset_x = (canvas_w - W) // 2
        self.offset_y = (canvas_h - H) // 2
        
        self.canvas.create_image(self.offset_x, self.offset_y, anchor="nw", image=self.foto)
        self.canvas.bind("<Button-1>", self.clic)

        self.errores = 0
        self.activo = True
        self.t0 = time.perf_counter()

    def clic(self, event):
        if not self.activo: return
        t = time.perf_counter() - self.t0
        d = self.diseno
        
        # Restamos el offset exacto donde se dibujó la imagen
        click_y = event.y - self.offset_y
        
        # Coordenadas exactas sincronizadas con el dibujado de las filas (y_inicio = 253, altura = 30)
        y_inicio_filas = 253
        altura_fila = 30
        fila_clicada = math.floor((click_y - y_inicio_filas) / altura_fila)
        
        if 0 <= fila_clicada < len(d["lineas_datos"]):
            if fila_clicada == d["fila_obj"]:
                self.terminar(t, encontrado=1)
            else:
                self.errores += 1

    def terminar(self, tiempo, encontrado):
        self.activo = False
        d = self.diseno
        fila = {
            "id_prueba": self.id_prueba,
            "doc_id": d["doc_id"],
            "num_filas": d["num_filas"],
            "nivel_posicion": d["nivel_posicion"],
            "nivel_contraste": d["nivel_contraste"],
            "tamano_fuente_doc": d["tamano_fuente_doc"],
            "distancia_usuario_m": self.distancia_usuario_m,
            "tiempo_s": round(tiempo, 3),
            "errores": self.errores,
            "encontrado": encontrado,
            "hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        guardar_fila(fila)
        self.mostrar_formulario_distancia()

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