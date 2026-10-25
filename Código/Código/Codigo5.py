import tkinter as tk
from PIL import Image, ImageDraw, ImageFont, ImageTk
import random
import csv
import os
import time

# ====================================
# CONFIGURACIÓN
# ====================================

ANCHO = 1400
ALTO = 900

FILAS = 10
COLUMNAS = 10

MIN_NUMEROS = 50
MAX_NUMEROS = 100

MIN_TAMANO = 20
MAX_TAMANO = 65

ESCALA = 0.7

# ====================================
# CARPETAS Y ARCHIVO CSV GLOBAL
# ====================================

os.makedirs("dataset", exist_ok=True)

if not os.path.exists("resultados.csv"):
    with open(
        "resultados.csv",
        "w",
        newline="",
        encoding="utf8"
    ) as f:
        writer = csv.writer(f)
        writer.writerow([
            "hoja",
            "objetivo",
            "tamano_fuente",
            "tiempo",
            "errores"
        ])

# ====================================
# VARIABLES GLOBALES
# ====================================

contador = 0
objetivo = ""
tamano_actual_fuente = 0
errores = 0
inicio = 0
numeros = []
foto = None
terminado = False

# ====================================
# GUARDAR RESULTADO
# ====================================

def guardar_resultado():
    with open(
        "resultados.csv",
        "a",
        newline="",
        encoding="utf8"
    ) as f:
        writer = csv.writer(f)
        writer.writerow([
            contador,
            objetivo,
            tamano_actual_fuente,
            round(
                time.time() - inicio,
                2
            ),
            errores
        ])

# ====================================
# GENERAR HOJA
# ====================================

def generar():
    global contador
    global objetivo
    global tamano_actual_fuente
    global errores
    global inicio
    global numeros
    global foto
    global terminado

    terminado = False
    contador += 1
    errores = 0
    numeros = []

    # Definir un tamaño de fuente base aleatorio para toda la hoja o por elemento
    # Aquí lo manejaremos variando por elemento dentro del rango, o un tamaño general.
    # Vamos a registrar el tamaño promedio o el del objetivo, pero lo ideal es registrar el del objetivo o un valor representativo.
    # Usemos un tamaño aleatorio para la hoja o dejémoslo por elemento y guardemos el del objetivo.
    
    fondo = random.choice([
        "white",
        "lightyellow",
        "lavender",
        "aliceblue",
        "honeydew"
    ])

    img = Image.new(
        "RGB",
        (ANCHO, ALTO),
        fondo
    )

    draw = ImageDraw.Draw(img)

    celda_x = ANCHO // COLUMNAS
    celda_y = ALTO // FILAS

    cantidad = random.randint(
        MIN_NUMEROS,
        MAX_NUMEROS
    )

    posiciones = [
        (f, c)
        for f in range(FILAS)
        for c in range(COLUMNAS)
    ]

    random.shuffle(posiciones)

    objetivo = str(
        random.randint(0, 99999)
    )

    posicion_objetivo = random.randint(
        0,
        cantidad - 1
    )

    csv_path = f"dataset/hoja_{contador}.csv"

    with open(
        csv_path,
        "w",
        newline="",
        encoding="utf8"
    ) as archivo:
        writer = csv.writer(archivo)
        writer.writerow([
            "numero",
            "fila",
            "columna",
            "x",
            "y",
            "ancho",
            "alto",
            "tamano_fuente",
            "objetivo"
        ])

        # Colores mucho más variados
        colores = [
            (255, 0, 0),      # Rojo
            (0, 0, 255),      # Azul
            (0, 128, 0),      # Verde
            (128, 0, 128),    # Morado
            (255, 140, 0),    # Naranja
            (0, 139, 139),    # Cian oscuro
            (139, 69, 19),    # Marrón
            (255, 20, 147),   # Rosa profundo
            (47, 79, 79),     # Gris pizarra oscuro
            (75, 0, 130),     # Índigo
            (0, 0, 0)         # Negro
        ]

        for i in range(cantidad):
            fila, col = posiciones[i]

            if i == posicion_objetivo:
                numero = objetivo
                es_obj = True
            else:
                while True:
                    numero = str(
                        random.randint(
                            0,
                            99999
                        )
                    )
                    if numero != objetivo:
                        break
                es_obj = False

            color = random.choice(colores)

            tam = random.randint(
                MIN_TAMANO,
                MAX_TAMANO
            )

            # Guardar el tamaño si es el objetivo para el registro global
            if es_obj:
                tamano_actual_fuente = tam

            while True:
                try:
                    fuente = ImageFont.truetype(
                        "arial.ttf",
                        tam
                    )
                except:
                    fuente = ImageFont.load_default()

                bbox = draw.textbbox(
                    (0, 0),
                    numero,
                    font=fuente
                )

                ancho_t = bbox[2] - bbox[0]
                alto_t = bbox[3] - bbox[1]

                if (
                    ancho_t < celda_x - 10
                    and
                    alto_t < celda_y - 10
                ):
                    break

                tam -= 2
                if tam < 10:
                    break

            x = (
                col * celda_x
                + (celda_x - ancho_t) // 2
            )

            y = (
                fila * celda_y
                + (celda_y - alto_t) // 2
            )

            draw.text(
                (x, y),
                numero,
                fill=color,
                font=fuente
            )

            numeros.append({
                "numero": numero,
                "objetivo": es_obj,
                "x1": x,
                "y1": y,
                "x2": x + ancho_t,
                "y2": y + alto_t
            })

            writer.writerow([
                numero,
                fila,
                col,
                x,
                y,
                ancho_t,
                alto_t,
                tam,
                es_obj
            ])

    draw.rectangle(
        [
            (0, ALTO - 80),
            (ANCHO, ALTO)
        ],
        fill="black"
    )

    try:
        fuente2 = ImageFont.truetype(
            "arial.ttf",
            50
        )
    except:
        fuente2 = ImageFont.load_default()

    draw.text(
        (20, ALTO - 65),
        f"BUSCAR: {objetivo}",
        fill="white",
        font=fuente2
    )

    img.save(
        f"dataset/hoja_{contador}.png"
    )

    vista = img.resize(
        (
            int(ANCHO * ESCALA),
            int(ALTO * ESCALA)
        )
    )

    foto = ImageTk.PhotoImage(
        vista
    )

    canvas.delete("all")

    canvas.config(
        width=vista.width,
        height=vista.height
    )

    canvas.create_image(
        0,
        0,
        anchor="nw",
        image=foto
    )

    lbl_objetivo.config(
        text=f"BUSCAR: {objetivo}"
    )

    lbl_estado.config(
        text=f"Hoja #{contador}"
    )

    inicio = time.time()

# ====================================
# CLIC
# ====================================

def clic(event):
    global errores
    global terminado

    if terminado:
        return

    x = event.x / ESCALA
    y = event.y / ESCALA

    for n in numeros:
        if (
            n["x1"] <= x <= n["x2"]
            and
            n["y1"] <= y <= n["y2"]
        ):
            x1 = n["x1"] * ESCALA
            y1 = n["y1"] * ESCALA
            x2 = n["x2"] * ESCALA
            y2 = n["y2"] * ESCALA

            if n["objetivo"]:
                canvas.create_oval(
                    x1 - 10,
                    y1 - 10,
                    x2 + 10,
                    y2 + 10,
                    outline="green",
                    width=4
                )

                tiempo = round(
                    time.time() - inicio,
                    2
                )

                guardar_resultado()

                lbl_estado.config(
                    text=(
                        f"✅ Correcto | "
                        f"{tiempo}s | "
                        f"{errores} errores"
                    )
                )

                terminado = True

                ventana.after(
                    1200,
                    generar
                )

            else:
                canvas.create_line(
                    x1,
                    y1,
                    x2,
                    y2,
                    fill="red",
                    width=3
                )

                canvas.create_line(
                    x1,
                    y2,
                    x2,
                    y1,
                    fill="red",
                    width=3
                )

                errores += 1

                lbl_estado.config(
                    text=f"❌ Errores: {errores}"
                )

            return

# ====================================
# INTERFAZ
# ====================================

ventana = tk.Tk()

ventana.title(
    "Búsqueda Visual Infinita"
)

lbl_objetivo = tk.Label(
    ventana,
    text="",
    font=(
        "Arial",
        22,
        "bold"
    ),
    fg="blue"
)

lbl_objetivo.pack(
    pady=5
)

canvas = tk.Canvas(
    ventana
)

canvas.pack()

canvas.bind(
    "<Button-1>",
    clic
)

lbl_estado = tk.Label(
    ventana,
    font=("Arial", 12)
)

lbl_estado.pack(
    pady=10
)

btn = tk.Button(
    ventana,
    text="Generar nueva hoja",
    command=generar
)

btn.pack(
    pady=5
)

generar()

ventana.mainloop()