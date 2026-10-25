"""
Generador de carteles industriales para el estudio de localización de información (RLM).

Cada cartel simula una FICHA DE CONTROL DE PRODUCCIÓN con campos (etiqueta + valor).
La persona debe encontrar el valor del campo "N.º DE LOTE" y señalarlo con el dedo.

Factores MANIPULADOS (los únicos dos que varía el grupo):
  x1 = cantidad de elementos de información (campos)  -> 16, 36, 64, 100
  x2 = distancia del campo objetivo al centro del cartel (cm) -> centro, media, periferia

Todo lo demás es CONSTANTE: tipo y tamaño de letra, colores, fondo, formato del objetivo
y tamaño de las celdas. Así, las demás variables (distancia de la persona, edad, lentes,
cansancio) quedan como variables OBSERVADAS.

Salidas (carpeta carteles_salida/):
  - cartel_XX.png            un archivo por cartel (A3 horizontal a 100 dpi)
  - carteles_A3.pdf          todos los carteles listos para imprimir (imprimir al 100 %, sin "ajustar")
  - diseno_carteles.csv      características exactas de cada cartel (x1, x2 en cm, valor objetivo...)
  - hoja_recoleccion.csv     una fila por participante, con cartel asignado en orden aleatorio
  - carteles_piloto_A3.pdf   (solo en MODO = "piloto") los 15 carteles de la premuestra, en orden

Uso:  python generador_carteles.py        (requiere:  pip install pillow)
"""
import csv
import math
import os
import random

from PIL import Image, ImageDraw, ImageFont

# ======================================================================
# CONFIGURACIÓN (modificar solo esta sección)
# ======================================================================
SEMILLA = 2026                 # misma semilla = mismos carteles (reproducible)
SALIDA = "carteles_salida"
DPI = 100
ANCHO_MM, ALTO_MM = 420, 297   # A3 horizontal

# cantidad de elementos -> (columnas, filas)
NIVELES_N = {16: (4, 4), 36: (6, 6), 64: (8, 8), 100: (10, 10)}
# nivel de posición -> fracción de la distancia máxima al centro (0 = centro, 1 = esquina)
NIVELES_POSICION = {"centro": 0.0, "media": 0.5, "periferia": 1.0}
REPLICAS = 3                   # carteles distintos por combinación (contenido aleatorio distinto)
PARTICIPANTES = 100            # filas de la hoja de recolección en modo "estudio"

# MODO = "piloto": premuestra de 15 personas (cubre las 12 combinaciones de cartel)
# MODO = "estudio": muestra completa (PARTICIPANTES)
MODO = "piloto"
PARTICIPANTES_PILOTO = 15

CELDA_W, CELDA_H = 160, 100    # px (constantes en todos los carteles)
FUENTE_ETIQUETA = 16           # px
FUENTE_VALOR = 30              # px
ALTO_ENCABEZADO = 110          # px
MARGEN_CELDA = 8             # px, margen izquierdo del texto dentro de la celda
Y_ETIQUETA = 10               # px, posición vertical de la etiqueta dentro de la celda
Y_VALOR = 42                  # px, posición vertical del valor dentro de la celda
ETIQUETA_OBJETIVO = "N.º DE LOTE"
PREGUNTA = "Encuentre el valor de N.º DE LOTE y señálelo con el dedo"

# Colores fijos (no se aleatorizan)
COLOR_FONDO = (255, 255, 255)
COLOR_BORDE = (190, 190, 190)
COLOR_ETIQUETA = (70, 70, 70)
COLOR_VALOR = (0, 0, 0)
COLOR_ENCABEZADO = (20, 50, 100)

# ======================================================================
# CAMPOS DE LA FICHA (etiqueta, generador de valor). Se eligen al azar los necesarios.
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
# FUENTES
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
    raise RuntimeError(
        "No se encontró una fuente en negrita. Instale Arial o Liberation Sans, "
        "o agregue la ruta de una fuente .ttf en cargar_fuente()."
    )


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


# ======================================================================
# GEOMETRÍA Y DISEÑO
# ======================================================================
def geometria(cols, filas):
    bloque_w, bloque_h = cols * CELDA_W, filas * CELDA_H
    x0 = (W - bloque_w) // 2
    y0 = ALTO_ENCABEZADO + (H - ALTO_ENCABEZADO - bloque_h) // 2
    return x0, y0, x0 + bloque_w / 2, y0 + bloque_h / 2


def elegir_celda_objetivo(cols, filas, fraccion, rng):
    """Elige la celda cuya distancia al centro del bloque se acerca más a fraccion*distancia máxima."""
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


def verificar_ancho(draw, texto, fuente, maximo, descripcion):
    ancho = draw.textlength(texto, font=fuente)
    if ancho > maximo:
        raise ValueError(f"{descripcion} '{texto}' no cabe en la celda ({ancho:.0f} px > {maximo} px)")


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


def generar_cartel(cartel_id, n, fraccion, rng):
    cols, filas = NIVELES_N[n]
    x0, y0, cx, cy = geometria(cols, filas)
    img = Image.new("RGB", (W, H), COLOR_FONDO)
    draw = ImageDraw.Draw(img)

    # Encabezado fijo
    draw.rectangle([(0, 0), (W, ALTO_ENCABEZADO)], fill=COLOR_ENCABEZADO)
    draw.text((40, 28), "FICHA DE CONTROL DE PRODUCCIÓN", fill="white", font=F_TITULO)
    texto_id = f"ID {cartel_id}"
    ancho_id = draw.textlength(texto_id, font=F_ID)
    draw.text((W - ancho_id - 40, 36), texto_id, fill=(200, 210, 230), font=F_ID)

    # Contenido: el objetivo + (n-1) campos al azar
    f_obj, c_obj, d_obj = elegir_celda_objetivo(cols, filas, fraccion, rng)
    valor_obj = f"L-{rng.randint(10000, 99999)}"
    otros = rng.sample(CAMPOS, n - 1)
    contenido = [(et, gen(rng)) for et, gen in otros]
    rng.shuffle(contenido)
    contenido.insert(f_obj * cols + c_obj, (ETIQUETA_OBJETIVO, valor_obj))

    margen = MARGEN_CELDA
    for i, (etiqueta, valor) in enumerate(contenido):
        f, c = divmod(i, cols)
        x = x0 + c * CELDA_W
        y = y0 + f * CELDA_H
        verificar_ancho(draw, etiqueta, F_ETIQ, CELDA_W - 2 * margen, "Etiqueta")
        verificar_ancho(draw, valor, F_VALOR, CELDA_W - 2 * margen, "Valor")
        draw.rectangle([(x, y), (x + CELDA_W, y + CELDA_H)], outline=COLOR_BORDE, width=2)
        draw.text((x + margen, y + Y_ETIQUETA), etiqueta, fill=COLOR_ETIQUETA, font=F_ETIQ)
        draw.text((x + margen, y + Y_VALOR), valor, fill=COLOR_VALOR, font=F_VALOR)

    archivo = os.path.join(SALIDA, f"cartel_{cartel_id}.png")
    img.save(archivo, dpi=(DPI, DPI))

    return img, {
        "cartel_id": cartel_id,
        "n_elementos": n,
        "columnas": cols,
        "filas": filas,
        "nivel_posicion": "",  # se completa afuera
        "fila_obj": f_obj + 1,
        "col_obj": c_obj + 1,
        "dist_centro_cm": round(d_obj * PX_A_MM / 10, 1),
        "valor_objetivo": valor_obj,
        "pregunta": PREGUNTA,
        "letra_etiqueta_mm": round(FUENTE_ETIQUETA * PX_A_MM, 1),
        "letra_valor_mm": round(FUENTE_VALOR * PX_A_MM, 1),
        "ancho_mm": ANCHO_MM,
        "alto_mm": ALTO_MM,
        "archivo": archivo,
    }


# ======================================================================
# PROGRAMA PRINCIPAL
# ======================================================================
def main():
    rng = random.Random(SEMILLA)
    os.makedirs(SALIDA, exist_ok=True)

    imagenes, filas_diseno = [], []
    contador = 0
    for n in NIVELES_N:
        for nivel, fraccion in NIVELES_POSICION.items():
            for _ in range(REPLICAS):
                contador += 1
                cartel_id = f"C{contador:02d}"
                img, fila = generar_cartel(cartel_id, n, fraccion, rng)
                fila["nivel_posicion"] = nivel
                imagenes.append(img)
                filas_diseno.append(fila)

    # PDF para imprimir (A3, 100 % de escala)
    imagenes[0].save(
        os.path.join(SALIDA, "carteles_A3.pdf"),
        save_all=True,
        append_images=imagenes[1:],
        resolution=DPI,
    )

    # Diseño de los carteles
    with open(os.path.join(SALIDA, "diseno_carteles.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(filas_diseno[0].keys()))
        w.writeheader()
        w.writerows(filas_diseno)

    # Hoja de recolección: asignación aleatoria y balanceada de carteles a participantes
    ids = [d["cartel_id"] for d in filas_diseno]
    info = {d["cartel_id"]: d for d in filas_diseno}
    secuencia = []
    while len(secuencia) < PARTICIPANTES:
        bloque = ids[:]
        rng.shuffle(bloque)
        secuencia.extend(bloque)
    secuencia = secuencia[:PARTICIPANTES]
    if MODO == "piloto":
        secuencia = secuencia_piloto(filas_diseno)
        por_id = dict(zip(ids, imagenes))
        pdf_piloto = [por_id[c] for c in secuencia]
        pdf_piloto[0].save(
            os.path.join(SALIDA, "carteles_piloto_A3.pdf"),
            save_all=True,
            append_images=pdf_piloto[1:],
            resolution=DPI,
        )

    columnas = [
        "id_participante", "cartel_id", "n_elementos", "dist_centro_cm",
        "tiempo_s", "distancia_persona_m",
        "edad", "usa_lentes_0_1", "horas_sueno", "hora", "observaciones",
    ]
    with open(os.path.join(SALIDA, "hoja_recoleccion.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(columnas)
        for i, cid in enumerate(secuencia, start=1):
            w.writerow([i, cid, info[cid]["n_elementos"], info[cid]["dist_centro_cm"]] + [""] * 7)

    print(f"Listo: {len(imagenes)} carteles en '{SALIDA}/'")
    if MODO == "piloto":
        print(f"Modo piloto: imprima carteles_piloto_A3.pdf ({PARTICIPANTES_PILOTO} carteles, en el orden de hoja_recoleccion.csv).")
    print("Imprima carteles_A3.pdf al 100 % de escala (sin 'ajustar a la página').")


if __name__ == "__main__":
    main()