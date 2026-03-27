"""
config.py — Constantes globales del calificador de trabajos cotidianos.
Escala MEP (Taxonomía de Bloom): 0-3 puntos por hoja.
"""

# ---------------------------------------------------------------------------
# Temas evaluables — id único, etiqueta visible y categoría para la GUI
# ---------------------------------------------------------------------------
TEMAS = [
    # --- Estructura ---
    {"id": "nombre_hoja",          "nombre": "Nombre de hoja",           "categoria": "Estructura"},
    {"id": "color_hoja",           "nombre": "Color de pestaña",         "categoria": "Estructura"},

    # --- Operaciones y Funciones ---
    {"id": "operaciones_basicas",  "nombre": "Operaciones básicas",      "categoria": "Funciones"},
    {"id": "concatenar",           "nombre": "CONCATENAR / CONCAT",      "categoria": "Funciones"},
    {"id": "contar",               "nombre": "CONTAR",                   "categoria": "Funciones"},
    {"id": "contara",              "nombre": "CONTARA",                  "categoria": "Funciones"},
    {"id": "contar_si",            "nombre": "CONTAR.SI",                "categoria": "Funciones"},
    {"id": "contar_blanco",        "nombre": "CONTAR.BLANCO",            "categoria": "Funciones"},
    {"id": "promedio",             "nombre": "PROMEDIO",                 "categoria": "Funciones"},
    {"id": "mediana",              "nombre": "MEDIANA",                  "categoria": "Funciones"},
    {"id": "moda_uno",             "nombre": "MODA.UNO",                 "categoria": "Funciones"},
    {"id": "max_min",              "nombre": "MAX / MIN",                "categoria": "Funciones"},
    {"id": "si_simple",            "nombre": "SI simple",                       "categoria": "Funciones"},
    {"id": "si_anidado",           "nombre": "SI anidado",                      "categoria": "Funciones"},
    {"id": "buscarv",              "nombre": "BUSCARV",                         "categoria": "Funciones"},
    {"id": "sumar_si",             "nombre": "SUMAR.SI",                        "categoria": "Funciones"},
    {"id": "si_conjunto",          "nombre": "SI.CONJUNTO / IFS",               "categoria": "Funciones"},
    {"id": "operaciones_combinadas", "nombre": "Operaciones combinadas",         "categoria": "Funciones"},
    {"id": "calculo_porcentaje",   "nombre": "Cálculo de porcentaje",            "categoria": "Funciones"},
    {"id": "si_con_calculo",       "nombre": "SI/SI.CONJUNTO con cálculo",       "categoria": "Funciones"},

    # --- Análisis y Visualización ---
    {"id": "tabla_dinamica",       "nombre": "Tabla dinámica",           "categoria": "Análisis"},
    {"id": "grafico_dinamico",     "nombre": "Gráfico dinámico",         "categoria": "Análisis"},
    {"id": "grafico_normal",       "nombre": "Gráfico normal",           "categoria": "Análisis"},
    {"id": "filtros",              "nombre": "Filtros automáticos",      "categoria": "Análisis"},
    {"id": "formato_condicional",  "nombre": "Formato condicional",      "categoria": "Análisis"},
    {"id": "validacion_datos",     "nombre": "Validación de datos",      "categoria": "Análisis"},

    # --- Formato ---
    {"id": "bordes",              "nombre": "Bordes de celdas",         "categoria": "Formato"},
    {"id": "relleno",             "nombre": "Relleno / Color de fondo",  "categoria": "Formato"},
    {"id": "color_fuente",        "nombre": "Color de fuente",           "categoria": "Formato"},
    {"id": "formato_moneda",      "nombre": "Formato de moneda",         "categoria": "Formato"},
    {"id": "formato_fecha",       "nombre": "Formato de fecha",          "categoria": "Formato"},
]

# Mapeo rápido id → nombre
TEMA_NOMBRE = {t["id"]: t["nombre"] for t in TEMAS}

# ---------------------------------------------------------------------------
# Escala de conversión — Porcentaje de logro → Puntos Bloom (0-3)
# ---------------------------------------------------------------------------
def porcentaje_a_bloom(pct: float) -> int:
    """
    Convierte un porcentaje de logro (0.0 - 100.0) a la escala MEP 0-3.

    0%          → 0  No presenta evidencia
    > 0% ≤ 33%  → 1  No logrado
    > 33% ≤ 66% → 2  En proceso
    > 66%       → 3  Logrado
    """
    if pct <= 0:
        return 0
    elif pct <= 33.0:
        return 1
    elif pct <= 66.0:
        return 2
    else:
        return 3


# ---------------------------------------------------------------------------
# Paleta de colores (consistente con las reglas del proyecto)
# ---------------------------------------------------------------------------
COLOR_PRIMARIO  = "#0056b3"
COLOR_FONDO     = "#f5f5f5"
COLOR_ALERTA    = "#dc3545"
COLOR_PROCESO   = "#ffc107"
COLOR_LOGRADO   = "#198754"
COLOR_NO_LOGO   = "#dc3545"

# Máximo de hojas por estudiante
MAX_HOJAS = 5

# Nombre de la hoja de configuración dentro del patrón
HOJA_CONFIG = "_CONFIG"

from pathlib import Path
BASE_DIR = Path(__file__).parent.parent

# Nombre del archivo patrón (ahora en la raíz del proyecto)
ARCHIVO_PATRON = "PLANTILLA.xlsx"
RUTA_PATRON = BASE_DIR / ARCHIVO_PATRON

# Nombre del archivo de resultados generado
ARCHIVO_RESULTADOS = "Resultados_Evaluacion.xlsx"
RUTA_RESULTADOS = BASE_DIR / ARCHIVO_RESULTADOS
