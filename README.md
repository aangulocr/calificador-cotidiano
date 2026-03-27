# Calificador de Trabajos Cotidianos

Herramienta de escritorio en Python para calificar automáticamente trabajos en Excel según la **escala MEP (0-3 por hoja)** basada en la Taxonomía de Bloom.

---

## Requisitos

- Python 3.9 o superior
- Instalar dependencias (una sola vez):

```powershell
pip install -r requirements.txt
```

---

## Cómo ejecutar

```powershell
python d:\REVISION_TC\main.py
```

---

## Preparar la carpeta de trabajo

```
Trabajos_Cotidianos\
├── PLANTILLA.xlsx            ← TU solucionario
├── Juan_Perez.xlsx
├── Maria_Lopez.xlsx
└── ...
```

> Los nombres de los archivos de estudiantes se usan como **Nombre del Estudiante** en el reporte.

---

## Preparar `PLANTILLA.xlsx`

### Paso 1 — Hoja `_CONFIG` (obligatoria)

Crea una hoja llamada **`_CONFIG`** (exactamente así) con los siguientes encabezados en la fila 1:

| hoja_nombre | hoja_indice | tema_id | celdas_evaluar | peso |
|---|---|---|---|---|
| Hoja1 | 1 | nombre_hoja | - | 1 |
| Hoja1 | 1 | color_hoja | - | 1 |
| Hoja1 | 1 | funciones_excel | B2:B10 | 2 |
| Hoja2 | 2 | tabla_dinamica | - | 3 |
| Hoja2 | 2 | filtros | - | 1 |
| Hoja3 | 3 | buscarv | C5:C15 | 2 |

**Columnas explicadas:**

| Columna | Descripción |
|---|---|
| `hoja_nombre` | Nombre exacto de la hoja (debe ser igual en los archivos del estudiante) |
| `hoja_indice` | Número de hoja (1-5) |
| `tema_id` | ID del tema (ver tabla abajo) |
| `celdas_evaluar` | Rango de celdas donde se evalúa (ej: `B2:C10`). Usa `-` si no aplica |
| `peso` | Puntos que vale ese tema en esa hoja |

### Paso 2 — Hojas del solucionario

Las hojas del patrón deben tener **los mismos nombres** que las hojas del estudiante, y contener las fórmulas correctas **escritas** (no valores ingresados a mano):

- ✅ `=BUSCARV(A2, $F$2:$H$20, 3, 0)` → el script la detecta
- ❌ `10` (valor calculado a mano) → el script no lo detectará

### IDs de temas disponibles (`tema_id`)

| tema_id | Descripción |
|---|---|
| `nombre_hoja` | Nombre de la pestaña |
| `color_hoja` | Color de la pestaña |
| `operaciones_basicas` | Fórmulas con +, -, *, / o Funciones Equivalentes (SUMA, PRODUCTO) |
| `concatenar` | CONCATENAR / CONCAT |
| `contar` | CONTAR |
| `contara` | CONTARA |
| `contar_si` | CONTAR.SI |
| `contar_blanco` | CONTAR.BLANCO |
| `promedio` | PROMEDIO |
| `mediana` | MEDIANA |
| `moda_uno` | MODA.UNO / MODE.SNGL |
| `max_min` | MAX y MIN |
| `si_simple` | SI simple |
| `si_anidado` | SI anidado |
| `buscarv` | BUSCARV |
| `sumar_si` | SUMAR.SI |
| `si_conjunto` | SI.CONJUNTO / IFS |
| `operaciones_combinadas` | Operaciones combinadas (Reglas de 3, etc.) |
| `calculo_porcentaje` | Cálculo de porcentaje (*0.13, /100, *13%) |
| `si_con_calculo` | SI / SI.CONJUNTO conteniendo operaciones aritméticas |
| `tabla_dinamica` | Tabla dinámica |
| `grafico_dinamico` | Gráfico dinámico |
| `grafico_normal` | Gráfico normal |
| `filtros` | Filtros automáticos |
| `formato_condicional` | Formato condicional |
| `validacion_datos` | Validación de datos |
| `bordes` | Bordes de celda |
| `relleno` | Relleno / Color de fondo |
| `color_fuente` | Color de fuente |
| `formato_moneda` | Formato de Moneda (simbolos o contabilidad) |
| `formato_fecha` | Formato de Fecha (ej. dd/mm/yyyy) |

---

## Escala MEP (Taxonomía de Bloom) por hoja

| % de logro | Puntos | Nivel |
|---|---|---|
| 0% | 0 | No presenta evidencia |
| > 0% y ≤ 33% | 1 | No logrado |
| > 33% y ≤ 66% | 2 | En proceso |
| > 66% | 3 | Logrado |

La **Nota Final** es la suma de los puntos de todas las hojas (máximo 3 × número de hojas).

---

## Reporte generado

El archivo `Resultados_Evaluacion.xlsx` incluye:

- Una fila por estudiante
- Puntos por hoja (0-3) con color: 🟢 Logrado / 🟡 En proceso / 🔴 No logrado
- Nota final con semáforo de color
- Columna de observaciones automáticas explicando qué falló
- Promedio del grupo al final
- Hoja **Leyenda** con la escala MEP

---

## Estructura del proyecto

```
d:\REVISION_TC\
├── main.py                   ← Ejecutar esto
├── requirements.txt
├── README.md
└── calificador\
    ├── __init__.py
    ├── config.py             ← Temas y escala MEP
    ├── comparadores.py       ← Evaluadores por tema
    ├── evaluador.py          ← Motor de evaluación
    └── exportador.py         ← Generador de reporte
```
# calificador-cotidiano
# calificador-cotidiano
# calificador-cotidiano
