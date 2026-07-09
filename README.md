# Banco de Ejecutores del SGR

Análisis de datos y plataforma web para el proyecto **Sistema de Análisis de Riesgo (SAR)** — Sistema General de Regalías, DNP.

El proyecto tiene dos partes que viven en este mismo repo:

1. **Análisis / metodología** (`notebooks/`, `src/ics/`): construcción y calibración del Índice de Cumplimiento Histórico (ICS), que da un puntaje de riesgo 0-100 a cada entidad ejecutora del SGR.
2. **Plataforma web** (`plataforma/`): aplicación Flask que consulta ese puntaje por código de ejecutor, muestra comparables del mismo departamento, y un tablero descriptivo agregado.

---

## Estructura del repositorio

```
Banco-de-Ejecutores/
│
├── data/                          ← NO sube a GitHub (.gitignore)
│   ├── raw/                       ← Archivos originales del SGR, no modificar
│   │   ├── Balance_seguimiento_SGR.xlsx
│   │   ├── Curva_S_total_proyectos_en_Ejecucion_y_terminados.xlsx
│   │   ├── Curva_Sl_*.xlsx
│   │   └── Reprogramaciones_no_permitidas.xlsx
│   └── processed/                 ← Datos limpios / transformados
│
├── notebooks/
│   ├── feature_selection/         ← Selección de variables
│   │   ├── 01_prueba_base.ipynb
│   │   ├── 02_prueba_igpr.ipynb
│   │   └── 03_resultado_final.ipynb   ← ✅ NOTEBOOK PRINCIPAL
│   ├── sgr/                       ← Análisis SGR general
│   │   └── 01_pipeline_sgr.ipynb
│   └── archive/                   ← Experimentos descartados
│
├── src/                           ← Funciones reutilizables en Python
│   ├── preprocessing.py
│   ├── variable_selection.py
│   ├── clustering.py
│   └── ics/                       ← Cálculo del Índice de Cumplimiento Histórico
│       ├── etl_normalizar_excels.py          # 4 excels fuente -> Excel maestro normalizado
│       ├── indicador_cumplimiento_historico.py     # v1 (tau=5%, normalización min-max) — referencia
│       ├── indicador_cumplimiento_historico_v2.py  # v2 (tau=10%, normalización por percentil) — la vigente
│       ├── analisis_sensibilidad_tau.py      # compara resultados variando tau
│       └── analizar_resultado_ics.py         # KPIs + gráficas de cualquier resultado del ICS
│
├── outputs/
│   ├── rankings/                  ← Rankings de variables
│   └── plots/                     ← Gráficas generadas
│
├── docs/
│   ├── decisiones_metodologicas.md
│   └── mockup.html                ← mockup original de la plataforma
│
├── plataforma/                    ← Aplicación web (Flask), autocontenida
│   ├── app/
│   │   ├── main.py                # backend: API + servir el frontend
│   │   └── static/
│   │       ├── index.html         # frontend (adaptado del mockup)
│   │       ├── app.js             # fetch a la API + render dinámico
│   │       └── styles.css
│   ├── data/                      ← ⚠️ EXCEPCIÓN al .gitignore de arriba (ver nota)
│   │   ├── EXCEL_MAESTRO_ICS.xlsx           # salida de src/ics/etl_normalizar_excels.py
│   │   └── Resultado_ICS_v2_corregido.xlsx  # salida de src/ics/indicador_cumplimiento_historico_v2.py
│   ├── build_db.py                # arma data/sar.db (SQLite) a partir de los 2 excels de arriba
│   ├── requirements.txt
│   ├── render.yaml                # config de despliegue en Render
│   └── README.md                  # documentación específica de la plataforma
│
└── README.md                      ← este archivo
```

### Nota sobre `.gitignore`

La carpeta `data/` general no sube a GitHub porque contiene los excels *crudos* del SGR (30-40MB cada uno, son insumos). Pero `plataforma/data/` solo contiene los 2 archivos *ya procesados* (~20MB en total) que la plataforma necesita para construir su base de datos en cada despliegue — sin ellos, Render no puede generar `sar.db`. Por eso se excluye puntualmente del ignore general:

```gitignore
data/
!plataforma/data/
```

---

## Metodología: Índice de Cumplimiento Histórico (ICS)

Da un puntaje de riesgo (0-100) a cada entidad ejecutora del SGR, calculado como:

```
ICS = TBC × FC × Pen
```

| Componente | Qué mide | Fórmula |
|---|---|---|
| **TBC** — Tasa Bruta de Cumplimiento | % de periodos donde la ejecución no se desvió más del τ (tau) de lo programado | periodos cumplidos / periodos evaluados |
| **FC** — Factor de Carga | Complejidad operativa del ejecutor (más proyectos / más valor = más carga) | `ln(1+N) + ln(1 + Ve/Vref)` |
| **Pen** — Penalización | Descuento por reprogramaciones no permitidas | `1 / (1 + p·ln(1+R))` |

El ICS crudo se normaliza dentro de cada grupo de capacidad institucional y se invierte a un puntaje de riesgo 0-100 (0 = mejor, 100 = peor), con 3 niveles: Bajo / Medio / Alto.

### v1 vs. v2 — por qué existen dos versiones

La v1 (tau=5%, normalización min-max) producía que el **90% de los ejecutores** cayeran en "Riesgo Alto". El diagnóstico (ver `docs/decisiones_metodologicas.md`) encontró que la causa principal no era el tau sino la normalización min-max, extremadamente sensible a outliers (un solo ejecutor gigante define el techo de toda la escala). La v2 corrige:

- `tau`: 5% → 10% (más realista frente a la desviación mediana real, ~19%)
- Normalización: min-max → **percentil dentro del grupo** (`rank(pct=True)`), insensible a outliers

Resultado: de 90.4% / 7.9% / 1.7% (Alto/Medio/Bajo) pasa a 39.8% / 30.0% / 30.2% — mucho más balanceado y creíble.

**La plataforma web usa exclusivamente la v2** (`Resultado_ICS_v2_corregido.xlsx`). La v1 se conserva en `src/ics/` solo como referencia histórica del diagnóstico.

### Pendientes de calibración (documentados, no implementados aún)

- Separar N (número de proyectos) en "En Ejecución" vs. "Terminados"
- Ponderar la penalización por reprogramación según el valor del proyecto, no solo el conteo
- Evaluar el cumplimiento sobre el valor acumulado en vez del mensual puntual (reduce ruido de calendario)
- Winsorizar Ve y N antes de calcular FC, como capa adicional de robustez

---

## Cómo correr el análisis (notebooks)

### Google Colab

1. Sube el repo a Drive en: `MyDrive/Banco-de-Ejecutores/`
2. Pon los datos en: `MyDrive/Banco-de-Ejecutores/data/raw/`
3. Abre `notebooks/feature_selection/03_resultado_final.ipynb` en Colab
4. La celda de configuración monta Drive y resuelve las rutas automáticamente
5. Si tu ruta en Drive es diferente, cambia esta línea:
   ```python
   BASE_DIR = pathlib.Path('/content/drive/MyDrive/Banco-de-Ejecutores')
   ```

### Local (VS Code / Jupyter)

```bash
git clone https://github.com/<usuario>/Banco-de-Ejecutores.git
cd Banco-de-Ejecutores
mkdir -p data/raw
# Copia tus archivos Excel en data/raw/
pip install -r requirements.txt
```

### Archivos de datos esperados en `data/raw/`

| Archivo | Descripción |
|---|---|
| `Base de proyectos _15_12_2025.xlsx` | Base principal de proyectos SGR |
| `Capacidades_institucionales.xlsx` | Capacidades de los ejecutores |
| `ResultadosIGPR.xlsx` | Resultados IGPR por entidad |
| `Balance_seguimiento_SGR.xlsx` | Catálogo de ejecutores y proyectos (para `src/ics/`) |
| `Curva_S_total_proyectos_en_Ejecucion_y_terminados.xlsx` | Avance mes a mes por proyecto |
| `Curva_Sl_*.xlsx` | Segundo corte de avance mes a mes |
| `Reprogramaciones_no_permitidas.xlsx` | Reprogramaciones no permitidas por BPIN |

### Pipeline de cálculo del ICS (`src/ics/`)

```bash
cd src/ics
python etl_normalizar_excels.py              # 4 excels -> EXCEL_MAESTRO_ICS.xlsx
python indicador_cumplimiento_historico_v2.py  # -> Resultado_ICS_v2_corregido.xlsx
python analizar_resultado_ics.py Resultado_ICS_v2_corregido.xlsx --hoja Sheet1  # KPIs + gráficas
```

---

## Cómo correr la plataforma web

```bash
cd plataforma
pip install -r requirements.txt

# copiar a plataforma/data/:
#   EXCEL_MAESTRO_ICS.xlsx           (salida de etl_normalizar_excels.py)
#   Resultado_ICS_v2_corregido.xlsx  (salida de indicador_cumplimiento_historico_v2.py)

python build_db.py     # construye plataforma/data/sar.db
python app/main.py     # sirve en http://localhost:8000
```

Búsqueda en la plataforma: **por código de ejecutor** (no por BPIN) — ej. `73000` para Gobernación del Tolima.

### Actualizar los datos de la plataforma

Cada vez que haya un nuevo corte:

1. `python src/ics/etl_normalizar_excels.py` sobre los 4 excels fuente → `EXCEL_MAESTRO_ICS.xlsx`
2. `python src/ics/indicador_cumplimiento_historico_v2.py` → `Resultado_ICS_v2_corregido.xlsx`
3. Copiar ambos a `plataforma/data/`
4. `python plataforma/build_db.py` → regenera `sar.db`
5. Commit + push (Render redespliega solo)

### Desplegar en Render

1. Sube el repo a GitHub (ver sección siguiente)
2. En Render: **New → Web Service**, conecta el repo
3. En la configuración del servicio (o en `plataforma/render.yaml`), define `rootDir: plataforma` ya que es un monorepo
4. Render detecta el build/start command de `render.yaml` automáticamente
5. Cada `git push` a `main` redespliega solo (el build corre `build_db.py`, así que solo hace falta mantener actualizados los 2 excels en `plataforma/data/`)

### Limitaciones conocidas de la plataforma

- **Capacidad Financiera e Institucional**: no integradas — requieren datos de SECOP, patrimonio y entes de control que hoy no están disponibles. Solo Capacidad Administrativa se aproxima a partir del TBC, marcado explícitamente como provisional en la UI.
- **Serie histórica por vigencia**: solo existe un corte de datos calculado; la plataforma lo indica en vez de mostrar una tendencia inventada. Se habilitará cuando se empiecen a guardar snapshots del `resultado_ics` por fecha de corte.
- **Sector y nombre de proyecto**: se agregaron al ETL pero requieren re-correr `etl_normalizar_excels.py` sobre el archivo completo de `Balance_seguimiento_SGR.xlsx` (no la muestra).
- **4 bandas de riesgo en el gauge** (Bajo/Medio/Alto/Crítico): la metodología v2 define 3 niveles; para calzar con el gauge de 4 colores del mockup se subdividió "Alto" en cortes de 30/60/85 puntos (`nivel_4_bandas()` en `plataforma/app/main.py`). Ajustar si el DNP define oficialmente otros cortes.

---

## Agregar un nuevo análisis

```bash
# 1. Crear carpeta temática dentro de notebooks/
mkdir notebooks/nombre_del_tema

# 2. Nombrar notebooks con prefijo numérico
# 01_exploración.ipynb → 02_modelo.ipynb → 03_resultado_final.ipynb

# 3. Si hay funciones reutilizables, moverlas a src/
# 4. Outputs van a outputs/rankings/ o outputs/plots/
# 5. Documentar decisiones en docs/decisiones_metodologicas.md
```

## Si cambia la metodología del ICS

La plataforma no conoce la fórmula del indicador — solo lee una tabla con columnas fijas (`codigo_ejecutor`, `puntaje_riesgo`, `nivel_riesgo`, etc.). Mientras el resultado nuevo mantenga esa forma, el único cambio es apuntar `plataforma/build_db.py` al nuevo archivo de resultado. Si cambia la *estructura* (nuevas capacidades, otros componentes), el ajuste queda acotado a `plataforma/app/main.py` (función `construir_perfil`) y las cajas de "Componentes"/"Capacidades" en el frontend — el resto de la plataforma (búsqueda, comparables, tablero descriptivo) no depende de la fórmula interna.

---

## Autores

Proyecto SAR — Dirección de Regalías, DNP
Análisis: Daniel Felipe Rambaut Lemus