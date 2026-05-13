# Banco de Ejecutores del SGR

Análisis de datos para el proyecto **Sistema de Análisis de Riesgo (SAR)** — Sistema General de Regalías, DNP.

---

## Estructura

```
Banco-de-Ejecutores/
│
├── data/                         
│   ├── raw/                       ← Archivos originales, no modificar
│   └── processed/                 ← Datos limpios / transformados
│
├── notebooks/
│   ├── feature_selection/         ← Selección de variables
│   │   ├── 01_prueba_base.ipynb
│   │   ├── 02_prueba_igpr.ipynb
│   │   └── 03_resultado_final.ipynb   ← ✅ NOTEBOOK PRINCIPAL
│   │
│   ├── sgr/                       ← Análisis SGR general
│   │   └── 01_pipeline_sgr.ipynb
│   │
│   └── archive/                   ← Experimentos descartados
│
├── src/                           ← Funciones reutilizables en Python
│   ├── preprocessing.py
│   ├── variable_selection.py
│   └── clustering.py
│
├── outputs/
│   ├── rankings/                  ← Rankings de variables
│   └── plots/                     ← Gráficas generadas
│
├── docs/
│   ├── decisiones_metodologicas.md
│   └── mockup.html
│
└── README.md
```

---

## Cómo correr el notebook principal

### Local (VS Code / Jupyter)

```bash
git clone https://github.com/drambaut/Banco-de-ejecutores-del-SGR.git
cd Banco-de-Ejecutores-del-SGR
```

La celda de configuración detecta el entorno local y resuelve las rutas desde la raíz del repo.

---

## Archivos de datos esperados

| Archivo en `data/raw/` | Descripción |
|------------------------|-------------|
| `Base de proyectos _15_12_2025.xlsx` | Base principal de proyectos SGR |
| `Capacidades_institucionales.xlsx` | Capacidades de los ejecutores |
| `ResultadosIGPR.xlsx` | Resultados IGPR por entidad |

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

---

## Autores

Proyecto SAR — Dirección de Regalías, DNP  
Análisis: Daniel Felipe Rambaut Lemus
