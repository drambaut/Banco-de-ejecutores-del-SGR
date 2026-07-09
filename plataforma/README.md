# SAR — Sistema de Análisis de Riesgo (SGR)

Plataforma web que muestra el perfil de riesgo de ejecutores del SGR y un
tablero de análisis descriptivo, a partir de los cálculos del Índice de
Cumplimiento Histórico (metodología v2: tau=10%, normalización por percentil).

## Estructura del repo

```
sar-plataforma/
├── app/
│   ├── main.py              # Backend Flask (API + servir frontend)
│   └── static/
│       ├── index.html       # Frontend (adaptado del mockup)
│       ├── styles.css       # Estilos (extraídos del mockup, sin cambios)
│       └── app.js           # Lógica de fetch + render dinámico
├── data/
│   ├── EXCEL_MAESTRO_ICS.xlsx        # Ejecutores + Proyectos (salida de etl_normalizar_excels.py)
│   └── Resultado_ICS_v2_corregido.xlsx  # Salida de indicador_cumplimiento_historico_v2.py
├── build_db.py               # Construye data/sar.db (SQLite) a partir de los 2 Excel de arriba
├── requirements.txt
├── render.yaml                # Config de despliegue en Render
└── README.md
```

## Cómo actualizar los datos

Cada vez que tengas un nuevo corte:

1. Corre `etl_normalizar_excels.py` sobre los 4 Excel fuente → genera `EXCEL_MAESTRO_ICS.xlsx`
2. Corre `indicador_cumplimiento_historico_v2.py` → genera `Resultado_ICS_v2_corregido.xlsx`
3. Copia ambos archivos a `data/`
4. Corre `python build_db.py` → regenera `data/sar.db`
5. Sube los cambios a GitHub (Render redepliega automáticamente)

## Correr localmente

```bash
pip install -r requirements.txt
python build_db.py
python app/main.py
```

Abre http://localhost:8000

## Desplegar en Render

1. Sube este repo a GitHub (ver sección siguiente)
2. En Render: **New > Web Service**, conecta el repo
3. Render detecta `render.yaml` automáticamente (Build Command y Start Command ya quedan configurados)
4. Cada `git push` a `main` redespliega solo

**Importante:** el archivo `data/sar.db` se regenera en cada build (`build_db.py`
corre como parte del build command), así que basta con mantener actualizados
los 2 Excel en `data/` y hacer commit — no hace falta subir el `.db` directamente.

## Subir a GitHub (primera vez)

```bash
cd sar-plataforma
git init
git add .
git commit -m "Plataforma SAR - versión inicial"
git branch -M main
git remote add origin https://github.com/TU-USUARIO/sar-plataforma.git
git push -u origin main
```

## Limitaciones conocidas de esta versión (ver plan de implementación)

- **Capacidad Financiera e Institucional**: no integradas (requieren SECOP,
  patrimonio, historial de entes de control). Solo Administrativa se aproxima
  a partir del TBC, marcado explícitamente como provisional en la UI.
- **Serie histórica por vigencia**: solo existe un corte de datos; la
  plataforma lo indica en vez de inventar una tendencia.
- **Sector / nombre de proyecto**: pendientes hasta re-correr
  `etl_normalizar_excels.py` sobre el archivo completo de Balance (ya
  actualizado para extraerlos, ver el script).
- **4 bandas de riesgo (Bajo/Medio/Alto/Crítico)**: la metodología v2 define
  3 niveles; para calzar con el gauge de 4 colores del mockup se subdividió
  "Alto" usando cortes en 30/60/85 puntos. Ajustar en `nivel_4_bandas()` en
  `app/main.py` si el DNP define oficialmente otros cortes.
