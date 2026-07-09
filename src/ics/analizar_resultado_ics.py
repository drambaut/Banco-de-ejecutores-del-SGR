"""
analizar_resultado_ics.py
============================
Analiza cualquier archivo de resultados del ICS (salida de
indicador_cumplimiento_historico.py / _v2.py) y genera:

  - KPIs en consola (conteos por nivel de riesgo, estadísticas descriptivas,
    casos "límite" cerca de un umbral, top mejores/peores, correlaciones)
  - Gráficas (PNG) guardadas en una carpeta de salida
  - Un resumen en texto (.txt) con los mismos KPIs, para archivar

USO
----
  # Analizar una sola hoja
  python analizar_resultado_ics.py ruta/archivo.xlsx --hoja Sheet1

  # Comparar varias hojas del mismo archivo (ej. distintos tau)
  python analizar_resultado_ics.py ruta/archivo.xlsx --comparar Resultado_tau_5pct Resultado_tau_10pct Resultado_tau_15pct

  # Si no se indica --hoja ni --comparar, se analiza la primera hoja del archivo

El script detecta automáticamente qué columnas están presentes (no todas
las hojas de resultados tienen exactamente las mismas: p.ej. una hoja de
resumen no tiene 'ics', y no pasa nada, esa parte del análisis simplemente
se omite) — está pensado para las salidas estándar de:
    - indicador_cumplimiento_historico.py / _v2.py  (columnas: codigo_ejecutor,
      capacidad_institucional, tbc, n_proyectos, ve, vref, fc,
      reprogramaciones_no_permitidas, pen, ics, ics_norm, puntaje_riesgo,
      nivel_riesgo)
    - analisis_sensibilidad_tau.py (mismas columnas + 'tau_usado')
"""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd

plt.style.use("seaborn-v0_8-whitegrid")
plt.rcParams.update({
    "figure.figsize": (10, 6),
    "figure.dpi": 150,
    "font.size": 11,
    "axes.titlesize": 14,
    "axes.titleweight": "bold",
    "axes.labelsize": 11,
    "xtick.labelsize": 10,
    "ytick.labelsize": 10,
    "legend.fontsize": 10,
})

COLOR_ALTO = "#C44E52"
COLOR_MEDIO = "#DD8452"
COLOR_BAJO = "#55A868"
COLORES_RIESGO = {"Riesgo Alto": COLOR_ALTO, "Riesgo Medio": COLOR_MEDIO, "Riesgo Bajo": COLOR_BAJO}
PALETTE = ["#4C72B0", "#DD8452", "#55A868", "#C44E52", "#8172B3", "#937860"]

UMBRAL_RIESGO_BAJO = 0.70
UMBRAL_RIESGO_MEDIO = 0.40
ORDEN_RIESGO = ["Riesgo Bajo", "Riesgo Medio", "Riesgo Alto"]


# =============================================================================
# CARGA
# =============================================================================

def cargar_hoja(path: str, hoja: str | None = None) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=hoja if hoja else 0)
    if "codigo_ejecutor" in df.columns:
        df["codigo_ejecutor"] = df["codigo_ejecutor"].astype(str)
    return df


def tiene_columnas_ics(df: pd.DataFrame) -> bool:
    return {"ics", "puntaje_riesgo", "nivel_riesgo"}.issubset(df.columns)


# =============================================================================
# KPIs EN TEXTO
# =============================================================================

def resumen_kpis(df: pd.DataFrame, titulo: str = "") -> str:
    lineas = []
    lineas.append(f"\n{'='*70}\nKPIs — {titulo}\n{'='*70}")

    total = len(df)
    lineas.append(f"\nTotal de ejecutores analizados: {total}")

    if "nivel_riesgo" in df.columns:
        conteo = df["nivel_riesgo"].value_counts().reindex(ORDEN_RIESGO).fillna(0).astype(int)
        lineas.append("\nDistribución por nivel de riesgo:")
        for nivel in ORDEN_RIESGO:
            n = conteo.get(nivel, 0)
            pct = n / total * 100 if total else 0
            lineas.append(f"  {nivel:<14}: {n:>5} ejecutores  ({pct:5.1f}%)")

    for col, etiqueta in [
        ("tbc", "Tasa Bruta de Cumplimiento (TBC)"),
        ("fc", "Factor de Carga (FC)"),
        ("pen", "Penalización por reprogramaciones (Pen)"),
        ("ics", "Índice de Cumplimiento Histórico (ICS)"),
        ("puntaje_riesgo", "Puntaje de riesgo (0-100)"),
        ("n_proyectos", "N° de proyectos"),
        ("reprogramaciones_no_permitidas", "Reprogramaciones no permitidas"),
    ]:
        if col in df.columns:
            s = df[col].dropna()
            lineas.append(
                f"\n{etiqueta}: media={s.mean():,.3f} | mediana={s.median():,.3f} | "
                f"p25={s.quantile(.25):,.3f} | p75={s.quantile(.75):,.3f} | "
                f"min={s.min():,.3f} | max={s.max():,.3f}"
            )

    if "ics_norm" in df.columns:
        margen = 0.05
        cerca_alto_medio = df[
            (df["ics_norm"] >= UMBRAL_RIESGO_MEDIO - margen) & (df["ics_norm"] < UMBRAL_RIESGO_MEDIO + margen)
        ]
        cerca_medio_bajo = df[
            (df["ics_norm"] >= UMBRAL_RIESGO_BAJO - margen) & (df["ics_norm"] < UMBRAL_RIESGO_BAJO + margen)
        ]
        lineas.append(
            f"\nCasos 'límite' (a ±{margen:.0%} de un umbral, sensibles a recalibración):"
        )
        lineas.append(f"  Cerca del umbral Alto/Medio (ics_norm≈{UMBRAL_RIESGO_MEDIO}): {len(cerca_alto_medio)} ejecutores")
        lineas.append(f"  Cerca del umbral Medio/Bajo (ics_norm≈{UMBRAL_RIESGO_BAJO}): {len(cerca_medio_bajo)} ejecutores")

    if "capacidad_institucional" in df.columns and "puntaje_riesgo" in df.columns:
        lineas.append("\nPuntaje de riesgo promedio por grupo de capacidad institucional:")
        agg = df.groupby("capacidad_institucional")["puntaje_riesgo"].agg(["count", "mean", "median"])
        for grupo, fila in agg.iterrows():
            lineas.append(
                f"  Grupo {grupo}: n={int(fila['count']):>4} | promedio={fila['mean']:6.1f} | mediana={fila['median']:6.1f}"
            )

    if {"tbc", "puntaje_riesgo"}.issubset(df.columns):
        lineas.append("\nCorrelación de cada componente con el puntaje de riesgo (Pearson):")
        for col in ["tbc", "fc", "pen", "n_proyectos", "reprogramaciones_no_permitidas"]:
            if col in df.columns:
                corr = df[[col, "puntaje_riesgo"]].dropna().corr().iloc[0, 1]
                lineas.append(f"  {col:<32}: {corr:+.3f}")

    if "codigo_ejecutor" in df.columns and "puntaje_riesgo" in df.columns:
        peores = df.nlargest(10, "puntaje_riesgo")[["codigo_ejecutor", "puntaje_riesgo", "nivel_riesgo"]]
        mejores = df.nsmallest(10, "puntaje_riesgo")[["codigo_ejecutor", "puntaje_riesgo", "nivel_riesgo"]]
        lineas.append("\nTop 10 mayor riesgo (codigo_ejecutor - puntaje):")
        for _, fila in peores.iterrows():
            lineas.append(f"  {fila['codigo_ejecutor']:<15} {fila['puntaje_riesgo']:6.1f}  {fila['nivel_riesgo']}")
        lineas.append("\nTop 10 menor riesgo (codigo_ejecutor - puntaje):")
        for _, fila in mejores.iterrows():
            lineas.append(f"  {fila['codigo_ejecutor']:<15} {fila['puntaje_riesgo']:6.1f}  {fila['nivel_riesgo']}")

    return "\n".join(lineas)


# =============================================================================
# GRÁFICAS (una sola hoja)
# =============================================================================

def graficar_hoja(df: pd.DataFrame, carpeta: Path, prefijo: str = ""):
    carpeta.mkdir(parents=True, exist_ok=True)

    # 1) Conteo por nivel de riesgo
    if "nivel_riesgo" in df.columns:
        conteo = df["nivel_riesgo"].value_counts().reindex(ORDEN_RIESGO).fillna(0)
        fig, ax = plt.subplots()
        colores = [COLORES_RIESGO[n] for n in conteo.index]
        barras = ax.bar(conteo.index, conteo.values, color=colores)
        for barra in barras:
            h = barra.get_height()
            ax.text(barra.get_x() + barra.get_width() / 2, h, f"{int(h)}\n({h/len(df)*100:.1f}%)",
                    ha="center", va="bottom", fontsize=10)
        ax.set_title(f"Ejecutores por nivel de riesgo{(' — ' + prefijo) if prefijo else ''}")
        ax.set_ylabel("N° de ejecutores")
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.savefig(carpeta / f"{prefijo}01_conteo_nivel_riesgo.png", bbox_inches="tight")
        plt.close(fig)

    # 2) Histograma de puntaje de riesgo con umbrales
    if "puntaje_riesgo" in df.columns:
        fig, ax = plt.subplots()
        ax.hist(df["puntaje_riesgo"], bins=30, color=PALETTE[0], edgecolor="white", alpha=0.85)
        ax.axvline((1 - UMBRAL_RIESGO_BAJO) * 100, color=COLOR_BAJO, linestyle="--", linewidth=2,
                   label=f"Umbral Bajo/Medio ({(1-UMBRAL_RIESGO_BAJO)*100:.0f} pts)")
        ax.axvline((1 - UMBRAL_RIESGO_MEDIO) * 100, color=COLOR_ALTO, linestyle="--", linewidth=2,
                   label=f"Umbral Medio/Alto ({(1-UMBRAL_RIESGO_MEDIO)*100:.0f} pts)")
        ax.set_title(f"Distribución del puntaje de riesgo{(' — ' + prefijo) if prefijo else ''}")
        ax.set_xlabel("Puntaje de riesgo (0=mejor, 100=peor)")
        ax.set_ylabel("N° de ejecutores")
        ax.legend()
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.savefig(carpeta / f"{prefijo}02_histograma_puntaje_riesgo.png", bbox_inches="tight")
        plt.close(fig)

    # 3) Histograma del ICS crudo (media y mediana)
    if "ics" in df.columns:
        fig, ax = plt.subplots()
        ax.hist(df["ics"], bins=40, color=PALETTE[1], edgecolor="white", alpha=0.85)
        ax.axvline(df["ics"].mean(), color="red", linestyle="--", linewidth=1.5, label=f"Media: {df['ics'].mean():.2f}")
        ax.axvline(df["ics"].median(), color="green", linestyle="--", linewidth=1.5, label=f"Mediana: {df['ics'].median():.2f}")
        ax.set_title(f"Distribución del ICS (crudo){(' — ' + prefijo) if prefijo else ''}")
        ax.set_xlabel("ICS")
        ax.set_ylabel("N° de ejecutores")
        ax.legend()
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.savefig(carpeta / f"{prefijo}03_histograma_ics.png", bbox_inches="tight")
        plt.close(fig)

    # 4) Boxplot de TBC por grupo de capacidad institucional
    if {"tbc", "capacidad_institucional"}.issubset(df.columns):
        grupos = sorted(df["capacidad_institucional"].dropna().unique())
        datos = [df.loc[df["capacidad_institucional"] == g, "tbc"].dropna() for g in grupos]
        fig, ax = plt.subplots()
        bp = ax.boxplot(datos, tick_labels=[str(g) for g in grupos], patch_artist=True)
        for patch, color in zip(bp["boxes"], PALETTE * 3):
            patch.set_facecolor(color)
            patch.set_alpha(0.6)
        ax.set_title(f"TBC por grupo de capacidad institucional{(' — ' + prefijo) if prefijo else ''}")
        ax.set_xlabel("Capacidad institucional")
        ax.set_ylabel("Tasa Bruta de Cumplimiento (TBC)")
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.savefig(carpeta / f"{prefijo}04_boxplot_tbc_por_grupo.png", bbox_inches="tight")
        plt.close(fig)

    # 5) Puntaje de riesgo promedio por grupo
    if {"puntaje_riesgo", "capacidad_institucional"}.issubset(df.columns):
        agg = df.groupby("capacidad_institucional")["puntaje_riesgo"].mean().sort_index()
        fig, ax = plt.subplots()
        ax.bar(agg.index.astype(str), agg.values, color=PALETTE[2])
        ax.set_title(f"Puntaje de riesgo promedio por grupo{(' — ' + prefijo) if prefijo else ''}")
        ax.set_xlabel("Capacidad institucional")
        ax.set_ylabel("Puntaje de riesgo promedio")
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.savefig(carpeta / f"{prefijo}05_puntaje_promedio_por_grupo.png", bbox_inches="tight")
        plt.close(fig)

    # 6) Dispersión TBC vs Puntaje de riesgo, coloreado por nivel
    if {"tbc", "puntaje_riesgo", "nivel_riesgo"}.issubset(df.columns):
        fig, ax = plt.subplots()
        for nivel in ORDEN_RIESGO:
            subset = df[df["nivel_riesgo"] == nivel]
            ax.scatter(subset["tbc"], subset["puntaje_riesgo"], s=18, alpha=0.6,
                       color=COLORES_RIESGO[nivel], label=nivel)
        ax.set_title(f"TBC vs. Puntaje de riesgo{(' — ' + prefijo) if prefijo else ''}")
        ax.set_xlabel("Tasa Bruta de Cumplimiento (TBC)")
        ax.set_ylabel("Puntaje de riesgo")
        ax.legend()
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.savefig(carpeta / f"{prefijo}06_dispersion_tbc_vs_riesgo.png", bbox_inches="tight")
        plt.close(fig)

    # 7) Dispersión FC vs Puntaje de riesgo
    if {"fc", "puntaje_riesgo", "nivel_riesgo"}.issubset(df.columns):
        fig, ax = plt.subplots()
        for nivel in ORDEN_RIESGO:
            subset = df[df["nivel_riesgo"] == nivel]
            ax.scatter(subset["fc"], subset["puntaje_riesgo"], s=18, alpha=0.6,
                       color=COLORES_RIESGO[nivel], label=nivel)
        ax.set_title(f"Factor de Carga (FC) vs. Puntaje de riesgo{(' — ' + prefijo) if prefijo else ''}")
        ax.set_xlabel("Factor de Carga (FC)")
        ax.set_ylabel("Puntaje de riesgo")
        ax.legend()
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.savefig(carpeta / f"{prefijo}07_dispersion_fc_vs_riesgo.png", bbox_inches="tight")
        plt.close(fig)


# =============================================================================
# COMPARACIÓN ENTRE VARIAS HOJAS (ej. distintos tau)
# =============================================================================

def graficar_comparacion(dfs: dict[str, pd.DataFrame], carpeta: Path):
    carpeta.mkdir(parents=True, exist_ok=True)

    # Barras agrupadas: distribución de riesgo por hoja/escenario
    resumen = {}
    for nombre, df in dfs.items():
        if "nivel_riesgo" not in df.columns:
            continue
        conteo = df["nivel_riesgo"].value_counts().reindex(ORDEN_RIESGO).fillna(0)
        resumen[nombre] = conteo / len(df) * 100

    if resumen:
        tabla = pd.DataFrame(resumen).T[ORDEN_RIESGO]
        fig, ax = plt.subplots(figsize=(11, 6))
        x = np.arange(len(tabla.index))
        ancho = 0.25
        for i, nivel in enumerate(ORDEN_RIESGO):
            ax.bar(x + (i - 1) * ancho, tabla[nivel], width=ancho, label=nivel, color=COLORES_RIESGO[nivel])
        ax.set_xticks(x)
        ax.set_xticklabels(tabla.index, rotation=15)
        ax.set_ylabel("% de ejecutores")
        ax.set_title("Comparación de distribución de riesgo entre escenarios")
        ax.legend()
        ax.spines[["top", "right"]].set_visible(False)
        plt.tight_layout()
        plt.savefig(carpeta / "00_comparacion_escenarios.png", bbox_inches="tight")
        plt.close(fig)


# =============================================================================
# CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Analiza resultados del ICS: KPIs y gráficas.")
    parser.add_argument("excel", help="Ruta al archivo Excel de resultados")
    parser.add_argument("--hoja", default=None, help="Nombre de la hoja a analizar (por defecto, la primera)")
    parser.add_argument("--comparar", nargs="+", default=None,
                         help="Lista de hojas a comparar entre sí (ej. distintos escenarios de tau)")
    parser.add_argument("--salida", default="analisis_ics_output", help="Carpeta donde guardar gráficas y resumen")
    args = parser.parse_args()

    carpeta_salida = Path(args.salida)
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    texto_resumen = []

    if args.comparar:
        dfs = {hoja: cargar_hoja(args.excel, hoja) for hoja in args.comparar}
        for nombre, df in dfs.items():
            if tiene_columnas_ics(df):
                texto_resumen.append(resumen_kpis(df, titulo=nombre))
                graficar_hoja(df, carpeta_salida / "por_escenario", prefijo=f"{nombre}_")
        graficar_comparacion(dfs, carpeta_salida)
    else:
        df = cargar_hoja(args.excel, args.hoja)
        if not tiene_columnas_ics(df):
            print(
                "Aviso: esta hoja no tiene las columnas estándar del ICS "
                "(ics, puntaje_riesgo, nivel_riesgo). Se muestran solo las "
                "columnas disponibles como referencia:\n", df.columns.tolist()
            )
        else:
            texto_resumen.append(resumen_kpis(df, titulo=args.hoja or "hoja 1"))
            graficar_hoja(df, carpeta_salida)

    resumen_final = "\n".join(texto_resumen)
    print(resumen_final)
    (carpeta_salida / "resumen_kpis.txt").write_text(resumen_final, encoding="utf-8")
    print(f"\nGráficas y resumen guardados en: {carpeta_salida.resolve()}")


if __name__ == "__main__":
    main()
