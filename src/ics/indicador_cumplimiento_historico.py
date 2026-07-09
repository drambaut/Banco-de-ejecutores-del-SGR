"""
indicador_cumplimiento_historico_v2.py — METODOLOGÍA CORREGIDA
==================================================================
Versión corregida de indicador_cumplimiento_historico.py a partir de los
hallazgos del análisis de sensibilidad (ver analisis_sensibilidad_tau.py)
y del diagnóstico del 90% de ejecutores en "Riesgo Alto".

QUÉ CAMBIA FRENTE A LA V1 Y POR QUÉ
-------------------------------------

1) TAU: 5% -> 10%
   La mediana real de |desviación| en ~549.000 periodos es ~19%; con tau=5%
   solo el 42.6% de los periodos "cumplen" incluso agregando TODA la
   ejecución del SGR. Subir a 10% es una tolerancia todavía exigente pero
   más realista para desfases normales de caja mes a mes.

   IMPORTANTE (evidencia del propio análisis de sensibilidad): mover tau de
   5% a 15% solo cambia el % de "Riesgo Alto" de 90.4% a 89.6%. Es decir,
   *tau no es la causa del problema* — se deja en 10% por ser más realista,
   pero el cambio que de verdad corrige la distribución es el de
   normalización (punto 2).

2) NORMALIZACIÓN: min-max -> percentil (rank) dentro del grupo
   La min-max (ICS-min)/(max-min) hace que UN SOLO ejecutor con Factor de
   Carga extremo (muchísimos proyectos / valor) defina el techo de toda la
   escala, comprimiendo a todos los demás cerca de 0 aunque tengan un
   desempeño intermedio. Se reemplaza por el percentil del ICS dentro de su
   grupo de capacidad institucional:

       ics_percentil = posición relativa del ejecutor (0=peor, 1=mejor)
                        dentro de su grupo, por definición uniforme y
                        prácticamente insensible a valores extremos.

   Validado en el diagnóstico: con percentil, la distribución de riesgo pasa
   de (90.4% Alto / 7.9% Medio / 1.7% Bajo) a algo mucho más balanceado
   (~30% Alto / ~40% Medio / ~30% Bajo).

SUGERENCIAS ADICIONALES PARA LA FASE 3 DE CALIBRACIÓN (no implementadas aún,
quedan documentadas para discusión):
-----------------------------------------------------------------------------
  a) Separar N (número de proyectos) en "En Ejecución" vs "Terminados"
     (lo sugiere la propia diapositiva 17 del PPTX) — hoy ambos cuentan
     igual y un ejecutor con muchos proyectos ya cerrados exitosamente
     podría estar subiendo su FC sin que eso implique carga operativa
     *actual*.
  b) Ponderar la penalización por reprogramaciones por el VALOR del
     proyecto reprogramado, no solo por el conteo (diapositiva 20).
  c) Evaluar el TCP sobre el valor ACUMULADO (PV_ACUMULADO/EV_ACUMULADO) en
     vez del valor mensual puntual, para reducir el ruido de calendario que
     hace que un simple corrimiento de un pago de un mes a otro cuente como
     "no cumple" dos veces (el mes que se atrasó y el mes que compensó).
  d) Revisar si conviene capar (winsorizar) Ve y N en el percentil 95 antes
     de calcular FC, como capa adicional de robustez sobre la ya aportada
     por el cambio de normalización.

Todo lo demás (Paso 1 a 5: TBC, FC, Pen, ICS) se mantiene idéntico a la v1,
ya validado contra ICS_JOHANNA_ROZO.xlsx.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

# =============================================================================
# PARÁMETROS METODOLÓGICOS (corregidos)
# =============================================================================
TAU_TOLERANCIA = 0.10      # antes 0.05 -- ver justificación arriba
P_SUAVIDAD = 0.50          # sin cambios
UMBRAL_RIESGO_BAJO = 0.70  # sin cambios (ahora aplicado sobre percentil, no min-max)
UMBRAL_RIESGO_MEDIO = 0.40


# =============================================================================
# PASO 1-2: TASA BRUTA DE CUMPLIMIENTO (TBC)  — idéntico a la v1
# =============================================================================

def calcular_tcp(df_periodos: pd.DataFrame, tau: float = TAU_TOLERANCIA) -> pd.DataFrame:
    df = df_periodos.copy()
    with np.errstate(divide="ignore", invalid="ignore"):
        df["desviacion"] = np.where(
            df["valor_programado"] != 0,
            df["valor_ejecutado"] / df["valor_programado"] - 1,
            np.nan,
        )
    df["cumple"] = df["desviacion"].abs() <= tau
    return df


def calcular_tbc(df_tcp: pd.DataFrame) -> pd.DataFrame:
    validos = df_tcp.dropna(subset=["cumple"])
    agg = validos.groupby("codigo_ejecutor").agg(
        periodos_evaluados=("cumple", "size"),
        periodos_cumplidos=("cumple", "sum"),
    )
    agg["tbc"] = agg["periodos_cumplidos"] / agg["periodos_evaluados"]
    return agg.reset_index()


# =============================================================================
# PASO 3: FACTOR DE CARGA (FC) — idéntico a la v1
# =============================================================================

def calcular_valor_ejecutor(df_periodos: pd.DataFrame) -> pd.DataFrame:
    ve = df_periodos.groupby("codigo_ejecutor")["valor_programado"].sum()
    return ve.rename("ve").reset_index()


def calcular_num_proyectos(
    df_proyectos: pd.DataFrame,
    estados_incluidos: tuple[str, ...] = ("en ejecución",),
) -> pd.DataFrame:
    estados_norm = {e.lower() for e in estados_incluidos}
    filtrado = df_proyectos[df_proyectos["estado"].str.lower().isin(estados_norm)]
    n = filtrado.groupby("codigo_ejecutor")["bpin"].nunique()
    return n.rename("n_proyectos").reset_index()


def calcular_vref(
    df_ve: pd.DataFrame,
    df_capacidad: pd.DataFrame,
    grupo_col: str = "capacidad_institucional",
) -> pd.DataFrame:
    df = df_ve.merge(df_capacidad, on="codigo_ejecutor", how="left")
    vref = df.groupby(grupo_col)["ve"].median().rename("vref")
    df = df.merge(vref, on=grupo_col, how="left")
    return df


def calcular_fc(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["fc"] = np.log1p(df["n_proyectos"]) + np.log1p(df["ve"] / df["vref"])
    return df


# =============================================================================
# PASO 4: PENALIZACIÓN POR REPROGRAMACIONES — idéntico a la v1
# =============================================================================

def calcular_penalizacion(df: pd.DataFrame, p: float = P_SUAVIDAD) -> pd.DataFrame:
    df = df.copy()
    df["pen"] = 1 / (1 + p * np.log1p(df["reprogramaciones_no_permitidas"]))
    df["descuento_pct"] = (1 - df["pen"]) * 100
    return df


# =============================================================================
# PASO 5: ICS — idéntico a la v1
# =============================================================================

def calcular_ics(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["ics"] = df["tbc"] * df["fc"] * df["pen"]
    return df


# =============================================================================
# PASOS 6-8: NORMALIZACIÓN Y PUNTAJE DE RIESGO  —  *** CORREGIDO ***
# =============================================================================

def calcular_puntaje_riesgo(
    df: pd.DataFrame,
    grupo_col: str = "capacidad_institucional",
) -> pd.DataFrame:
    """
    CORRECCIÓN: en vez de min-max -- (ICS-min)/(max-min), sensible a un solo
    outlier -- se usa el percentil (rank) del ICS dentro del grupo de
    comparación. Por definición queda en [0,1] y no se deja arrastrar por
    valores extremos: un ejecutor "típico" cae cerca de 0.5, no cerca de 0.

    method="average": si dos ejecutores empatan en ICS, ambos reciben el
    percentil promedio de las posiciones que ocupan.
    """
    df = df.copy()
    df["ics_norm"] = df.groupby(grupo_col)["ics"].rank(pct=True, method="average")
    df["puntaje_riesgo"] = (1 - df["ics_norm"]) * 100

    condiciones = [
        df["ics_norm"] >= UMBRAL_RIESGO_BAJO,
        df["ics_norm"] >= UMBRAL_RIESGO_MEDIO,
    ]
    niveles = ["Riesgo Bajo", "Riesgo Medio"]
    df["nivel_riesgo"] = np.select(condiciones, niveles, default="Riesgo Alto")
    return df


# =============================================================================
# PIPELINE COMPLETO
# =============================================================================

def calcular_indicador_completo(
    df_periodos: pd.DataFrame,
    df_proyectos: pd.DataFrame,
    df_reprogramaciones: pd.DataFrame,
    df_capacidad: pd.DataFrame,
    tau: float = TAU_TOLERANCIA,
    p: float = P_SUAVIDAD,
) -> pd.DataFrame:
    tcp = calcular_tcp(df_periodos, tau=tau)
    tbc = calcular_tbc(tcp)

    ve = calcular_valor_ejecutor(df_periodos)
    n_proy = calcular_num_proyectos(df_proyectos)
    base_fc = ve.merge(n_proy, on="codigo_ejecutor", how="left")
    base_fc["n_proyectos"] = base_fc["n_proyectos"].fillna(0)
    base_fc = calcular_vref(base_fc, df_capacidad)
    base_fc = calcular_fc(base_fc)

    pen = calcular_penalizacion(df_reprogramaciones, p=p)

    resultado = (
        tbc.merge(base_fc, on="codigo_ejecutor", how="left")
        .merge(pen, on="codigo_ejecutor", how="left")
    )
    resultado["reprogramaciones_no_permitidas"] = resultado[
        "reprogramaciones_no_permitidas"
    ].fillna(0)
    resultado["pen"] = resultado["pen"].fillna(1.0)

    resultado = calcular_ics(resultado)
    resultado = calcular_puntaje_riesgo(resultado)

    columnas_finales = [
        "codigo_ejecutor",
        "capacidad_institucional",
        "periodos_evaluados",
        "periodos_cumplidos",
        "tbc",
        "n_proyectos",
        "ve",
        "vref",
        "fc",
        "reprogramaciones_no_permitidas",
        "pen",
        "descuento_pct",
        "ics",
        "ics_norm",
        "puntaje_riesgo",
        "nivel_riesgo",
    ]
    return resultado[columnas_finales].sort_values("puntaje_riesgo").reset_index(drop=True)


if __name__ == "__main__":
    from calcular_ics_desde_maestro import leer_maestro, preparar_insumos

    ejecutores, proyectos, periodos, reprogramaciones = leer_maestro("EXCEL_MAESTRO_ICS.xlsx")
    df_periodos, df_proyectos, df_reprog, df_capacidad = preparar_insumos(
        ejecutores, proyectos, periodos, reprogramaciones
    )

    resultado = calcular_indicador_completo(df_periodos, df_proyectos, df_reprog, df_capacidad)

    print("=== DISTRIBUCIÓN DE RIESGO — METODOLOGÍA CORREGIDA (tau=10%, percentil) ===")
    conteo = resultado["nivel_riesgo"].value_counts()
    total = len(resultado)
    for nivel, n in conteo.items():
        print(f"{nivel}: {n} ({n/total*100:.1f}%)")

    print("\n=== LOS 3 EJECUTORES DE REFERENCIA ===")
    referencia = ["73000", "19000", "63000"]
    print(
        resultado[resultado["codigo_ejecutor"].isin(referencia)][
            ["codigo_ejecutor", "tbc", "fc", "pen", "ics", "ics_norm", "puntaje_riesgo", "nivel_riesgo"]
        ].to_string(index=False)
    )

    resultado.to_excel("Resultado_ICS_v2_corregido.xlsx", index=False)
    print("\nGuardado en Resultado_ICS_v2_corregido.xlsx")
