"""
etl_normalizar_excels.py
=========================
Lee los 4 Excel fuente que hoy se manejan por separado y los normaliza en
UN SOLO Excel maestro (EXCEL_MAESTRO_ICS.xlsx) con 4 hojas limpias:

    Ejecutores        -> catálogo de entidades ejecutoras (1 fila por ejecutor)
    Proyectos         -> catálogo de proyectos (1 fila por BPIN)
    Periodos          -> avance mes a mes por proyecto (1 fila por BPIN+periodo)
    Reprogramaciones  -> reprogramaciones no permitidas (1 fila por BPIN)

ARCHIVOS FUENTE Y COLUMNAS QUE SE EXTRAEN DE CADA UNO
------------------------------------------------------
1) Balance_seguimiento_SGR.xlsx  (hoja "PROYECTOS APROBADOS ", con espacio al final)
   Catálogo maestro de proyectos. Header real está en la fila 8 de Excel
   (índice 7 en pandas) porque las primeras filas son título/notas.
   Columnas usadas:
     BPIN, ESTADO GENERAL, TOTAL PROYECTO, CÓDIGO EJECUTOR, ENTIDAD EJECUTORA,
     DEPARTAMENTO LOCALIZACIÓN DEL EJECUTOR, REGIÓN LOCALIZACIÓN DEL EJECUTOR,
     NIT ENTIDAD EJECUTORA, TIPO EJECUTOR, CAPACIDAD INSTITUCIONAL,
     FECHA INICIAL DE LA PROGRAMACIÓN, FECHA FINAL DE LA PROGRAMACIÓN

2) Curva_S_total_proyectos_en_Ejecucion_y_terminados_al_*.xlsx
   (hoja "CURVASTOTALPROYENEJECUCION_2601")
   Avance mes a mes (fuente principal de periodos).
   Columnas usadas:
     CODIGO_EJECUTOR, BPIN, PERIODO, PERIODO_FECHA, PV_VALOR_MES, EV_VALOR_MES,
     FECHA_CORTE

3) Curva_Sl_*.xlsx  (hoja "Curva S")
   Segundo corte de avance mes a mes, mismas columnas clave que el archivo
   anterior pero sin AC_/ACUMULADO ni Departamento/Región. Se UNE (union) con
   la Curva S total: si un BPIN+PERIODO aparece en ambos archivos, se
   conserva el de FECHA_CORTE más reciente.
   >>> Si en tu operación real estos dos archivos representan cosas distintas
   >>> (p.ej. dos fuentes que NO deberían mezclarse), dímelo y separamos la
   >>> lógica de unión.

4) Reprogramaciones_no_permitidas.xlsx  (hoja "BASE FINAL")
   Columnas usadas:
     BPIN, TOTAL REPROGRAMACIONES NO PERMITIDAS,
     TOTAL REPROGRAMACIONES PERMITIDAS AJUSTADA, TOTAL REPROGRAMACIONES REALIZADAS,
     TECHO DE MEDICIÓN

CLAVES DE CRUCE
----------------
- BPIN conecta: Proyectos <-> Periodos <-> Reprogramaciones
- CÓDIGO EJECUTOR (o NIT, si se prefiere) conecta: Proyectos <-> Ejecutores
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

ARCHIVO_BALANCE = "Balance_seguimiento_SGR.xlsx"
ARCHIVO_CURVA_S = "Curva_S_total_proyectos_en_Ejecucion_y_terminados_al_23062026.xlsx"
ARCHIVO_CURVA_SI = "Curva_Sl_23062026.xlsx"
ARCHIVO_REPROGRAMACIONES = "Reprogramaciones_no_permitidas.xlsx"

SALIDA_MAESTRO = "EXCEL_MAESTRO_ICS.xlsx"


# =============================================================================
# NORMALIZACIÓN DE ESTADOS (distintos archivos pueden usar distinta escritura)
# =============================================================================

def _normalizar_estado(valor: str) -> str:
    if pd.isna(valor):
        return "Desconocido"
    v = str(valor).strip().upper()
    if "TERMIN" in v:
        return "Terminado"
    if "EJECUC" in v:
        return "En Ejecución"
    return str(valor).strip().title()


# =============================================================================
# 1. EJECUTORES + PROYECTOS  <-  Balance_seguimiento_SGR.xlsx
# =============================================================================

def cargar_proyectos_y_ejecutores(path_balance: str):
    df = pd.read_excel(path_balance, sheet_name="PROYECTOS APROBADOS ", header=7)

    columnas_proyecto = {
        "BPIN": "bpin",
        "NOMBRE DEL PROYECTO": "nombre_proyecto",
        "SECTOR": "sector",
        "ESTADO GENERAL": "estado",
        "TOTAL PROYECTO": "valor_total_proyecto",
        "CÓDIGO EJECUTOR": "codigo_ejecutor",
        "FECHA INICIAL DE LA PROGRAMACIÓN": "fecha_inicial_programacion",
        "FECHA FINAL DE LA PROGRAMACIÓN": "fecha_final_programacion",
    }
    df_proyectos = df[list(columnas_proyecto.keys())].rename(columns=columnas_proyecto)
    df_proyectos["bpin"] = df_proyectos["bpin"].astype(str)
    df_proyectos["codigo_ejecutor"] = df_proyectos["codigo_ejecutor"].astype(str)
    df_proyectos["estado"] = df_proyectos["estado"].apply(_normalizar_estado)
    df_proyectos = df_proyectos.drop_duplicates(subset="bpin")

    columnas_ejecutor = {
        "CÓDIGO EJECUTOR": "codigo_ejecutor",
        "ENTIDAD EJECUTORA": "nombre_ejecutor",
        "NIT ENTIDAD EJECUTORA": "nit",
        "DEPARTAMENTO LOCALIZACIÓN DEL EJECUTOR": "departamento",
        "REGIÓN LOCALIZACIÓN DEL EJECUTOR": "region",
        "TIPO EJECUTOR": "tipo_ejecutor",
        "CAPACIDAD INSTITUCIONAL": "capacidad_institucional",
    }
    df_ejecutores = df[list(columnas_ejecutor.keys())].rename(columns=columnas_ejecutor)
    df_ejecutores["codigo_ejecutor"] = df_ejecutores["codigo_ejecutor"].astype(str)
    df_ejecutores = df_ejecutores.drop_duplicates(subset="codigo_ejecutor")

    return df_proyectos, df_ejecutores


# =============================================================================
# 2-3. PERIODOS  <-  Curva S total + Curva SI  (unión, sin duplicar bpin+periodo)
# =============================================================================

def _cargar_curva(path: str, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet)
    columnas = {
        "CODIGO_EJECUTOR": "codigo_ejecutor",
        "BPIN": "bpin",
        "PERIODO": "periodo",
        "PERIODO_FECHA": "periodo_fecha",
        "PV_VALOR_MES": "valor_programado",
        "EV_VALOR_MES": "valor_ejecutado",
        "FECHA_CORTE": "fecha_corte",
    }
    faltantes = [c for c in columnas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Columnas esperadas no encontradas en {path}: {faltantes}")
    df = df[list(columnas.keys())].rename(columns=columnas)
    df["bpin"] = df["bpin"].astype(str)
    df["codigo_ejecutor"] = df["codigo_ejecutor"].astype(str)
    return df


def cargar_periodos(path_curva_s: str, path_curva_si: str) -> pd.DataFrame:
    curva_s = _cargar_curva(path_curva_s, "CURVASTOTALPROYENEJECUCION_2601")
    curva_si = _cargar_curva(path_curva_si, "Curva S")

    unido = pd.concat([curva_s, curva_si], ignore_index=True)
    unido = unido.sort_values("fecha_corte", ascending=False)
    # Si el mismo BPIN+periodo aparece en ambas fuentes, se conserva el corte más reciente
    unido = unido.drop_duplicates(subset=["bpin", "periodo"], keep="first")
    return unido.sort_values(["codigo_ejecutor", "bpin", "periodo"]).reset_index(drop=True)


# =============================================================================
# 4. REPROGRAMACIONES  <-  Reprogramaciones_no_permitidas.xlsx
# =============================================================================

def cargar_reprogramaciones(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="BASE FINAL")
    columnas = {
        "BPIN": "bpin",
        "TOTAL REPROGRAMACIONES PERMITIDAS AJUSTADA": "reprogramaciones_permitidas",
        "TOTAL REPROGRAMACIONES REALIZADAS": "reprogramaciones_realizadas",
        "TOTAL REPROGRAMACIONES NO PERMITIDAS": "reprogramaciones_no_permitidas",
        "TECHO DE MEDICIÓN": "techo_medicion",
    }
    df = df[list(columnas.keys())].rename(columns=columnas)
    df["bpin"] = df["bpin"].astype(str)
    return df


# =============================================================================
# PIPELINE DE NORMALIZACIÓN COMPLETO -> EXCEL MAESTRO
# =============================================================================

def construir_excel_maestro(
    path_balance: str = ARCHIVO_BALANCE,
    path_curva_s: str = ARCHIVO_CURVA_S,
    path_curva_si: str = ARCHIVO_CURVA_SI,
    path_reprogramaciones: str = ARCHIVO_REPROGRAMACIONES,
    salida: str = SALIDA_MAESTRO,
) -> dict[str, pd.DataFrame]:

    df_proyectos, df_ejecutores = cargar_proyectos_y_ejecutores(path_balance)
    df_periodos = cargar_periodos(path_curva_s, path_curva_si)
    df_reprogramaciones = cargar_reprogramaciones(path_reprogramaciones)

    tablas = {
        "Ejecutores": df_ejecutores,
        "Proyectos": df_proyectos,
        "Periodos": df_periodos,
        "Reprogramaciones": df_reprogramaciones,
    }

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        for nombre_hoja, tabla in tablas.items():
            tabla.to_excel(writer, sheet_name=nombre_hoja, index=False)

    return tablas


if __name__ == "__main__":
    tablas = construir_excel_maestro()
    print(f"Excel maestro generado en: {Path(SALIDA_MAESTRO).resolve()}\n")
    for nombre, tabla in tablas.items():
        print(f"--- {nombre} ({len(tabla)} filas) ---")
        print(tabla.head(5).to_string())
        print()
