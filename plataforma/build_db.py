"""
build_db.py — Construye la base de datos SQLite que consume la plataforma SAR.

Fuentes:
    data/EXCEL_MAESTRO_ICS.xlsx        -> hojas Ejecutores, Proyectos
    data/Resultado_ICS_v2_corregido.xlsx -> resultado del ICS (metodología v2)

Salida:
    data/sar.db  (SQLite)

Correr cada vez que haya datos nuevos:
    python build_db.py
"""

import sqlite3
from pathlib import Path

import pandas as pd

DIR_DATOS = Path(__file__).parent / "data"
MAESTRO = DIR_DATOS / "EXCEL_MAESTRO_ICS.xlsx"
RESULTADO = DIR_DATOS / "Resultados_m1.xlsx" # "Resultados_m2.xlsx"
DB_PATH = DIR_DATOS / "sar.db"


def construir_bd():
    ejecutores = pd.read_excel(MAESTRO, sheet_name="Ejecutores", dtype={"codigo_ejecutor": str})
    proyectos = pd.read_excel(MAESTRO, sheet_name="Proyectos", dtype={"bpin": str, "codigo_ejecutor": str})
    resultado = pd.read_excel(RESULTADO, dtype={"codigo_ejecutor": str})

    # columnas opcionales que pueden no existir todavía en el maestro (se agregan
    # cuando se re-corra etl_normalizar_excels.py sobre el Balance completo)
    for col in ["sector", "nombre_proyecto"]:
        if col not in proyectos.columns:
            proyectos[col] = None

    # BPIN representativo de cada ejecutor: el de mayor valor total (para el
    # click en "otras entidades sugeridas" -> mantener la búsqueda por BPIN)
    proyectos_validos = proyectos.dropna(subset=["valor_total_proyecto"])
    idx_max = proyectos_validos.groupby("codigo_ejecutor")["valor_total_proyecto"].idxmax()
    bpin_representativo = proyectos_validos.loc[idx_max, ["codigo_ejecutor", "bpin"]]
    bpin_representativo = bpin_representativo.rename(columns={"bpin": "bpin_representativo"})

    ejecutores = ejecutores.merge(bpin_representativo, on="codigo_ejecutor", how="left")

    # conteo de proyectos y sector más frecuente, por ejecutor (para la ficha)
    conteo_proy = proyectos.groupby("codigo_ejecutor")["bpin"].nunique().rename("total_proyectos")
    sector_frecuente = (
        proyectos.dropna(subset=["sector"])
        .groupby("codigo_ejecutor")["sector"]
        .agg(lambda s: s.value_counts().idxmax() if len(s) else None)
        .rename("sector_principal")
    )
    ejecutores = ejecutores.merge(conteo_proy, on="codigo_ejecutor", how="left")
    ejecutores = ejecutores.merge(sector_frecuente, on="codigo_ejecutor", how="left")
    ejecutores["total_proyectos"] = ejecutores["total_proyectos"].fillna(0).astype(int)

    con = sqlite3.connect(DB_PATH)
    ejecutores.to_sql("ejecutores", con, if_exists="replace", index=False)
    proyectos.to_sql("proyectos", con, if_exists="replace", index=False)
    resultado.to_sql("resultado_ics", con, if_exists="replace", index=False)

    con.execute("CREATE INDEX IF NOT EXISTS idx_proy_bpin ON proyectos(bpin)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_proy_ejecutor ON proyectos(codigo_ejecutor)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_ejec_codigo ON ejecutores(codigo_ejecutor)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_res_codigo ON resultado_ics(codigo_ejecutor)")
    con.commit()
    con.close()

    print(f"Base de datos generada en {DB_PATH}")
    print(f"  ejecutores: {len(ejecutores)} filas")
    print(f"  proyectos: {len(proyectos)} filas")
    print(f"  resultado_ics: {len(resultado)} filas")


if __name__ == "__main__":
    construir_bd()
