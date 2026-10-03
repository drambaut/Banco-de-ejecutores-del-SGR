"""
build_db.py — Construye la base de datos SQLite que consume la plataforma SAR.

Fuentes:
    data/Resultados_consolidado.xlsx   -> hojas Ejecutores, Proyectos, Resultados

Salida:
    data/sar.db  (SQLite)

Correr cada vez que haya datos nuevos:
    python build_db.py
"""

import re
import sqlite3
from pathlib import Path

import pandas as pd

DIR_DATOS = Path(__file__).parent / "data"
CONSOLIDADO = DIR_DATOS / "Resultados_consolidado.xlsx"
MAESTRO = DIR_DATOS / "EXCEL_MAESTRO_ICS.xlsx"
BALANCE = DIR_DATOS / "Balance seguimiento SGR.xlsx"
IGPR = Path(__file__).parent.parent / "data" / "raw" / "Resultados IGPR - IV trimestre 2025 19032026.xlsx"
DB_PATH = DIR_DATOS / "sar.db"

# Abreviaturas que aparecen entre paréntesis en LOCALIZACIÓN DEL PROYECTO
# del Balance. Se validaron contra las 34 categorías territoriales presentes
# en el corte; no se infieren departamentos a partir del domicilio del ejecutor.
DEPARTAMENTOS_BALANCE = {
    "AMAZ": "AMAZONAS", "ANTI": "ANTIOQUIA", "ARAU": "ARAUCA", "ATLA": "ATLÁNTICO",
    "BOGO": "BOGOTÁ D.C.", "BOLI": "BOLÍVAR", "BOYA": "BOYACÁ", "CALD": "CALDAS",
    "CAQU": "CAQUETÁ", "CASA": "CASANARE", "CAUC": "CAUCA", "CESA": "CESAR",
    "CHOC": "CHOCÓ", "CORD": "CÓRDOBA", "CUND": "CUNDINAMARCA", "GUAI": "GUAINÍA",
    "GUAJ": "LA GUAJIRA", "GUAV": "GUAVIARE", "HUIL": "HUILA", "LA G": "LA GUAJIRA",
    "MAGD": "MAGDALENA", "META": "META", "NARI": "NARIÑO", "NORT": "NORTE DE SANTANDER",
    "PUTU": "PUTUMAYO", "QUIN": "QUINDÍO", "RISA": "RISARALDA",
    "SAN": "SAN ANDRÉS, PROVIDENCIA Y SANTA CATALINA", "SANT": "SANTANDER", "SUCR": "SUCRE",
    "TOLI": "TOLIMA", "VALL": "VALLE DEL CAUCA", "VAUP": "VAUPÉS", "VICH": "VICHADA",
}


def extraer_departamentos_localizacion(valor: str) -> list[str]:
    codigos = {codigo.strip() for codigo in re.findall(r"\(([^)]+)\)", str(valor))}
    desconocidos = codigos.difference(DEPARTAMENTOS_BALANCE)
    if desconocidos:
        raise ValueError(f"Balance/LOCALIZACIÓN DEL PROYECTO: códigos territoriales no reconocidos: {sorted(desconocidos)}")
    return sorted({DEPARTAMENTOS_BALANCE[codigo] for codigo in codigos})


def validar_unicidad(df, columna: str, fuente: str) -> None:
    if df[columna].isna().any() or df[columna].duplicated().any():
        raise ValueError(f"{fuente}: '{columna}' contiene nulos o duplicados; no se realiza una unión ambigua.")


def construir_bd():
    ejecutores = pd.read_excel(CONSOLIDADO, sheet_name="Ejecutores", dtype={"codigo_ejecutor": str})
    proyectos = pd.read_excel(CONSOLIDADO, sheet_name="Proyectos", dtype={"bpin": str, "codigo_ejecutor": str})
    resultado = pd.read_excel(CONSOLIDADO, sheet_name="Resultados", dtype={"codigo_ejecutor": str})

    # El maestro y el corte IGPR se usan solo como enriquecimiento; no alteran
    # los resultados metodológicos contenidos en el consolidado.
    maestro = pd.read_excel(MAESTRO, sheet_name="Proyectos", dtype={"bpin": str})
    campos_maestro = ["bpin", "nombre_proyecto", "fecha_inicial_programacion", "fecha_final_programacion"]
    validar_unicidad(proyectos, "bpin", "Resultados_consolidado/Proyectos")
    validar_unicidad(maestro, "bpin", "EXCEL_MAESTRO_ICS/Proyectos")
    if not proyectos["bpin"].isin(maestro["bpin"]).all():
        raise ValueError("El maestro no cubre todos los BPIN del consolidado; se cancela la unión.")
    proyectos = proyectos.drop(columns=[c for c in campos_maestro[1:] if c in proyectos], errors="ignore")
    proyectos = proyectos.merge(maestro[campos_maestro], on="bpin", how="left", validate="one_to_one")

    entidad = pd.read_excel(IGPR, sheet_name="Entidad", header=7, dtype=str)
    codigo_igpr = "CODIGO ENTIDAD EJECUTORA O BENEFICIARIA \n(A MEDIR)"
    entidad[codigo_igpr] = entidad[codigo_igpr].str.replace(r"\.0$", "", regex=True).str.strip()
    columnas_igpr = {
        codigo_igpr: "codigo_ejecutor",
        "IGPR FINAL DE LA ENTIDAD \nIV TRIM 2025": "igpr_final_entidad",
        "PROMEDIO IGPR PROYECTOS MEDIDOS": "promedio_igpr_proyectos_medidos",
        "RECURSOS SGR\n PROY MEDIDOS": "recursos_sgr_proyectos_medidos",
    }
    igpr_entidad = entidad[list(columnas_igpr)].rename(columns=columnas_igpr)
    validar_unicidad(igpr_entidad, "codigo_ejecutor", "IGPR/Entidad")
    for col in columnas_igpr.values():
        if col != "codigo_ejecutor":
            igpr_entidad[col] = pd.to_numeric(igpr_entidad[col], errors="raise")
    igpr_entidad["periodo_referencia"] = "IV trimestre de 2025"

    # Valor SGR completo del mismo universo de BPIN del consolidado. Se toma
    # del Balance, no del subconjunto de proyectos medidos por IGPR.
    # Metadatos descriptivos para la ficha de proyecto. Se mantienen separados
    # del cálculo MIEE: el consolidado continúa siendo la fuente de puntajes y
    # del valor total del portafolio.
    balance = pd.read_excel(
        BALANCE, sheet_name="PROYECTOS APROBADOS ", header=7,
        usecols=[
            "BPIN", "VALOR SGR", "LOCALIZACIÓN DEL PROYECTO", "PROGRAMA",
            "SUBPROGRAMA", "FECHA APROBACIÓN", "INSTANCIA DE APROBACIÓN INICIAL ",
            "TIPO DE INSTANCIA INICIAL",
        ], dtype={"BPIN": str},
    )
    balance["BPIN"] = balance["BPIN"].str.replace(r"\.0$", "", regex=True).str.strip()
    validar_unicidad(balance, "BPIN", "Balance/Proyectos aprobados")
    campos_balance = balance.rename(columns={
        "BPIN": "bpin",
        "VALOR SGR": "valor_sgr",
        "LOCALIZACIÓN DEL PROYECTO": "localizacion_proyecto",
        "PROGRAMA": "programa",
        "SUBPROGRAMA": "subprograma",
        "FECHA APROBACIÓN": "fecha_aprobacion",
        "INSTANCIA DE APROBACIÓN INICIAL ": "instancia_aprobacion_inicial",
        "TIPO DE INSTANCIA INICIAL": "tipo_instancia_inicial",
    })
    validar_unicidad(campos_balance, "bpin", "Balance/Proyectos aprobados")
    if not proyectos["bpin"].isin(campos_balance["bpin"]).all():
        raise ValueError("El Balance no cubre todos los BPIN del consolidado; se cancela la unión.")
    campos_balance["valor_sgr"] = pd.to_numeric(campos_balance["valor_sgr"], errors="raise")
    proyectos = proyectos.merge(campos_balance, on="bpin", how="left", validate="one_to_one")

    # La localización puede contener varios municipios y departamentos por BPIN.
    # Se conserva en una tabla aparte para no duplicar filas ni valores del portafolio.
    registros_territoriales = []
    for fila in balance[["BPIN", "LOCALIZACIÓN DEL PROYECTO"]].itertuples(index=False):
        for departamento in extraer_departamentos_localizacion(fila[1]):
            registros_territoriales.append({"bpin": fila[0], "departamento": departamento})
    proyecto_departamento = pd.DataFrame(registros_territoriales, columns=["bpin", "departamento"])
    if proyecto_departamento.duplicated(["bpin", "departamento"]).any():
        raise ValueError("Balance/LOCALIZACIÓN DEL PROYECTO: se generaron duplicados BPIN-departamento.")
    if not proyecto_departamento["bpin"].isin(proyectos["bpin"]).all():
        raise ValueError("Balance/LOCALIZACIÓN DEL PROYECTO contiene BPIN fuera del consolidado.")

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
    proyecto_departamento.to_sql("proyecto_departamento", con, if_exists="replace", index=False)
    resultado.to_sql("resultado_ics", con, if_exists="replace", index=False)
    igpr_entidad.to_sql("igpr_entidad", con, if_exists="replace", index=False)

    con.execute("CREATE INDEX IF NOT EXISTS idx_proy_bpin ON proyectos(bpin)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_proy_ejecutor ON proyectos(codigo_ejecutor)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_proy_depto_bpin ON proyecto_departamento(bpin)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_ejec_codigo ON ejecutores(codigo_ejecutor)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_res_codigo ON resultado_ics(codigo_ejecutor)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_igpr_codigo ON igpr_entidad(codigo_ejecutor)")
    con.commit()
    con.close()

    print(f"Base de datos generada en {DB_PATH}")
    print(f"  ejecutores: {len(ejecutores)} filas")
    print(f"  proyectos: {len(proyectos)} filas")
    print(f"  proyecto_departamento: {len(proyecto_departamento)} filas ({balance['BPIN'].nunique() - proyecto_departamento['bpin'].nunique()} BPIN sin localización clasificable)")
    print(f"  resultado_ics: {len(resultado)} filas")
    print(f"  igpr_entidad: {len(igpr_entidad)} filas")


if __name__ == "__main__":
    construir_bd()
