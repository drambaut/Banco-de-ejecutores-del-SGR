"""
Funciones de preprocesamiento compartidas entre notebooks.
Importar con: from src.preprocessing import <funcion>
"""

def ingestar_excel(ruta, hoja, header_rows, skiprows=None, usecols=None):
    """Lee un Excel con headers multinivel y aplana columnas."""
    import pandas as pd
    df = pd.read_excel(
        ruta, sheet_name=hoja, header=header_rows,
        skiprows=skiprows, usecols=usecols, engine="openpyxl"
    )
    if hasattr(df.columns, 'levels'):
        nuevas = []
        for lvls in df.columns:
            partes = [str(p).strip() for p in lvls if "Unnamed" not in str(p)]
            nuevas.append("_".join(partes) if partes else "SIN_NOMBRE")
        df.columns = nuevas
    else:
        df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all").reset_index(drop=True)
    return df
