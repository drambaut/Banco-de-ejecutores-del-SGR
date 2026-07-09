"""
analisis_sensibilidad_tau.py
===============================
Corre el pipeline del ICS con distintos valores de TAU (tolerancia de
desviación en el cumplimiento por periodo) para evaluar qué tan sensible es
el resultado a ese parámetro, y para escoger un valor de tau que tenga
sentido frente al comportamiento real de la ejecución del SGR.

No modifica la normalización ni ningún otro paso: solo cambia
TAU_TOLERANCIA, tal como se pidió, para aislar el efecto de ese único
parámetro.

POR QUÉ ESTOS 3 VALORES DE TAU (5%, 10%, 15%)
------------------------------------------------
Se calculó la desviación |Ejecutado/Programado - 1| sobre las ~549.000 filas
reales de la hoja "Periodos" del Excel maestro. Hallazgos:

    Percentil 50 (mediana) de |desviación|:  ~19%
    % de periodos que cumplen con tau=5%  :  42.6%
    % de periodos que cumplen con tau=10% :  46.1%
    % de periodos que cumplen con tau=15% :  48.4%
    % de periodos que cumplen con tau=30% :  53.2%  (ya no aporta mucho más)

Es decir, la mediana de desviación real (~19%) ya es 4 veces mayor que el
5% actual, lo que sugiere que 5% es una tolerancia poco realista para
ejecución mensual de proyectos de infraestructura (donde el desfase de caja
entre lo programado y lo pagado en un mes puntual es normal). Subir a 10-15%
sigue siendo exigente, pero no penaliza automáticamente a casi todo el
universo por ruido de calendario. No se prueba con valores como 50-90%
porque ahí ya no se estaría midiendo cumplimiento sino solamente que el
proyecto avanzó algo, lo cual pierde sentido como indicador de riesgo.
"""

import pandas as pd

from indicador_cumplimiento_historico import calcular_indicador_completo
from calcular_ics_desde_maestro import leer_maestro, preparar_insumos

VALORES_TAU_A_PROBAR = [0.05, 0.10, 0.15]  # 5% (actual), 10% y 15% (alternativas razonables)

MAESTRO = "EXCEL_MAESTRO_ICS.xlsx"
SALIDA = "ANALISIS_SENSIBILIDAD_TAU.xlsx"

EJECUTORES_REFERENCIA = {
    "73000": "DEPARTAMENTO DEL TOLIMA",
    "19000": "DEPARTAMENTO DEL CAUCA",
    "63000": "DEPARTAMENTO DEL QUINDIO",
}


def correr_para_tau(df_periodos, df_proyectos, df_reprog, df_capacidad, tau):
    resultado = calcular_indicador_completo(
        df_periodos, df_proyectos, df_reprog, df_capacidad, tau=tau
    )
    resultado["tau_usado"] = tau
    return resultado


def resumir(resultado: pd.DataFrame, tau: float) -> dict:
    conteo = resultado["nivel_riesgo"].value_counts()
    total = len(resultado)
    return {
        "tau": tau,
        "tbc_mediana": resultado["tbc"].median(),
        "tbc_p25": resultado["tbc"].quantile(0.25),
        "tbc_p75": resultado["tbc"].quantile(0.75),
        "riesgo_alto_n": conteo.get("Riesgo Alto", 0),
        "riesgo_alto_pct": conteo.get("Riesgo Alto", 0) / total * 100,
        "riesgo_medio_n": conteo.get("Riesgo Medio", 0),
        "riesgo_medio_pct": conteo.get("Riesgo Medio", 0) / total * 100,
        "riesgo_bajo_n": conteo.get("Riesgo Bajo", 0),
        "riesgo_bajo_pct": conteo.get("Riesgo Bajo", 0) / total * 100,
    }


if __name__ == "__main__":
    ejecutores, proyectos, periodos, reprogramaciones = leer_maestro(MAESTRO)
    df_periodos, df_proyectos, df_reprog, df_capacidad = preparar_insumos(
        ejecutores, proyectos, periodos, reprogramaciones
    )

    resultados_por_tau = {}
    resumenes = []

    for tau in VALORES_TAU_A_PROBAR:
        print(f"Calculando con tau={tau:.0%} ...")
        resultado = correr_para_tau(df_periodos, df_proyectos, df_reprog, df_capacidad, tau)
        resultados_por_tau[tau] = resultado
        resumenes.append(resumir(resultado, tau))

    df_resumen = pd.DataFrame(resumenes)

    print("\n=== RESUMEN COMPARATIVO POR TAU (misma normalización min-max actual) ===\n")
    print(df_resumen.to_string(index=False))

    print("\n=== LOS 3 EJECUTORES DE REFERENCIA, PARA CADA TAU ===\n")
    for tau, resultado in resultados_por_tau.items():
        print(f"--- tau = {tau:.0%} ---")
        filtro = resultado[resultado["codigo_ejecutor"].isin(EJECUTORES_REFERENCIA)]
        print(filtro[["codigo_ejecutor", "tbc", "ics", "ics_norm", "puntaje_riesgo", "nivel_riesgo"]]
              .to_string(index=False))
        print()

    with pd.ExcelWriter(SALIDA, engine="openpyxl") as writer:
        df_resumen.to_excel(writer, sheet_name="Resumen_Comparacion", index=False)
        for tau, resultado in resultados_por_tau.items():
            resultado.to_excel(writer, sheet_name=f"Resultado_tau_{int(tau*100)}pct", index=False)

    print(f"\nArchivo generado: {SALIDA}")
