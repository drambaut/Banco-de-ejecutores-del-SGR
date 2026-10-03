"""
main.py — Backend de la plataforma SAR (Sistema de Análisis de Riesgo).

Expone una API JSON que alimenta el frontend (mockup adaptado) con datos
reales calculados por indicador_cumplimiento_historico_v2.py, leídos desde
data/sar.db (generada por build_db.py).

Correr localmente:
    python app/main.py
  (o en producción: gunicorn app.main:app)

Endpoints:
    GET /api/perfil/<codigo_ejecutor>  -> perfil de riesgo del ejecutor
    GET /api/descriptivo               -> KPIs y agregados para el tablero descriptivo
    GET /api/buscar?q=texto            -> autocompletar ejecutores por nombre/código/NIT
    GET /api/departamentos             -> lista de departamentos (para el filtro)
"""

import sqlite3
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory

BASE_DIR = Path(__file__).parent.parent
DB_PATH = BASE_DIR / "data" / "sar.db"
STATIC_DIR = Path(__file__).parent / "static"

app = Flask(__name__, static_folder=None)


def conectar():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


# =============================================================================
# LÓGICA DE NEGOCIO
# =============================================================================

# Umbrales de 4 bandas para calzar con el gauge de 4 colores del mockup
# (Bajo/Medio/Alto/Crítico). La metodología v2 define 3 niveles
# (Bajo/Medio/Alto); aquí se subdivide "Alto" en Alto/Crítico solo para la
# visualización. AJUSTAR si el DNP define oficialmente estos cortes.
def nivel_4_bandas(puntaje: float) -> str:
    if puntaje < 30:
        return "Bajo"
    if puntaje < 60:
        return "Medio"
    if puntaje < 85:
        return "Alto"
    return "Crítico"


def construir_perfil(con, codigo_ejecutor: str):
    ejecutor = con.execute(
        "SELECT * FROM ejecutores WHERE codigo_ejecutor = ?", (codigo_ejecutor,)
    ).fetchone()
    if ejecutor is None:
        return None

    resultado = con.execute(
        "SELECT * FROM resultado_ics WHERE codigo_ejecutor = ?", (codigo_ejecutor,)
    ).fetchone()
    # proyecto representativo del ejecutor (el de mayor valor), solo para
    # mostrar contexto -- ya no se busca por BPIN
    proyecto = con.execute(
        "SELECT * FROM proyectos WHERE bpin = ?", (ejecutor["bpin_representativo"],)
    ).fetchone()

    perfil_riesgo = None
    capacidades = {"administrativa": None, "financiera": None, "institucional": None}

    if resultado is not None:
        puntaje_preciso = resultado["puntaje_riesgo"]
        puntaje = round(puntaje_preciso, 1)
        claves_indicadores = [
            ("ICH", "puntaje_ich"), ("ICCI", "puntaje_icci"), ("IE", "puntaje_ie"),
            ("IMA", "puntaje_ima"), ("IAG", "puntaje_iag"),
        ]
        disponibles = [(sigla, clave, resultado[clave]) for sigla, clave in claves_indicadores if resultado[clave] is not None]
        peso_efectivo = 1 / len(disponibles) if disponibles else None
        desglose_indicadores = [
            {"sigla": sigla, "valor": valor, "peso": peso_efectivo, "aporte": valor * peso_efectivo}
            if valor is not None else {"sigla": sigla, "valor": None, "peso": None, "aporte": None}
            for sigla, clave in claves_indicadores
            for valor in [resultado[clave]]
        ]
        perfil_riesgo = {
            "puntaje": puntaje,
            "puntaje_preciso": puntaje_preciso,
            "nivel_3_bandas": resultado["nivel_riesgo"],
            "nivel_4_bandas": nivel_4_bandas(puntaje),
            "base_e": round(resultado["base_e"], 3) if resultado["base_e"] is not None else None,
            "inc_valor": round(resultado["inc_valor"], 3) if resultado["inc_valor"] is not None else None,
            "inc_n": round(resultado["inc_n"], 3) if resultado["inc_n"] is not None else None,
            "ich": round(resultado["ich"], 3) if resultado["ich"] is not None else None,
            "n_proyectos": int(resultado["n_proyectos"]) if resultado["n_proyectos"] is not None else 0,
            "ve_e": resultado["ve_e"],
            "vref": resultado["vref"],
            "grupo_capacidad_institucional": ejecutor["capacidad_institucional"],
            "puntaje_ich":  round(resultado["puntaje_ich"],  1) if resultado["puntaje_ich"]  is not None else None,
            "puntaje_icci": round(resultado["puntaje_icci"], 1) if resultado["puntaje_icci"] is not None else None,
            "puntaje_ie":   round(resultado["puntaje_ie"],   1) if resultado["puntaje_ie"]   is not None else None,
            "puntaje_ima":  round(resultado["puntaje_ima"],  1) if resultado["puntaje_ima"]  is not None else None,
            "puntaje_iag":  round(resultado["puntaje_iag"],  1) if resultado["puntaje_iag"]  is not None else None,
            "desglose_indicadores": desglose_indicadores,
        }

        # Base_e ya combina cumplimiento y penalización por reprogramación
        # ponderados por proyecto (metodología M3); se aproxima a una escala
        # 0-100 para la caja "Administrativa" de la UI.
        base_pct = resultado["base_e"] * 100 if resultado["base_e"] is not None else None
        capacidades["administrativa"] = {
            "score": round(base_pct) if base_pct is not None else None,
            "disponible": base_pct is not None,
            "variables": [
                {"nombre": "Cumplimiento ponderado (Base_e)", "puntos": round(base_pct) if base_pct is not None else None},
                {"nombre": "Éxito en contratación", "puntos": None},
                {"nombre": "Experiencia en el sector", "puntos": None},
            ],
            "nota": "Requieren datos de SECOP.",
        }
        capacidades["financiera"] = {
            "score": None, "disponible": False,
            "nota": "Requiere ejecución presupuestal, patrimonio y desviación en costo (no integrado aún).",
        }
        capacidades["institucional"] = {
            "score": None, "disponible": False,
            "nota": "Requiere histórico de entes de control / sanciones (no integrado aún).",
        }

    # Agregados reales del portafolio. No se calculan ni modifican indicadores.
    filas_estado = con.execute(
        """
        SELECT COALESCE(estado, 'Sin estado') AS nombre, COUNT(DISTINCT bpin) AS proyectos,
               COALESCE(SUM(valor_total_proyecto), 0) AS valor
        FROM proyectos WHERE codigo_ejecutor = ? GROUP BY COALESCE(estado, 'Sin estado')
        ORDER BY valor DESC
        """, (codigo_ejecutor,)
    ).fetchall()
    filas_sector = con.execute(
        """
        SELECT COALESCE(sector, 'Sin clasificar') AS nombre, COUNT(DISTINCT bpin) AS proyectos,
               COALESCE(SUM(valor_total_proyecto), 0) AS valor
        FROM proyectos WHERE codigo_ejecutor = ? GROUP BY COALESCE(sector, 'Sin clasificar')
        ORDER BY valor DESC LIMIT 8
        """, (codigo_ejecutor,)
    ).fetchall()
    total_valor = sum(f["valor"] for f in filas_estado)
    sgr_balance = con.execute(
        """
        SELECT COUNT(DISTINCT bpin) AS proyectos_con_valor_sgr, COALESCE(SUM(valor_sgr), 0) AS valor_sgr
        FROM proyectos WHERE codigo_ejecutor = ? AND valor_sgr IS NOT NULL
        """, (codigo_ejecutor,)
    ).fetchone()
    # Los conteos incluyen cada BPIN por departamento real. El valor solo se
    # agrega si el BPIN tiene un único departamento, para no duplicarlo ni
    # repartirlo arbitrariamente en localizaciones múltiples.
    filas_departamento = con.execute(
        """
        WITH ubicaciones AS (
            SELECT bpin, COUNT(DISTINCT departamento) AS n_departamentos
            FROM proyecto_departamento GROUP BY bpin
        )
        SELECT pd.departamento,
               COUNT(DISTINCT pd.bpin) AS proyectos,
               COALESCE(SUM(CASE WHEN u.n_departamentos = 1 THEN p.valor_total_proyecto END), 0) AS valor_total
        FROM proyecto_departamento pd
        JOIN proyectos p ON p.bpin = pd.bpin
        JOIN ubicaciones u ON u.bpin = pd.bpin
        WHERE p.codigo_ejecutor = ?
        GROUP BY pd.departamento
        ORDER BY valor_total DESC, proyectos DESC, pd.departamento ASC
        """, (codigo_ejecutor,)
    ).fetchall()
    resumen_territorial = con.execute(
        """
        WITH ubicaciones AS (
            SELECT bpin, COUNT(DISTINCT departamento) AS n_departamentos
            FROM proyecto_departamento GROUP BY bpin
        )
        SELECT
            COALESCE(SUM(CASE WHEN u.n_departamentos = 1 THEN 1 ELSE 0 END), 0) AS proyectos_unidepartamentales,
            COALESCE(SUM(CASE WHEN u.n_departamentos > 1 THEN 1 ELSE 0 END), 0) AS proyectos_multidepartamentales,
            COALESCE(SUM(CASE WHEN u.n_departamentos = 1 THEN p.valor_total_proyecto END), 0) AS valor_territorial_total,
            COUNT(u.bpin) AS proyectos_clasificados
        FROM proyectos p
        LEFT JOIN ubicaciones u ON u.bpin = p.bpin
        WHERE p.codigo_ejecutor = ?
        """, (codigo_ejecutor,),
    ).fetchone()
    portafolio = {
        "total_proyectos": ejecutor["total_proyectos"],
        "valor_total": total_valor,
        "por_estado": [dict(f) for f in filas_estado],
        "por_sector": [dict(f) for f in filas_sector],
        "valor_sgr": sgr_balance["valor_sgr"],
        "proyectos_con_valor_sgr": sgr_balance["proyectos_con_valor_sgr"],
        "total_sectores": con.execute(
            "SELECT COUNT(DISTINCT sector) FROM proyectos WHERE codigo_ejecutor = ? AND sector IS NOT NULL",
            (codigo_ejecutor,),
        ).fetchone()[0],
        "por_departamento": [dict(f) for f in filas_departamento],
        "proyectos_unidepartamentales": resumen_territorial["proyectos_unidepartamentales"],
        "proyectos_multidepartamentales": resumen_territorial["proyectos_multidepartamentales"],
        "proyectos_sin_localizacion": ejecutor["total_proyectos"] - resumen_territorial["proyectos_clasificados"],
        "valor_territorial_total": resumen_territorial["valor_territorial_total"],
    }

    igpr_fila = con.execute(
        "SELECT * FROM igpr_entidad WHERE codigo_ejecutor = ?", (codigo_ejecutor,)
    ).fetchone()
    igpr = dict(igpr_fila) if igpr_fila is not None else None

    comparables = []
    if ejecutor["departamento"] and resultado is not None:
        filas = con.execute(
            """
            SELECT e.codigo_ejecutor, e.nombre_ejecutor, e.tipo_ejecutor, e.bpin_representativo,
                   r.puntaje_riesgo, r.nivel_riesgo
            FROM resultado_ics r
            JOIN ejecutores e ON e.codigo_ejecutor = r.codigo_ejecutor
            WHERE e.departamento = ?
              AND e.codigo_ejecutor != ?
              AND r.puntaje_riesgo < ?
            ORDER BY r.puntaje_riesgo ASC
            LIMIT 3
            """,
            (ejecutor["departamento"], codigo_ejecutor, resultado["puntaje_riesgo"]),
        ).fetchall()
        etiquetas = ["1° Entidad Sugerida", "2° Entidad Sugerida", "3° Entidad Sugerida"]
        for i, fila in enumerate(filas[:3]):
            comparables.append({
                "etiqueta": etiquetas[i] if i < len(etiquetas) else f"{i+1}° Entidad Sugerida",
                "codigo_ejecutor": fila["codigo_ejecutor"],
                "nombre_ejecutor": fila["nombre_ejecutor"],
                "tipo": fila["tipo_ejecutor"],
                "puntaje": round(fila["puntaje_riesgo"], 1),
                "nivel": fila["nivel_riesgo"],
                "codigo_para_buscar": fila["codigo_ejecutor"],
            })

    return {
        "ejecutor": {
            "codigo_ejecutor": ejecutor["codigo_ejecutor"],
            "nombre_ejecutor": ejecutor["nombre_ejecutor"],
            "etiqueta": "Entidad objeto de análisis",
            "nit": ejecutor["nit"],
            "departamento": ejecutor["departamento"],
            "region": ejecutor["region"],
            "tipo_ejecutor": ejecutor["tipo_ejecutor"],
            "total_proyectos": ejecutor["total_proyectos"],
            "sector_principal": ejecutor["sector_principal"] or "No disponible",
        },
        "proyecto_representativo": ({
            "bpin": proyecto["bpin"],
            "nombre_proyecto": proyecto["nombre_proyecto"],
            "estado": proyecto["estado"],
            "valor_total": proyecto["valor_total_proyecto"],
            "fecha_inicial_programacion": proyecto["fecha_inicial_programacion"],
            "fecha_final_programacion": proyecto["fecha_final_programacion"],
            "valor_sgr": proyecto["valor_sgr"],
        } if proyecto else None),
        "perfil_riesgo": perfil_riesgo,
        "portafolio": portafolio,
        "igpr": igpr,
        "capacidades": capacidades,
        "comparables": comparables,
    }


# =============================================================================
# ENDPOINTS
# =============================================================================

@app.route("/api/perfil/<codigo_ejecutor>")
def perfil_por_codigo(codigo_ejecutor):
    con = conectar()
    try:
        perfil = construir_perfil(con, codigo_ejecutor)
        if perfil is None:
            return jsonify({"error": f"No se encontró el ejecutor con código {codigo_ejecutor}"}), 404
        return jsonify(perfil)
    finally:
        con.close()


@app.route("/api/buscar")
def buscar():
    q = request.args.get("q", "").strip()
    departamento = request.args.get("departamento", "").strip()
    region = request.args.get("region", "").strip()
    tipo_ejecutor = request.args.get("tipo_ejecutor", "").strip()
    if len(q) == 1:
        return jsonify({"resultados": [], "total": 0})
    con = conectar()
    try:
        condiciones: list[str] = []
        parametros: list[str] = []
        if q:
            like = f"%{q}%"
            condiciones.append("(e.nombre_ejecutor LIKE ? OR e.nit LIKE ? OR e.codigo_ejecutor LIKE ?)")
            parametros.extend([like, like, like])
        if departamento:
            condiciones.append("UPPER(e.departamento) = UPPER(?)")
            parametros.append(departamento)
        if region:
            condiciones.append("UPPER(e.region) = UPPER(?)")
            parametros.append(region)
        if tipo_ejecutor:
            condiciones.append("UPPER(e.tipo_ejecutor) = UPPER(?)")
            parametros.append(tipo_ejecutor)
        where = f"WHERE {' AND '.join(condiciones)}" if condiciones else ""
        filas = con.execute(
            f"""
            SELECT DISTINCT e.codigo_ejecutor, e.nombre_ejecutor, e.nit, e.departamento, e.region, e.tipo_ejecutor
            FROM ejecutores e
            {where}
            ORDER BY e.nombre_ejecutor
            LIMIT 100
            """, parametros,
        ).fetchall()
        resultados = [
            {"codigo_ejecutor": f["codigo_ejecutor"], "nombre_ejecutor": f["nombre_ejecutor"],
             "nit": f["nit"], "departamento": f["departamento"], "region": f["region"],
             "tipo_ejecutor": f["tipo_ejecutor"]}
            for f in filas
        ]
        return jsonify({"resultados": resultados, "total": len(resultados)})
    finally:
        con.close()


@app.route("/api/departamentos")
def departamentos():
    con = conectar()
    try:
        filas = con.execute(
            "SELECT DISTINCT departamento FROM ejecutores WHERE departamento IS NOT NULL ORDER BY departamento"
        ).fetchall()
        return jsonify({"departamentos": [f["departamento"] for f in filas]})
    finally:
        con.close()


def proyecto_ocad(con, bpin: str):
    """Ficha descriptiva de un BPIN; no aplica reglas de designación."""
    fila = con.execute(
        """
        SELECT p.*, e.nombre_ejecutor, e.nit, e.tipo_ejecutor,
               e.departamento AS departamento_ejecutor, e.region,
               r.puntaje_riesgo, r.nivel_riesgo
        FROM proyectos p
        LEFT JOIN ejecutores e ON e.codigo_ejecutor = p.codigo_ejecutor
        LEFT JOIN resultado_ics r ON r.codigo_ejecutor = p.codigo_ejecutor
        WHERE p.bpin = ?
        """, (bpin,)
    ).fetchone()
    if fila is None:
        return None

    departamentos = [x["departamento"] for x in con.execute(
        "SELECT departamento FROM proyecto_departamento WHERE bpin = ? ORDER BY departamento", (bpin,)
    ).fetchall()]
    referencias = []
    if departamentos:
        marcas = ",".join("?" for _ in departamentos)
        referencias = [dict(x) for x in con.execute(
            f"""
            SELECT e.codigo_ejecutor, e.nombre_ejecutor, e.nit, e.departamento,
                   e.tipo_ejecutor, r.puntaje_riesgo, r.nivel_riesgo
            FROM ejecutores e
            JOIN resultado_ics r ON r.codigo_ejecutor = e.codigo_ejecutor
            WHERE e.tipo_ejecutor = 'DEPARTAMENTO'
              AND e.departamento IN ({marcas})
            ORDER BY e.departamento, e.nombre_ejecutor, e.codigo_ejecutor
            """, departamentos,
        ).fetchall()]

    proyecto = dict(fila)
    proyecto["departamentos_localizacion"] = departamentos
    proyecto["referencias_territoriales"] = referencias
    # El texto se conserva sin inferir municipios; sirve como trazabilidad de
    # la localización que reporta el Balance.
    proyecto["municipios_localizacion"] = proyecto["localizacion_proyecto"]
    return proyecto


@app.route("/api/proyectos")
def buscar_proyectos():
    q = request.args.get("q", "").strip()
    if len(q) == 1:
        return jsonify({"resultados": [], "total": 0})
    con = conectar()
    try:
        condiciones, parametros = [], []
        if q:
            like = f"%{q}%"
            condiciones.append("(p.bpin LIKE ? OR p.nombre_proyecto LIKE ?)")
            parametros.extend([like, like])
        where = f"WHERE {' AND '.join(condiciones)}" if condiciones else ""
        filas = con.execute(
            f"""
            SELECT p.bpin, p.nombre_proyecto, p.estado, p.codigo_ejecutor,
                   e.nombre_ejecutor
            FROM proyectos p
            LEFT JOIN ejecutores e ON e.codigo_ejecutor = p.codigo_ejecutor
            {where}
            ORDER BY CASE WHEN p.bpin = ? THEN 0 ELSE 1 END, p.nombre_proyecto, p.bpin
            LIMIT 20
            """, [*parametros, q],
        ).fetchall()
        return jsonify({"resultados": [dict(x) for x in filas], "total": len(filas)})
    finally:
        con.close()


@app.route("/api/proyecto/<bpin>")
def proyecto_por_bpin(bpin):
    con = conectar()
    try:
        proyecto = proyecto_ocad(con, bpin)
        if proyecto is None:
            return jsonify({"error": f"No se encontró el proyecto con BPIN {bpin}"}), 404
        return jsonify(proyecto)
    finally:
        con.close()


@app.route("/api/ranking")
def ranking():
    tipo_ejecutor = request.args.get("tipo_ejecutor")
    region        = request.args.get("region")
    departamento  = request.args.get("departamento")

    con = conectar()
    try:
        filtros = []
        params: list = []
        if tipo_ejecutor and tipo_ejecutor != "Todos":
            filtros.append("e.tipo_ejecutor = ?")
            params.append(tipo_ejecutor)
        if region and region != "Todas":
            filtros.append("e.region = ?")
            params.append(region)
        if departamento and departamento != "Todos":
            filtros.append("e.departamento = ?")
            params.append(departamento)
        where = f"WHERE {' AND '.join(filtros)}" if filtros else ""

        base_sql = f"""
            SELECT e.codigo_ejecutor, e.tipo_ejecutor, e.region, r.puntaje_riesgo, r.nivel_riesgo
            FROM resultado_ics r JOIN ejecutores e ON e.codigo_ejecutor = r.codigo_ejecutor
            {where}
        """
        mejores = con.execute(base_sql + " ORDER BY r.puntaje_riesgo ASC  LIMIT 5", params).fetchall()
        peores  = con.execute(base_sql + " ORDER BY r.puntaje_riesgo DESC LIMIT 5", params).fetchall()
        return jsonify({
            "mejores": [dict(f) for f in mejores],
            "peores":  [dict(f) for f in peores],
        })
    finally:
        con.close()


@app.route("/api/tipos")
def tipos():
    con = conectar()
    try:
        filas = con.execute(
            "SELECT DISTINCT tipo_ejecutor FROM ejecutores WHERE tipo_ejecutor IS NOT NULL ORDER BY tipo_ejecutor"
        ).fetchall()
        return jsonify({"tipos": [f["tipo_ejecutor"] for f in filas]})
    finally:
        con.close()


@app.route("/api/regiones")
def regiones():
    con = conectar()
    try:
        filas = con.execute(
            "SELECT DISTINCT region FROM ejecutores WHERE region IS NOT NULL ORDER BY region"
        ).fetchall()
        return jsonify({"regiones": [f["region"] for f in filas]})
    finally:
        con.close()


@app.route("/api/descriptivo")
def descriptivo():
    tipo_ejecutor = request.args.get("tipo_ejecutor")
    region = request.args.get("region")
    departamento = request.args.get("departamento")

    con = conectar()
    try:
        filtros = []
        params: list = []
        if tipo_ejecutor and tipo_ejecutor != "Todos":
            filtros.append("e.tipo_ejecutor = ?")
            params.append(tipo_ejecutor)
        if region and region != "Todas":
            filtros.append("e.region = ?")
            params.append(region)
        if departamento and departamento != "Todos":
            filtros.append("e.departamento = ?")
            params.append(departamento)
        where = f"WHERE {' AND '.join(filtros)}" if filtros else ""

        filas = con.execute(
            f"""
            SELECT r.*, e.tipo_ejecutor, e.region, e.departamento
            FROM resultado_ics r
            JOIN ejecutores e ON e.codigo_ejecutor = r.codigo_ejecutor
            {where}
            """,
            params,
        ).fetchall()

        total_ejecutores = len(filas)
        if total_ejecutores == 0:
            return jsonify({"total_ejecutores": 0})

        puntajes = [f["puntaje_riesgo"] for f in filas]

        bins = [0] * 10
        for p in puntajes:
            idx = min(int(p // 10), 9)
            bins[idx] += 1

        conteo_4_bandas = {"Bajo": 0, "Medio": 0, "Alto": 0, "Crítico": 0}
        for p in puntajes:
            conteo_4_bandas[nivel_4_bandas(p)] += 1

        por_tipo: dict = {}
        for f in filas:
            t = f["tipo_ejecutor"] or "Sin clasificar"
            por_tipo.setdefault(t, []).append(f["puntaje_riesgo"])
        promedio_por_tipo = {t: round(sum(v) / len(v), 1) for t, v in por_tipo.items()}

        por_region: dict = {}
        for f in filas:
            r = f["region"] or "Sin clasificar"
            por_region.setdefault(r, []).append(f["puntaje_riesgo"])
        promedio_por_region = {r: round(sum(v) / len(v), 1) for r, v in por_region.items()}

        codigos_filtrados = [f["codigo_ejecutor"] for f in filas]
        placeholders = ",".join("?" * len(codigos_filtrados))
        total_proyectos = con.execute(
            f"SELECT COUNT(*) as n FROM proyectos WHERE codigo_ejecutor IN ({placeholders})",
            codigos_filtrados,
        ).fetchone()["n"]
        valor_total = con.execute(
            f"SELECT SUM(valor_total_proyecto) as v FROM proyectos WHERE codigo_ejecutor IN ({placeholders})",
            codigos_filtrados,
        ).fetchone()["v"]
        estado_counts = con.execute(
            f"SELECT estado, COUNT(*) as n FROM proyectos WHERE codigo_ejecutor IN ({placeholders}) GROUP BY estado",
            codigos_filtrados,
        ).fetchall()

        return jsonify({
            "total_ejecutores": total_ejecutores,
            "total_proyectos": total_proyectos,
            "valor_total_proyectos": valor_total,
            "puntaje_promedio": round(sum(puntajes) / len(puntajes), 1),
            "pct_alto_critico": round(
                (conteo_4_bandas["Alto"] + conteo_4_bandas["Crítico"]) / total_ejecutores * 100, 1
            ),
            "histograma_bins_10": bins,
            "conteo_4_bandas": conteo_4_bandas,
            "promedio_riesgo_por_tipo": promedio_por_tipo,
            "promedio_riesgo_por_region": promedio_por_region,
            "estado_proyectos": {f["estado"] or "Sin estado": f["n"] for f in estado_counts},
        })
    finally:
        con.close()


# =============================================================================
# FRONTEND (archivos estáticos)
# =============================================================================

@app.route("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.route("/<path:filename>")
def estaticos(filename):
    return send_from_directory(STATIC_DIR, filename)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
