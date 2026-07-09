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
        puntaje = round(resultado["puntaje_riesgo"], 1)
        perfil_riesgo = {
            "puntaje": puntaje,
            "nivel_3_bandas": resultado["nivel_riesgo"],
            "nivel_4_bandas": nivel_4_bandas(puntaje),
            "tbc": round(resultado["tbc"], 3),
            "fc": round(resultado["fc"], 3),
            "pen": round(resultado["pen"], 3),
            "ics": round(resultado["ics"], 3),
            "n_proyectos": int(resultado["n_proyectos"]),
            "reprogramaciones_no_permitidas": int(resultado["reprogramaciones_no_permitidas"]),
            "descuento_pct_por_reprogramacion": round(resultado["descuento_pct"], 1),
            "grupo_capacidad_institucional": ejecutor["capacidad_institucional"],
        }

        tbc_pct = resultado["tbc"] * 100
        capacidades["administrativa"] = {
            "score": round(tbc_pct),
            "disponible": True,
            "variables": [
                {"nombre": "Cumplimiento de metas (TBC)", "puntos": round(tbc_pct)},
                {"nombre": "Retraso programación mensual", "puntos": round(100 - tbc_pct)},
                {"nombre": "Éxito en contratación", "puntos": None},
                {"nombre": "Experiencia en el sector", "puntos": None},
            ],
            "nota": "Aproximado a partir de TBC. 'Éxito en contratación' y "
                    "'experiencia en el sector' requieren datos de SECOP (pendiente).",
        }
        capacidades["financiera"] = {
            "score": None, "disponible": False,
            "nota": "Requiere ejecución presupuestal, patrimonio y desviación en costo (no integrado aún).",
        }
        capacidades["institucional"] = {
            "score": None, "disponible": False,
            "nota": "Requiere histórico de entes de control / sanciones (no integrado aún).",
        }

    comparables = []
    if ejecutor["departamento"]:
        filas = con.execute(
            """
            SELECT e.codigo_ejecutor, e.nombre_ejecutor, e.tipo_ejecutor, e.bpin_representativo,
                   r.puntaje_riesgo, r.nivel_riesgo
            FROM resultado_ics r
            JOIN ejecutores e ON e.codigo_ejecutor = r.codigo_ejecutor
            WHERE e.departamento = ? AND e.codigo_ejecutor != ?
            ORDER BY r.puntaje_riesgo ASC
            LIMIT 3
            """,
            (ejecutor["departamento"], codigo_ejecutor),
        ).fetchall()
        etiquetas = ["1° Entidad Sugerida", "2° Entidad Sugerida", "3° Entidad Sugerida"]
        for i, fila in enumerate(filas[:3]):
            comparables.append({
                "etiqueta": etiquetas[i] if i < len(etiquetas) else f"{i+1}° Entidad Sugerida",
                "codigo_ejecutor": fila["codigo_ejecutor"],
                "nombre": fila["nombre_ejecutor"],
                "tipo": fila["tipo_ejecutor"],
                "puntaje": round(fila["puntaje_riesgo"], 1),
                "nivel": fila["nivel_riesgo"],
                "codigo_para_buscar": fila["codigo_ejecutor"],
            })

    return {
        "ejecutor": {
            "codigo_ejecutor": ejecutor["codigo_ejecutor"],
            "nombre": ejecutor["nombre_ejecutor"],
            "nit": ejecutor["nit"],
            "departamento": ejecutor["departamento"],
            "region": ejecutor["region"],
            "tipo_ejecutor": ejecutor["tipo_ejecutor"],
            "total_proyectos": ejecutor["total_proyectos"],
            "sector_principal": ejecutor["sector_principal"] or "No disponible",
        },
        "proyecto_representativo": ({
            "nombre": proyecto["nombre_proyecto"] if proyecto["nombre_proyecto"] else "No disponible",
            "estado": proyecto["estado"],
            "valor_total": proyecto["valor_total_proyecto"],
        } if proyecto else None),
        "perfil_riesgo": perfil_riesgo,
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
    q = request.args.get("q", "")
    if len(q) < 2:
        return jsonify({"resultados": []})
    con = conectar()
    try:
        like = f"%{q}%"
        filas = con.execute(
            """
            SELECT DISTINCT e.codigo_ejecutor, e.nombre_ejecutor, e.departamento
            FROM ejecutores e
            WHERE e.nombre_ejecutor LIKE ? OR e.nit LIKE ? OR e.codigo_ejecutor LIKE ?
            LIMIT 15
            """,
            (like, like, like),
        ).fetchall()
        resultados = [
            {"codigo_ejecutor": f["codigo_ejecutor"], "nombre": f["nombre_ejecutor"],
             "departamento": f["departamento"]}
            for f in filas
        ]
        return jsonify({"resultados": resultados})
    finally:
        con.close()
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


@app.route("/api/ranking")
def ranking():
    con = conectar()
    try:
        mejores = con.execute(
            """
            SELECT e.nombre_ejecutor, e.tipo_ejecutor, e.region, r.puntaje_riesgo, r.nivel_riesgo
            FROM resultado_ics r JOIN ejecutores e ON e.codigo_ejecutor = r.codigo_ejecutor
            ORDER BY r.puntaje_riesgo ASC LIMIT 5
            """
        ).fetchall()
        peores = con.execute(
            """
            SELECT e.nombre_ejecutor, e.tipo_ejecutor, e.region, r.puntaje_riesgo, r.nivel_riesgo
            FROM resultado_ics r JOIN ejecutores e ON e.codigo_ejecutor = r.codigo_ejecutor
            ORDER BY r.puntaje_riesgo DESC LIMIT 5
            """
        ).fetchall()
        return jsonify({
            "mejores": [dict(f) for f in mejores],
            "peores": [dict(f) for f in peores],
        })
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

        total_proyectos = con.execute("SELECT COUNT(*) as n FROM proyectos").fetchone()["n"]
        valor_total = con.execute("SELECT SUM(valor_total_proyecto) as v FROM proyectos").fetchone()["v"]
        estado_counts = con.execute(
            "SELECT estado, COUNT(*) as n FROM proyectos GROUP BY estado"
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
