"""
comps_routes.py — Blueprint con todas las rutas del módulo Auditoría COMPS.
Adaptado de COMPS-MDA (sqlite3) a SQLAlchemy text() con parámetros nombrados.
"""
import os
from io import BytesIO
from datetime import datetime

import pandas as pd
from flask import Blueprint, render_template, request, redirect, url_for, flash, jsonify, send_file
from flask_login import login_required
from sqlalchemy import text
from werkzeug.utils import secure_filename

comps_bp = Blueprint('comps', __name__, url_prefix='/comps')


@comps_bp.app_context_processor
def inject_meses_nombre():
    return dict(MESES_NOMBRE=MESES_NOMBRE)


UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {'.xls', '.xlsx'}

MESES_NOMBRE = {
    '01': 'Enero', '02': 'Febrero', '03': 'Marzo', '04': 'Abril',
    '05': 'Mayo', '06': 'Junio', '07': 'Julio', '08': 'Agosto',
    '09': 'Septiembre', '10': 'Octubre', '11': 'Noviembre', '12': 'Diciembre'
}


def _get_db():
    try:
        from sgos_web.extensions import db
    except ImportError:
        from extensions import db
    return db


def allowed_file(filename):
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS


def build_date_filter(col, anio, mes):
    """Construye cláusula WHERE y dict de params para filtro año/mes (named params)."""
    conditions = []
    params = {}
    if anio:
        conditions.append(f"SUBSTR({col}, 1, 4) = :anio_{col.replace('.', '_')}")
        params[f"anio_{col.replace('.', '_')}"] = anio
    if mes:
        conditions.append(f"SUBSTR({col}, 6, 2) = :mes_{col.replace('.', '_')}")
        params[f"mes_{col.replace('.', '_')}"] = mes
    where = "WHERE " + " AND ".join(conditions) if conditions else ""
    return where, params


def get_anios_meses():
    """Obtiene años y meses disponibles de las tablas COMPS."""
    db = _get_db()
    rows = db.session.execute(text("""
        SELECT DISTINCT fecha FROM (
            SELECT SUBSTR(gaming_date, 1, 7) as fecha FROM srw_jugadores WHERE gaming_date IS NOT NULL
            UNION
            SELECT SUBSTR(fecha_jornada, 1, 7) FROM cortesias WHERE fecha_jornada IS NOT NULL
            UNION
            SELECT SUBSTR(fecha_jornada, 1, 7) FROM premios_comps WHERE fecha_jornada IS NOT NULL
            UNION
            SELECT SUBSTR(fecha_operacion, 1, 7) FROM mesas_puntos WHERE fecha_operacion IS NOT NULL
        ) sub ORDER BY fecha
    """)).fetchall()
    anios = sorted(set(r[0][:4] for r in rows if r[0]))
    meses_num = sorted(set(r[0][5:7] for r in rows if r[0]))
    return anios, meses_num


def _exec(sql_str, params=None):
    db = _get_db()
    result = db.session.execute(text(sql_str), params or {})
    cols = result.keys()
    return [dict(zip(cols, row)) for row in result.fetchall()]


def _exec_one(sql_str, params=None):
    rows = _exec(sql_str, params)
    return rows[0] if rows else {}


# ─────────────────────────── Index + Carga ───────────────────────────

@comps_bp.route('/')
@login_required
def comps_index():
    db = _get_db()
    log_rows = db.session.execute(text("SELECT * FROM carga_log ORDER BY fecha_carga DESC")).fetchall()
    log = [dict(r._mapping) for r in log_rows]
    stats = {}
    for tabla in ['srw_jugadores', 'cortesias', 'premios_comps', 'mesas_puntos']:
        row = db.session.execute(text(f"SELECT COUNT(*) as cnt FROM {tabla}")).fetchone()
        stats[tabla] = row[0]
    return render_template('comps/index.html', log=log, stats=stats)


@comps_bp.route('/cargar', methods=['POST'])
@login_required
def cargar_datos():
    try:
        from sgos_web.comps_engine import (
            cargar_srw, cargar_cortesias, cargar_premios_comps,
            cargar_mesas_puntos, guardar_comps_db, guardar_jefaturas_db,
            actualizar_nombres_cortesias
        )
    except ImportError:
        from comps_engine import (
            cargar_srw, cargar_cortesias, cargar_premios_comps,
            cargar_mesas_puntos, guardar_comps_db, guardar_jefaturas_db,
            actualizar_nombres_cortesias
        )

    db = _get_db()
    resultados = []
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

    file_map = {
        'archivo_srw': ('srw_jugadores', cargar_srw, 'gaming_date'),
        'archivo_cortesias': ('cortesias', cargar_cortesias, 'fecha_jornada'),
        'archivo_premios': ('premios_comps', cargar_premios_comps, 'fecha_jornada'),
        'archivo_mesas_puntos': ('mesas_puntos', cargar_mesas_puntos, 'fecha_operacion'),
    }

    try:
        for field, (tabla, etl_fn, col_fecha) in file_map.items():
            f = request.files.get(field)
            if not f or f.filename == '':
                continue
            if not allowed_file(f.filename):
                flash(f'Archivo no válido: {f.filename}. Solo .xls y .xlsx', 'error')
                continue

            filename = secure_filename(f.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            f.save(filepath)

            filas = guardar_comps_db(db, tabla, etl_fn, filepath, col_fecha, {})

            # Log de carga
            db.session.execute(text(
                "INSERT INTO carga_log (tabla, archivo, filas, fecha_carga) VALUES (:t, :a, :f, :fc)"
            ), {"t": tabla, "a": filename, "f": filas, "fc": datetime.now().isoformat()})
            db.session.commit()

            resultados.append(f"{tabla}: {filas} filas cargadas ({filename})")

        # Jefaturas (archivo separado)
        f_jef = request.files.get('archivo_jefaturas')
        if f_jef and f_jef.filename != '' and allowed_file(f_jef.filename):
            filename = secure_filename(f_jef.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            f_jef.save(filepath)
            guardar_jefaturas_db(db, filepath)
            resultados.append(f"jefaturas+categorias: cargadas ({filename})")

        if resultados:
            actualizar_nombres_cortesias(db)
            flash(' | '.join(resultados), 'success')
        else:
            flash('No se seleccionó ningún archivo.', 'error')

    except Exception as e:
        db.session.rollback()
        import logging
        logging.getLogger(__name__).error(f'Error al cargar datos COMPS: {e}')
        flash('Error al cargar datos. Contacta al administrador.', 'error')

    return redirect(url_for('comps.comps_index'))


# ─────────────────────────── Análisis Cortesías ───────────────────────────

@comps_bp.route('/analisis/cortesias')
@login_required
def analisis_cortesias():
    anio = request.args.get('anio', '')
    mes = request.args.get('mes', '')
    anios, meses_disp = get_anios_meses()

    cw, cp = build_date_filter('c.fecha_jornada', anio, mes)
    cw_solo, cp_solo = build_date_filter('fecha_jornada', anio, mes)
    sw, sp = build_date_filter('gaming_date', anio, mes)
    mw_cort, mp_cort = build_date_filter('m.fecha_operacion', anio, mes)

    all_params = {**sp, **mp_cort, **cp}

    resumen = _exec(f"""
        SELECT
            c.cliente_id,
            c.nombre_cliente,
            COUNT(c.id) as total_cortesias,
            SUM(c.micros) as monto_cortesias,
            COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0) as total_coin_in,
            COALESCE(MAX(s.total_promo_in), 0) as total_promo_in,
            COALESCE(MAX(s.total_games), 0) as total_games,
            COALESCE(MAX(s.player_level), '-') as player_level,
            CASE WHEN (COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0)) > 0
                 THEN ROUND((SUM(c.micros) * 100.0 / (COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0)))::numeric, 4)
                 ELSE 0 END as pct_cortesia_coin_in
        FROM cortesias c
        LEFT JOIN (
            SELECT player_id,
                   SUM(coin_in) as total_coin_in,
                   SUM(promo_in) as total_promo_in,
                   SUM(total_games) as total_games,
                   MAX(player_level) as player_level
            FROM srw_jugadores {sw}
            GROUP BY player_id
        ) s ON c.cliente_id = s.player_id
        LEFT JOIN (
            SELECT cliente_id,
                   SUM(coin_in_puntos) as coin_in_mesas
            FROM mesas_puntos m {mw_cort}
            GROUP BY cliente_id
        ) m ON c.cliente_id = m.cliente_id
        {cw}
        GROUP BY c.cliente_id, c.nombre_cliente
        ORDER BY monto_cortesias DESC
    """, all_params)

    por_categoria = _exec(f"""
        SELECT descripcion_cat, COUNT(*) as cantidad,
               SUM(micros) as monto_total
        FROM cortesias {cw_solo}
        GROUP BY descripcion_cat
        ORDER BY monto_total DESC
    """, cp_solo)

    productos_rows = _exec(f"""
        SELECT descripcion_cat, descripcion_prod, COUNT(*) as cantidad,
               SUM(micros) as monto_total
        FROM cortesias {cw_solo}
        GROUP BY descripcion_cat, descripcion_prod
        ORDER BY descripcion_cat, monto_total DESC
    """, cp_solo)
    productos_por_cat = {}
    for r in productos_rows:
        cat = r['descripcion_cat']
        if cat not in productos_por_cat:
            productos_por_cat[cat] = []
        productos_por_cat[cat].append(r)

    # Cortesías por día
    dia_where, dia_params = build_date_filter('fecha_jornada', anio, mes)
    if dia_where:
        dia_where = dia_where + " AND fecha_jornada IS NOT NULL"
    else:
        dia_where = "WHERE fecha_jornada IS NOT NULL"
    por_dia_raw = _exec(f"""
        SELECT fecha_jornada, COUNT(*) as cantidad,
               SUM(micros) as monto_total
        FROM cortesias {dia_where}
        GROUP BY fecha_jornada
        ORDER BY fecha_jornada
    """, dia_params)

    # Coin In por día - MDA
    sw_dia, sp_dia = build_date_filter('gaming_date', anio, mes)
    coin_mda_dia = _exec(f"""
        SELECT gaming_date as fecha, SUM(coin_in) as coin_in
        FROM srw_jugadores {sw_dia}
        {'AND' if sw_dia else 'WHERE'} gaming_date IS NOT NULL
        GROUP BY gaming_date
    """, sp_dia)
    mda_por_fecha = {r['fecha']: r['coin_in'] or 0 for r in coin_mda_dia}

    # Coin In por día - MDJ
    mw_dia, mp_dia = build_date_filter('fecha_operacion', anio, mes)
    coin_mdy_dia = _exec(f"""
        SELECT fecha_operacion as fecha, SUM(coin_in_puntos) as coin_in
        FROM mesas_puntos {mw_dia}
        {'AND' if mw_dia else 'WHERE'} fecha_operacion IS NOT NULL
        GROUP BY fecha_operacion
    """, mp_dia)
    mdy_por_fecha = {r['fecha']: r['coin_in'] or 0 for r in coin_mdy_dia}

    dias_semana = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
    por_dia = []
    for d in por_dia_raw:
        row = dict(d)
        fecha_str = str(row['fecha_jornada'])
        try:
            fecha = datetime.strptime(fecha_str, '%Y-%m-%d')
            row['dia_semana'] = dias_semana[fecha.weekday()]
        except Exception:
            row['dia_semana'] = ''
        row['coin_in_mda'] = mda_por_fecha.get(fecha_str, 0)
        row['coin_in_mdy'] = mdy_por_fecha.get(fecha_str, 0)
        por_dia.append(row)

    totales = _exec_one(f"""
        SELECT COUNT(*) as total_cortesias,
               SUM(micros) as monto_total,
               COUNT(DISTINCT cliente_id) as clientes_unicos
        FROM cortesias {cw_solo}
    """, cp_solo)

    total_coin_in_srw = _exec_one(f"SELECT SUM(coin_in) as total FROM srw_jugadores {sw}", sp)
    mw_total, mp_total = build_date_filter('fecha_operacion', anio, mes)
    total_coin_in_mesas = _exec_one(f"SELECT SUM(coin_in_puntos) as total FROM mesas_puntos {mw_total}", mp_total)
    total_coin_in_combined = (total_coin_in_srw.get('total') or 0) + (total_coin_in_mesas.get('total') or 0)

    return render_template('comps/analisis_cortesias.html',
                           resumen=resumen,
                           por_categoria=por_categoria,
                           productos_por_cat=productos_por_cat,
                           por_dia=por_dia,
                           totales=totales,
                           total_coin_in=total_coin_in_combined,
                           anios=anios, meses_disp=meses_disp,
                           anio_actual=anio, mes_actual=mes)


# ─────────────────────────── Análisis Premios ───────────────────────────

@comps_bp.route('/analisis/premios')
@login_required
def analisis_premios():
    anio = request.args.get('anio', '')
    mes = request.args.get('mes', '')
    anios, meses_disp = get_anios_meses()

    pw, pp = build_date_filter('p.fecha_jornada', anio, mes)
    pw_solo, pp_solo = build_date_filter('fecha_jornada', anio, mes)
    sw, sp = build_date_filter('gaming_date', anio, mes)
    mw_prem, mp_prem = build_date_filter('m.fecha_operacion', anio, mes)

    all_params = {**sp, **mp_prem, **pp}

    por_jugador = _exec(f"""
        SELECT
            p.cliente_id,
            COALESCE(MAX(s.full_name), MAX(m.cliente_nombre), '(Sin nombre)') as nombre,
            COALESCE(MAX(s.player_level), '-') as player_level,
            COUNT(p.id) as total_premios,
            SUM(p.transferencia_final) as monto_total,
            COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0) as total_coin_in,
            COALESCE(MAX(s.total_promo_in), 0) as total_promo_in,
            COALESCE(MAX(s.total_games), 0) as total_games,
            CASE WHEN (COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0)) > 0
                 THEN ROUND((SUM(p.transferencia_final) * 100.0 / (COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0)))::numeric, 4)
                 ELSE 0 END as pct_premio_coin_in
        FROM premios_comps p
        LEFT JOIN (
            SELECT player_id, MAX(full_name) as full_name,
                   MAX(player_level) as player_level,
                   SUM(coin_in) as total_coin_in,
                   SUM(promo_in) as total_promo_in,
                   SUM(total_games) as total_games
            FROM srw_jugadores {sw} GROUP BY player_id
        ) s ON p.cliente_id = s.player_id
        LEFT JOIN (
            SELECT cliente_id, MAX(cliente_nombre) as cliente_nombre,
                   SUM(coin_in_puntos) as coin_in_mesas
            FROM mesas_puntos m {mw_prem}
            GROUP BY cliente_id
        ) m ON p.cliente_id = m.cliente_id
        {pw}
        GROUP BY p.cliente_id
        ORDER BY monto_total DESC
    """, all_params)

    por_tipo = _exec(f"""
        SELECT tipo_pago, COUNT(*) as cantidad,
               SUM(transferencia_final) as monto_total
        FROM premios_comps {pw_solo}
        GROUP BY tipo_pago
        ORDER BY monto_total DESC
    """, pp_solo)

    dia_where, dia_params = build_date_filter('fecha_jornada', anio, mes)
    if dia_where:
        dia_where = dia_where + " AND fecha_jornada IS NOT NULL"
    else:
        dia_where = "WHERE fecha_jornada IS NOT NULL"
    por_dia = _exec(f"""
        SELECT fecha_jornada, COUNT(*) as cantidad,
               SUM(transferencia_final) as monto_total
        FROM premios_comps {dia_where}
        GROUP BY fecha_jornada
        ORDER BY fecha_jornada
    """, dia_params)

    totales = _exec_one(f"""
        SELECT COUNT(*) as total_premios,
               SUM(transferencia_final) as monto_total,
               COUNT(DISTINCT cliente_id) as clientes_unicos
        FROM premios_comps {pw_solo}
    """, pp_solo)

    return render_template('comps/analisis_premios.html',
                           por_jugador=por_jugador,
                           por_tipo=por_tipo,
                           por_dia=por_dia,
                           totales=totales,
                           anios=anios, meses_disp=meses_disp,
                           anio_actual=anio, mes_actual=mes)


# ─────────────────────────── Análisis Resumen ───────────────────────────

@comps_bp.route('/analisis/resumen')
@login_required
def analisis_resumen():
    anio = request.args.get('anio', '')
    mes = request.args.get('mes', '')
    anios, meses_disp = get_anios_meses()

    sw, sp = build_date_filter('gaming_date', anio, mes)
    cw, cparam = build_date_filter('fecha_jornada', anio, mes)
    pw, pparam = build_date_filter('fecha_jornada', anio, mes)
    mw_res, mp_res = build_date_filter('m.fecha_operacion', anio, mes)
    mw_res_solo, mp_res_solo = build_date_filter('fecha_operacion', anio, mes)

    kpis_srw = _exec_one(f"""
        SELECT
            COALESCE(SUM(coin_in), 0) as total_coin_in,
            COALESCE(SUM(promo_in), 0) as total_promo_in,
            COALESCE(SUM(total_games), 0) as total_games,
            COUNT(DISTINCT player_id) as jugadores_srw
        FROM srw_jugadores {sw}
    """, sp)
    kpis_mesas = _exec_one(f"""
        SELECT COALESCE(SUM(coin_in_puntos), 0) as total_coin_in_mesas
        FROM mesas_puntos {mw_res_solo}
    """, mp_res_solo)
    kpis_cort = _exec_one(f"""
        SELECT COALESCE(SUM(micros), 0) as total_cortesias,
               COUNT(DISTINCT cliente_id) as clientes_cortesias
        FROM cortesias {cw}
    """, cparam)
    kpis_prem = _exec_one(f"""
        SELECT COALESCE(SUM(transferencia_final), 0) as total_premios,
               COUNT(DISTINCT cliente_id) as clientes_premios
        FROM premios_comps {pw}
    """, pparam)

    total_coin_in_global = (kpis_srw.get('total_coin_in') or 0) + (kpis_mesas.get('total_coin_in_mesas') or 0)
    kpis = {
        'total_coin_in': total_coin_in_global,
        'total_promo_in': kpis_srw.get('total_promo_in') or 0,
        'total_games': kpis_srw.get('total_games') or 0,
        'jugadores_srw': kpis_srw.get('jugadores_srw') or 0,
        'total_cortesias': kpis_cort.get('total_cortesias') or 0,
        'clientes_cortesias': kpis_cort.get('clientes_cortesias') or 0,
        'total_premios': kpis_prem.get('total_premios') or 0,
        'clientes_premios': kpis_prem.get('clientes_premios') or 0,
    }

    all_params = {**sp, **cparam, **pparam, **mp_res}

    jugadores_raw = _exec(f"""
        SELECT
            s.player_id,
            s.full_name,
            s.player_level,
            s.total_coin_in + COALESCE(m.coin_in_mesas, 0) as total_coin_in,
            s.total_promo_in,
            s.total_games,
            s.dias_jugados,
            COALESCE(c.total_cortesias, 0) as total_cortesias,
            COALESCE(c.monto_cortesias, 0) as monto_cortesias,
            COALESCE(p.total_premios, 0) as total_premios,
            COALESCE(p.monto_premios, 0) as monto_premios
        FROM (
            SELECT player_id, MAX(full_name) as full_name,
                   MAX(player_level) as player_level,
                   SUM(coin_in) as total_coin_in,
                   SUM(promo_in) as total_promo_in,
                   SUM(total_games) as total_games,
                   COUNT(DISTINCT gaming_date) as dias_jugados
            FROM srw_jugadores {sw} GROUP BY player_id
        ) s
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as total_cortesias,
                   SUM(micros) as monto_cortesias
            FROM cortesias {cw} GROUP BY cliente_id
        ) c ON s.player_id = c.cliente_id
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as total_premios,
                   SUM(transferencia_final) as monto_premios
            FROM premios_comps {pw} GROUP BY cliente_id
        ) p ON s.player_id = p.cliente_id
        LEFT JOIN (
            SELECT cliente_id, SUM(coin_in_puntos) as coin_in_mesas
            FROM mesas_puntos m {mw_res}
            GROUP BY cliente_id
        ) m ON s.player_id = m.cliente_id
        WHERE COALESCE(c.total_cortesias, 0) > 0 OR COALESCE(p.total_premios, 0) > 0
        ORDER BY total_coin_in DESC
    """, all_params)

    jugadores = []
    for j in jugadores_raw:
        row = dict(j)
        if total_coin_in_global > 0:
            row['pct_total_coin_in'] = round((row['total_coin_in'] or 0) * 100.0 / total_coin_in_global, 3)
        else:
            row['pct_total_coin_in'] = 0
        jugadores.append(row)

    return render_template('comps/analisis_resumen.html', jugadores=jugadores, kpis=kpis,
                           anios=anios, meses_disp=meses_disp,
                           anio_actual=anio, mes_actual=mes)


# ─────────────────────────── API Charts ───────────────────────────

@comps_bp.route('/api/cortesias-dia')
@login_required
def api_cortesias_dia():
    rows = _exec("""
        SELECT fecha_jornada as fecha, SUM(micros) as monto
        FROM cortesias WHERE fecha_jornada IS NOT NULL
        GROUP BY fecha_jornada ORDER BY fecha_jornada
    """)
    return jsonify(rows)


@comps_bp.route('/api/coin-in-dia')
@login_required
def api_coin_in_dia():
    rows = _exec("""
        SELECT gaming_date as fecha, SUM(coin_in) as monto
        FROM srw_jugadores WHERE gaming_date IS NOT NULL
        GROUP BY gaming_date ORDER BY gaming_date
    """)
    return jsonify(rows)


@comps_bp.route('/api/premios-tipo')
@login_required
def api_premios_tipo():
    rows = _exec("""
        SELECT tipo_pago as tipo, COUNT(*) as cantidad,
               SUM(transferencia_final) as monto
        FROM premios_comps GROUP BY tipo_pago ORDER BY monto DESC
    """)
    return jsonify(rows)


# ─────────────────────────── Control Invitaciones ───────────────────────────

def _get_invitaciones_config():
    prim_row = _exec_one("SELECT porcentaje FROM categorias_nivel WHERE categoria = 'Primario'")
    pct_primario = prim_row.get('porcentaje', 0) or 0
    cat_rows = _exec("SELECT categoria, porcentaje FROM categorias_nivel WHERE categoria != 'Primario'")
    pct_categoria = {r['categoria']: r['porcentaje'] for r in cat_rows}
    return pct_primario, pct_categoria


def _calc_invitaciones(jugadores_raw, pct_primario, pct_categoria, dias_totales):
    resultados = []
    for j in jugadores_raw:
        nivel = j.get('nivel') or ''
        pct_cat = pct_categoria.get(nivel, 0)
        coin_in = j.get('coin_in_mensual') or 0
        invitacion_mensual = coin_in * pct_primario * pct_cat
        monto_micros = j.get('monto_micros') or 0
        saldo = invitacion_mensual - monto_micros
        dias = j.get('dias_asistidos') or 0
        pct_asistencia = round(dias * 100.0 / dias_totales, 1) if dias_totales > 0 else 0
        resultados.append({
            'nombre': j.get('nombre'),
            'nivel': nivel,
            'dias_asistidos': dias,
            'pct_asistencia': pct_asistencia,
            'cant_premios': j.get('cant_premios') or 0,
            'monto_premios': j.get('monto_premios') or 0,
            'coin_in_mensual': coin_in,
            'total_cortesias': j.get('total_cortesias') or 0,
            'monto_micros': monto_micros,
            'invitacion_mensual': round(invitacion_mensual),
            'saldo': round(saldo),
            'pct_cat': pct_cat,
        })
    return resultados


@comps_bp.route('/control/invitaciones')
@login_required
def control_invitaciones():
    anio = request.args.get('anio', '')
    mes = request.args.get('mes', '')
    area = request.args.get('area', '')
    jefe = request.args.get('jefe', '')
    anios, meses_disp = get_anios_meses()

    areas = [r['area'] for r in _exec("SELECT DISTINCT area FROM jefaturas WHERE area != '' ORDER BY area")]

    if area:
        jefes_disp = [(r['usuario_id'], r['nombre']) for r in _exec(
            "SELECT usuario_id, nombre FROM jefaturas WHERE area = :area ORDER BY nombre", {"area": area})]
    else:
        jefes_disp = [(r['usuario_id'], r['nombre']) for r in _exec(
            "SELECT usuario_id, nombre FROM jefaturas ORDER BY nombre")]

    jefe_filter_sql = ""
    jefe_p = {}
    if jefe:
        jefe_filter_sql = " AND c.usuario_id = :jefe_id"
        jefe_p = {"jefe_id": jefe}
    elif area:
        jefe_filter_sql = " AND c.usuario_id IN (SELECT usuario_id FROM jefaturas WHERE area = :area_filter)"
        jefe_p = {"area_filter": area}

    sw, sp = build_date_filter('s.gaming_date', anio, mes)
    cw, cparam = build_date_filter('c.fecha_jornada', anio, mes)
    pw, pparam = build_date_filter('p.fecha_jornada', anio, mes)
    mw, mparam = build_date_filter('m.fecha_operacion', anio, mes)

    sw_solo, sp_solo = build_date_filter('gaming_date', anio, mes)
    mw_solo_dias, mp_solo_dias = build_date_filter('fecha_operacion', anio, mes)
    dias_totales = (_exec_one(f"""
        SELECT COUNT(DISTINCT fecha) as dias FROM (
            SELECT gaming_date as fecha FROM srw_jugadores {sw_solo}
            UNION
            SELECT fecha_operacion as fecha FROM mesas_puntos {mw_solo_dias}
        ) sub
    """, {**sp_solo, **mp_solo_dias})).get('dias') or 1

    pct_primario, pct_categoria = _get_invitaciones_config()

    cw_inner = cw
    if jefe_filter_sql:
        cw_inner = (cw + jefe_filter_sql) if cw else ("WHERE 1=1" + jefe_filter_sql)

    sw_plain, sp_plain = build_date_filter('gaming_date', anio, mes)
    mw_plain, mp_plain = build_date_filter('fecha_operacion', anio, mes)

    all_params = {**cparam, **jefe_p, **pparam, **mparam, **sp_plain, **mp_plain, **sp}

    jugadores_srw = _exec(f"""
        SELECT
            s.player_id,
            MAX(s.full_name) as nombre,
            MAX(s.player_level) as nivel,
            SUM(s.coin_in) + COALESCE(MAX(m.coin_in_mesas), 0) as coin_in_mensual,
            COALESCE(MAX(d.dias_combinados), COUNT(DISTINCT s.gaming_date)) as dias_asistidos,
            COALESCE(MAX(c.total_cortesias), 0) as total_cortesias,
            COALESCE(MAX(c.monto_micros), 0) as monto_micros,
            COALESCE(MAX(p.cant_premios), 0) as cant_premios,
            COALESCE(MAX(p.monto_premios), 0) as monto_premios
        FROM srw_jugadores s
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_micros
            FROM cortesias c {cw_inner}
            GROUP BY cliente_id
        ) c ON s.player_id = c.cliente_id
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
            FROM premios_comps p {pw}
            GROUP BY cliente_id
        ) p ON s.player_id = p.cliente_id
        LEFT JOIN (
            SELECT cliente_id, SUM(coin_in_puntos) as coin_in_mesas
            FROM mesas_puntos m {mw}
            GROUP BY cliente_id
        ) m ON s.player_id = m.cliente_id
        LEFT JOIN (
            SELECT cliente_id, COUNT(DISTINCT fecha) as dias_combinados FROM (
                SELECT player_id as cliente_id, gaming_date as fecha FROM srw_jugadores {sw_plain}
                UNION
                SELECT cliente_id, fecha_operacion as fecha FROM mesas_puntos {mw_plain}
            ) sub GROUP BY cliente_id
        ) d ON s.player_id = d.cliente_id
        {sw}
        GROUP BY s.player_id
        HAVING COALESCE(MAX(c.total_cortesias), 0) > 0
        ORDER BY coin_in_mensual DESC
    """, all_params)

    # Jugadores solo-mesas (no en SRW)
    sw_excl, sp_excl = build_date_filter('gaming_date', anio, mes)
    mw_inner, mparam_inner = build_date_filter('mp.fecha_operacion', anio, mes)
    mesas_excl = f"mp.cliente_id NOT IN (SELECT DISTINCT player_id FROM srw_jugadores {sw_excl})"
    if mparam_inner:
        mw_conditions = []
        if anio:
            mw_conditions.append(f"SUBSTR(mp.fecha_operacion, 1, 4) = :anio_mp_fecha_operacion")
        if mes:
            mw_conditions.append(f"SUBSTR(mp.fecha_operacion, 6, 2) = :mes_mp_fecha_operacion")
        mesas_where = "WHERE " + " AND ".join(mw_conditions) + " AND " + mesas_excl
    else:
        mesas_where = "WHERE " + mesas_excl

    mesas_params = {**cparam, **jefe_p, **pparam, **mparam_inner, **sp_excl}

    jugadores_mesas = _exec(f"""
        SELECT
            mp.cliente_id as player_id,
            MAX(mp.cliente_nombre) as nombre,
            COALESCE(MAX(mp.categoria), 'Sin Categoria') as nivel,
            SUM(mp.coin_in_puntos) as coin_in_mensual,
            COUNT(DISTINCT mp.fecha_operacion) as dias_asistidos,
            COALESCE(MAX(c.total_cortesias), 0) as total_cortesias,
            COALESCE(MAX(c.monto_micros), 0) as monto_micros,
            COALESCE(MAX(p.cant_premios), 0) as cant_premios,
            COALESCE(MAX(p.monto_premios), 0) as monto_premios
        FROM mesas_puntos mp
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_micros
            FROM cortesias c {cw_inner}
            GROUP BY cliente_id
        ) c ON mp.cliente_id = c.cliente_id
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
            FROM premios_comps p {pw}
            GROUP BY cliente_id
        ) p ON mp.cliente_id = p.cliente_id
        {mesas_where}
        GROUP BY mp.cliente_id
        HAVING COALESCE(MAX(c.total_cortesias), 0) > 0
        ORDER BY coin_in_mensual DESC
    """, mesas_params)

    # Jugadores que solo tienen cortesías (no aparecen en SRW ni mesas_puntos)
    sw_excl2, sp_excl2 = build_date_filter('gaming_date', anio, mes)
    mw_excl2, mp_excl2 = build_date_filter('fecha_operacion', anio, mes)
    ids_ya = set(j['player_id'] for j in jugadores_srw) | set(j['player_id'] for j in jugadores_mesas)
    cort_only_params = {**cparam, **jefe_p, **pparam}
    jugadores_cort_only = _exec(f"""
        SELECT
            c.cliente_id as player_id,
            MAX(c.nombre_cliente) as nombre,
            '-' as nivel,
            0 as coin_in_mensual,
            0 as dias_asistidos,
            COUNT(c.id) as total_cortesias,
            SUM(c.micros) as monto_micros,
            COALESCE(MAX(p.cant_premios), 0) as cant_premios,
            COALESCE(MAX(p.monto_premios), 0) as monto_premios
        FROM cortesias c
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
            FROM premios_comps p {pw}
            GROUP BY cliente_id
        ) p ON c.cliente_id = p.cliente_id
        {cw_inner}
        GROUP BY c.cliente_id
        ORDER BY monto_micros DESC
    """, cort_only_params)
    jugadores_cort_only = [j for j in jugadores_cort_only if j['player_id'] not in ids_ya]

    jugadores = list(jugadores_srw) + list(jugadores_mesas) + list(jugadores_cort_only)
    resultados = _calc_invitaciones(jugadores, pct_primario, pct_categoria, dias_totales)

    # Gráfico de torta
    cw_chart, cp_chart = build_date_filter('c.fecha_jornada', anio, mes)
    chart_params = dict(cp_chart)

    if area:
        extra = " AND j.area = :chart_area" if cw_chart else "WHERE j.area = :chart_area"
        chart_params["chart_area"] = area
        chart_rows = _exec(f"""
            SELECT j.nombre as etiqueta, COUNT(*) as cantidad
            FROM cortesias c
            LEFT JOIN jefaturas j ON c.usuario_id = j.usuario_id
            {cw_chart}{extra}
            GROUP BY j.nombre ORDER BY cantidad DESC
        """, chart_params)
        chart_titulo = f"Cortesías por Jefe — {area}"
    else:
        extra = " AND j.area IS NOT NULL AND j.area != ''" if cw_chart else "WHERE j.area IS NOT NULL AND j.area != ''"
        chart_rows = _exec(f"""
            SELECT j.area as etiqueta, COUNT(*) as cantidad
            FROM cortesias c
            LEFT JOIN jefaturas j ON c.usuario_id = j.usuario_id
            {cw_chart}{extra}
            GROUP BY j.area ORDER BY cantidad DESC
        """, chart_params)
        chart_titulo = "Cortesías por Sección"

    chart_labels = [r['etiqueta'] or 'Sin asignar' for r in chart_rows]
    chart_cantidades = [r['cantidad'] for r in chart_rows]

    # KPIs
    cw_kpi, cp_kpi = build_date_filter('fecha_jornada', anio, mes)
    kpi_cortesias = _exec_one(f"SELECT COALESCE(SUM(micros), 0) as total FROM cortesias {cw_kpi}", cp_kpi)
    sw_kpi, sp_kpi = build_date_filter('gaming_date', anio, mes)
    kpi_coin_srw = _exec_one(f"SELECT COALESCE(SUM(coin_in), 0) as total FROM srw_jugadores {sw_kpi}", sp_kpi)
    mw_kpi, mp_kpi = build_date_filter('fecha_operacion', anio, mes)
    kpi_coin_mesas = _exec_one(f"SELECT COALESCE(SUM(coin_in_puntos), 0) as total FROM mesas_puntos {mw_kpi}", mp_kpi)
    total_cortesias_periodo = kpi_cortesias.get('total') or 0
    total_coin_in_periodo = (kpi_coin_srw.get('total') or 0) + (kpi_coin_mesas.get('total') or 0)
    pct_cortesias_coin_in = round(total_cortesias_periodo * 100.0 / total_coin_in_periodo, 3) if total_coin_in_periodo > 0 else 0

    # Coin-In solo de jugadores CON cortesías
    cw_cc, cp_cc = build_date_filter('fecha_jornada', anio, mes)
    subq_cort = f"SELECT DISTINCT cliente_id FROM cortesias {cw_cc}"
    sw_cc, sp_cc = build_date_filter('gaming_date', anio, mes)
    coin_srw_cc = _exec_one(f"SELECT COALESCE(SUM(coin_in), 0) as total FROM srw_jugadores {sw_cc} {'AND' if sw_cc else 'WHERE'} player_id IN ({subq_cort})", {**sp_cc, **cp_cc})
    mw_cc, mp_cc = build_date_filter('fecha_operacion', anio, mes)
    coin_mesas_cc = _exec_one(f"SELECT COALESCE(SUM(coin_in_puntos), 0) as total FROM mesas_puntos {mw_cc} {'AND' if mw_cc else 'WHERE'} cliente_id IN ({subq_cort})", {**mp_cc, **cp_cc})
    coin_in_con_cortesias = (coin_srw_cc.get('total') or 0) + (coin_mesas_cc.get('total') or 0)

    return render_template('comps/control_invitaciones.html',
                           resultados=resultados,
                           dias_totales=dias_totales,
                           pct_primario=pct_primario,
                           pct_categoria=pct_categoria,
                           chart_labels=chart_labels,
                           chart_cantidades=chart_cantidades,
                           chart_titulo=chart_titulo,
                           total_cortesias_periodo=total_cortesias_periodo,
                           total_coin_in_periodo=total_coin_in_periodo,
                           pct_cortesias_coin_in=pct_cortesias_coin_in,
                           coin_in_con_cortesias=coin_in_con_cortesias,
                           anios=anios, meses_disp=meses_disp,
                           areas=areas, jefes_disp=jefes_disp,
                           anio_actual=anio, mes_actual=mes,
                           area_actual=area, jefe_actual=jefe)


# ─────────────────────────── Control Invitaciones MDA ───────────────────────────

@comps_bp.route('/control/invitaciones-mda')
@login_required
def control_invitaciones_mda():
    anio = request.args.get('anio', '')
    mes = request.args.get('mes', '')
    jefe = request.args.get('jefe', '')
    anios, meses_disp = get_anios_meses()

    jefes_disp = [(r['usuario_id'], r['nombre']) for r in _exec(
        "SELECT usuario_id, nombre FROM jefaturas WHERE area = 'MDA' ORDER BY nombre")]

    jefe_filter_sql = ""
    jefe_p = {}
    if jefe:
        jefe_filter_sql = " AND c.usuario_id = :jefe_id"
        jefe_p = {"jefe_id": jefe}
    else:
        jefe_filter_sql = " AND c.usuario_id IN (SELECT usuario_id FROM jefaturas WHERE area = 'MDA')"

    sw, sp = build_date_filter('s.gaming_date', anio, mes)
    cw, cparam = build_date_filter('c.fecha_jornada', anio, mes)
    pw, pparam = build_date_filter('p.fecha_jornada', anio, mes)

    sw_solo, sp_solo = build_date_filter('gaming_date', anio, mes)
    dias_totales = (_exec_one(f"""
        SELECT COUNT(DISTINCT gaming_date) as dias FROM srw_jugadores {sw_solo}
    """, sp_solo)).get('dias') or 1

    pct_primario, pct_categoria = _get_invitaciones_config()

    cw_inner = (cw + jefe_filter_sql) if cw else ("WHERE 1=1" + jefe_filter_sql)

    all_params = {**cparam, **jefe_p, **pparam, **sp}

    jugadores_raw = _exec(f"""
        SELECT
            s.player_id,
            MAX(s.full_name) as nombre,
            MAX(s.player_level) as nivel,
            SUM(s.coin_in) as coin_in_mensual,
            COUNT(DISTINCT s.gaming_date) as dias_asistidos,
            COALESCE(MAX(c.total_cortesias), 0) as total_cortesias,
            COALESCE(MAX(c.monto_micros), 0) as monto_micros,
            COALESCE(MAX(p.cant_premios), 0) as cant_premios,
            COALESCE(MAX(p.monto_premios), 0) as monto_premios
        FROM srw_jugadores s
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_micros
            FROM cortesias c {cw_inner}
            GROUP BY cliente_id
        ) c ON s.player_id = c.cliente_id
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
            FROM premios_comps p {pw}
            GROUP BY cliente_id
        ) p ON s.player_id = p.cliente_id
        {sw}
        GROUP BY s.player_id
        HAVING COALESCE(MAX(c.total_cortesias), 0) > 0
        ORDER BY coin_in_mensual DESC
    """, all_params)

    resultados = _calc_invitaciones(jugadores_raw, pct_primario, pct_categoria, dias_totales)

    # Jugadores solo-cortesías MDA (no en SRW)
    ids_ya_mda = set(j['player_id'] for j in jugadores_raw)
    cort_only_mda_params = {**cparam, **jefe_p, **pparam}
    jugadores_cort_only_mda = _exec(f"""
        SELECT
            c.cliente_id as player_id,
            MAX(c.nombre_cliente) as nombre,
            '-' as nivel,
            0 as coin_in_mensual,
            0 as dias_asistidos,
            COUNT(c.id) as total_cortesias,
            SUM(c.micros) as monto_micros,
            COALESCE(MAX(p.cant_premios), 0) as cant_premios,
            COALESCE(MAX(p.monto_premios), 0) as monto_premios
        FROM cortesias c
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
            FROM premios_comps p {pw}
            GROUP BY cliente_id
        ) p ON c.cliente_id = p.cliente_id
        {cw_inner}
        GROUP BY c.cliente_id
        ORDER BY monto_micros DESC
    """, cort_only_mda_params)
    jugadores_cort_only_mda = [j for j in jugadores_cort_only_mda if j['player_id'] not in ids_ya_mda]

    all_resultados = list(resultados) + _calc_invitaciones(jugadores_cort_only_mda, pct_primario, pct_categoria, dias_totales)

    # Gráfico de torta MDA
    cw_chart, cp_chart = build_date_filter('c.fecha_jornada', anio, mes)
    chart_params = dict(cp_chart)
    extra_chart = " AND j.area = 'MDA'" if cw_chart else "WHERE j.area = 'MDA'"
    chart_rows = _exec(f"""
        SELECT j.nombre as etiqueta, COUNT(*) as cantidad
        FROM cortesias c
        LEFT JOIN jefaturas j ON c.usuario_id = j.usuario_id
        {cw_chart}{extra_chart}
        GROUP BY j.nombre ORDER BY cantidad DESC
    """, chart_params)
    chart_labels = [r['etiqueta'] or 'Sin asignar' for r in chart_rows]
    chart_cantidades = [r['cantidad'] for r in chart_rows]
    chart_titulo = "Cortesías por Jefe — MDA"

    # KPIs
    cw_kpi, cp_kpi = build_date_filter('fecha_jornada', anio, mes)
    kpi_cortesias = _exec_one(f"SELECT COALESCE(SUM(micros), 0) as total FROM cortesias {cw_kpi}", cp_kpi)
    sw_kpi, sp_kpi = build_date_filter('gaming_date', anio, mes)
    kpi_coin = _exec_one(f"SELECT COALESCE(SUM(coin_in), 0) as total FROM srw_jugadores {sw_kpi}", sp_kpi)
    total_cortesias_periodo = kpi_cortesias.get('total') or 0
    total_coin_in_periodo = kpi_coin.get('total') or 0
    pct_cortesias_coin_in = round(total_cortesias_periodo * 100.0 / total_coin_in_periodo, 3) if total_coin_in_periodo > 0 else 0

    # Coin-In solo de jugadores CON cortesías (MDA)
    cw_cc, cp_cc = build_date_filter('fecha_jornada', anio, mes)
    subq_cort_mda = f"SELECT DISTINCT cliente_id FROM cortesias {cw_cc}"
    sw_cc, sp_cc = build_date_filter('gaming_date', anio, mes)
    coin_srw_cc = _exec_one(f"SELECT COALESCE(SUM(coin_in), 0) as total FROM srw_jugadores {sw_cc} {'AND' if sw_cc else 'WHERE'} player_id IN ({subq_cort_mda})", {**sp_cc, **cp_cc})
    coin_in_con_cortesias = coin_srw_cc.get('total') or 0

    return render_template('comps/control_invitaciones_mda.html',
                           resultados=all_resultados, dias_totales=dias_totales,
                           pct_primario=pct_primario, pct_categoria=pct_categoria,
                           chart_labels=chart_labels,
                           chart_cantidades=chart_cantidades,
                           chart_titulo=chart_titulo,
                           total_cortesias_periodo=total_cortesias_periodo,
                           total_coin_in_periodo=total_coin_in_periodo,
                           pct_cortesias_coin_in=pct_cortesias_coin_in,
                           coin_in_con_cortesias=coin_in_con_cortesias,
                           anios=anios, meses_disp=meses_disp,
                           jefes_disp=jefes_disp,
                           anio_actual=anio, mes_actual=mes, jefe_actual=jefe)


# ─────────────────────────── Control Invitaciones MDJ ───────────────────────────

@comps_bp.route('/control/invitaciones-mdj')
@login_required
def control_invitaciones_mdj():
    anio = request.args.get('anio', '')
    mes = request.args.get('mes', '')
    jefe = request.args.get('jefe', '')
    anios, meses_disp = get_anios_meses()

    jefes_disp = [(r['usuario_id'], r['nombre']) for r in _exec(
        "SELECT usuario_id, nombre FROM jefaturas WHERE area = 'MDJ' ORDER BY nombre")]

    jefe_filter_sql = ""
    jefe_p = {}
    if jefe:
        jefe_filter_sql = " AND c.usuario_id = :jefe_id"
        jefe_p = {"jefe_id": jefe}
    else:
        jefe_filter_sql = " AND c.usuario_id IN (SELECT usuario_id FROM jefaturas WHERE area = 'MDJ')"

    cw, cparam = build_date_filter('c.fecha_jornada', anio, mes)
    pw, pparam = build_date_filter('p.fecha_jornada', anio, mes)
    mw, mparam = build_date_filter('mp.fecha_operacion', anio, mes)

    mw_solo, mp_solo = build_date_filter('fecha_operacion', anio, mes)
    dias_totales = (_exec_one(f"""
        SELECT COUNT(DISTINCT fecha_operacion) as dias FROM mesas_puntos {mw_solo}
    """, mp_solo)).get('dias') or 1

    pct_primario, pct_categoria = _get_invitaciones_config()

    cw_inner = (cw + jefe_filter_sql) if cw else ("WHERE 1=1" + jefe_filter_sql)

    all_params = {**cparam, **jefe_p, **pparam, **mparam}

    jugadores_raw = _exec(f"""
        SELECT
            mp.cliente_id as player_id,
            MAX(mp.cliente_nombre) as nombre,
            COALESCE(MAX(mp.categoria), COALESCE(MAX(s.player_level), 'Sin Categoria')) as nivel,
            SUM(mp.coin_in_puntos) as coin_in_mensual,
            COUNT(DISTINCT mp.fecha_operacion) as dias_asistidos,
            COALESCE(MAX(c.total_cortesias), 0) as total_cortesias,
            COALESCE(MAX(c.monto_micros), 0) as monto_micros,
            COALESCE(MAX(p.cant_premios), 0) as cant_premios,
            COALESCE(MAX(p.monto_premios), 0) as monto_premios
        FROM mesas_puntos mp
        LEFT JOIN (
            SELECT player_id, MAX(player_level) as player_level FROM srw_jugadores GROUP BY player_id
        ) s ON mp.cliente_id = s.player_id
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_micros
            FROM cortesias c {cw_inner}
            GROUP BY cliente_id
        ) c ON mp.cliente_id = c.cliente_id
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
            FROM premios_comps p {pw}
            GROUP BY cliente_id
        ) p ON mp.cliente_id = p.cliente_id
        {mw}
        GROUP BY mp.cliente_id
        HAVING COALESCE(MAX(c.total_cortesias), 0) > 0
        ORDER BY coin_in_mensual DESC
    """, all_params)

    resultados = _calc_invitaciones(jugadores_raw, pct_primario, pct_categoria, dias_totales)

    # Jugadores solo-cortesías MDJ (no en mesas_puntos)
    ids_ya_mdj = set(j['player_id'] for j in jugadores_raw)
    cort_only_mdj_params = {**cparam, **jefe_p, **pparam}
    jugadores_cort_only_mdj = _exec(f"""
        SELECT
            c.cliente_id as player_id,
            MAX(c.nombre_cliente) as nombre,
            '-' as nivel,
            0 as coin_in_mensual,
            0 as dias_asistidos,
            COUNT(c.id) as total_cortesias,
            SUM(c.micros) as monto_micros,
            COALESCE(MAX(p.cant_premios), 0) as cant_premios,
            COALESCE(MAX(p.monto_premios), 0) as monto_premios
        FROM cortesias c
        LEFT JOIN (
            SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
            FROM premios_comps p {pw}
            GROUP BY cliente_id
        ) p ON c.cliente_id = p.cliente_id
        {cw_inner}
        GROUP BY c.cliente_id
        ORDER BY monto_micros DESC
    """, cort_only_mdj_params)
    jugadores_cort_only_mdj = [j for j in jugadores_cort_only_mdj if j['player_id'] not in ids_ya_mdj]

    all_resultados_mdj = list(resultados) + _calc_invitaciones(jugadores_cort_only_mdj, pct_primario, pct_categoria, dias_totales)

    # Gráfico de torta MDJ
    cw_chart, cp_chart = build_date_filter('c.fecha_jornada', anio, mes)
    chart_params = dict(cp_chart)
    extra_chart = " AND j.area = 'MDJ'" if cw_chart else "WHERE j.area = 'MDJ'"
    chart_rows = _exec(f"""
        SELECT j.nombre as etiqueta, COUNT(*) as cantidad
        FROM cortesias c
        LEFT JOIN jefaturas j ON c.usuario_id = j.usuario_id
        {cw_chart}{extra_chart}
        GROUP BY j.nombre ORDER BY cantidad DESC
    """, chart_params)
    chart_labels = [r['etiqueta'] or 'Sin asignar' for r in chart_rows]
    chart_cantidades = [r['cantidad'] for r in chart_rows]
    chart_titulo = "Cortesías por Jefe — MDJ"

    # KPIs
    cw_kpi, cp_kpi = build_date_filter('fecha_jornada', anio, mes)
    kpi_cortesias = _exec_one(f"SELECT COALESCE(SUM(micros), 0) as total FROM cortesias {cw_kpi}", cp_kpi)
    mw_kpi, mp_kpi = build_date_filter('fecha_operacion', anio, mes)
    kpi_coin = _exec_one(f"SELECT COALESCE(SUM(coin_in_puntos), 0) as total FROM mesas_puntos {mw_kpi}", mp_kpi)
    total_cortesias_periodo = kpi_cortesias.get('total') or 0
    total_coin_in_periodo = kpi_coin.get('total') or 0
    pct_cortesias_coin_in = round(total_cortesias_periodo * 100.0 / total_coin_in_periodo, 3) if total_coin_in_periodo > 0 else 0

    # Coin-In solo de jugadores CON cortesías (MDJ)
    cw_cc, cp_cc = build_date_filter('fecha_jornada', anio, mes)
    subq_cort_mdj = f"SELECT DISTINCT cliente_id FROM cortesias {cw_cc}"
    mw_cc, mp_cc = build_date_filter('fecha_operacion', anio, mes)
    coin_mesas_cc = _exec_one(f"SELECT COALESCE(SUM(coin_in_puntos), 0) as total FROM mesas_puntos {mw_cc} {'AND' if mw_cc else 'WHERE'} cliente_id IN ({subq_cort_mdj})", {**mp_cc, **cp_cc})
    coin_in_con_cortesias = coin_mesas_cc.get('total') or 0

    return render_template('comps/control_invitaciones_mdj.html',
                           resultados=all_resultados_mdj, dias_totales=dias_totales,
                           pct_primario=pct_primario, pct_categoria=pct_categoria,
                           chart_labels=chart_labels,
                           chart_cantidades=chart_cantidades,
                           chart_titulo=chart_titulo,
                           total_cortesias_periodo=total_cortesias_periodo,
                           total_coin_in_periodo=total_coin_in_periodo,
                           pct_cortesias_coin_in=pct_cortesias_coin_in,
                           coin_in_con_cortesias=coin_in_con_cortesias,
                           anios=anios, meses_disp=meses_disp,
                           jefes_disp=jefes_disp,
                           anio_actual=anio, mes_actual=mes, jefe_actual=jefe)


# ─────────────────────────── Auditoría Coin-In Cero ───────────────────────────

@comps_bp.route('/auditoria/coinin-cero')
@login_required
def auditoria_coinin_cero():
    anio = request.args.get('anio', '')
    mes = request.args.get('mes', '')
    area = request.args.get('area', '')
    jefe = request.args.get('jefe', '')
    anios, meses_disp = get_anios_meses()

    areas = [r['area'] for r in _exec("SELECT DISTINCT area FROM jefaturas WHERE area != '' ORDER BY area")]

    if area:
        jefes_disp = [(r['usuario_id'], r['nombre']) for r in _exec(
            "SELECT usuario_id, nombre FROM jefaturas WHERE area = :area ORDER BY nombre", {"area": area})]
    else:
        jefes_disp = [(r['usuario_id'], r['nombre']) for r in _exec(
            "SELECT usuario_id, nombre FROM jefaturas ORDER BY nombre")]

    cw, cp = build_date_filter('c.fecha_jornada', anio, mes)

    extra_conditions = []
    extra_params = {}
    if jefe:
        extra_conditions.append("c.usuario_id = :jefe_id")
        extra_params["jefe_id"] = jefe
    elif area:
        extra_conditions.append("c.usuario_id IN (SELECT usuario_id FROM jefaturas WHERE area = :area_filter)")
        extra_params["area_filter"] = area

    where_parts = []
    all_params = {}
    if cw:
        where_parts.append(cw.replace("WHERE ", ""))
        all_params.update(cp)
    where_parts.append("""(
        NOT EXISTS (
            SELECT 1 FROM srw_jugadores s
            WHERE s.player_id = c.cliente_id
              AND s.gaming_date = c.fecha_jornada
              AND s.coin_in > 0
        )
        AND NOT EXISTS (
            SELECT 1 FROM mesas_puntos m
            WHERE m.cliente_id = c.cliente_id
              AND m.fecha_operacion = c.fecha_jornada
              AND m.coin_in_puntos > 0
        )
    )""")
    if extra_conditions:
        where_parts.extend(extra_conditions)
        all_params.update(extra_params)

    where_clause = "WHERE " + " AND ".join(where_parts)

    resultados = _exec(f"""
        SELECT
            c.fecha_jornada as jornada,
            c.cliente_id,
            MAX(c.nombre_cliente) as nombre_cliente,
            COALESCE(MAX(s.coin_in_dia), 0) + COALESCE(MAX(me.coin_in_mesas), 0) as coin_in,
            COUNT(c.id) as cant_cortesias,
            SUM(c.micros) as monto_cortesias,
            COALESCE(MAX(p.cant_premios), 0) as cant_premios,
            COALESCE(MAX(p.monto_premios), 0) as monto_premios,
            COALESCE(MAX(j.nombre), '') as jefe_nombre,
            COALESCE(MAX(j.area), '') as jefe_area
        FROM cortesias c
        LEFT JOIN (
            SELECT player_id, gaming_date, SUM(coin_in) as coin_in_dia
            FROM srw_jugadores GROUP BY player_id, gaming_date
        ) s ON c.cliente_id = s.player_id AND c.fecha_jornada = s.gaming_date
        LEFT JOIN (
            SELECT cliente_id, fecha_operacion, SUM(coin_in_puntos) as coin_in_mesas
            FROM mesas_puntos GROUP BY cliente_id, fecha_operacion
        ) me ON c.cliente_id = me.cliente_id AND c.fecha_jornada = me.fecha_operacion
        LEFT JOIN (
            SELECT cliente_id, fecha_jornada, COUNT(*) as cant_premios,
                   SUM(transferencia_final) as monto_premios
            FROM premios_comps GROUP BY cliente_id, fecha_jornada
        ) p ON c.cliente_id = p.cliente_id AND c.fecha_jornada = p.fecha_jornada
        LEFT JOIN jefaturas j ON c.usuario_id = j.usuario_id
        {where_clause}
        GROUP BY c.fecha_jornada, c.cliente_id, c.usuario_id
        ORDER BY c.fecha_jornada DESC, monto_cortesias DESC
    """, all_params)

    # Gráfico
    cw_chart, cp_chart = build_date_filter('c.fecha_jornada', anio, mes)
    chart_params = dict(cp_chart)
    coin_zero_cond = """(
        NOT EXISTS (SELECT 1 FROM srw_jugadores s WHERE s.player_id = c.cliente_id AND s.gaming_date = c.fecha_jornada AND s.coin_in > 0)
        AND NOT EXISTS (SELECT 1 FROM mesas_puntos m WHERE m.cliente_id = c.cliente_id AND m.fecha_operacion = c.fecha_jornada AND m.coin_in_puntos > 0)
    )"""
    if cw_chart:
        chart_where = cw_chart.replace("WHERE ", "WHERE " + coin_zero_cond + " AND ")
    else:
        chart_where = "WHERE " + coin_zero_cond

    if area:
        chart_where += " AND j.area = :chart_area"
        chart_params["chart_area"] = area
        chart_rows = _exec(f"""
            SELECT j.nombre as etiqueta, COUNT(*) as cantidad
            FROM cortesias c LEFT JOIN jefaturas j ON c.usuario_id = j.usuario_id
            {chart_where} GROUP BY j.nombre ORDER BY cantidad DESC
        """, chart_params)
        chart_titulo = f"Casos Coin In Cero por Jefe — {area}"
    else:
        chart_where += " AND j.area IS NOT NULL AND j.area != ''"
        chart_rows = _exec(f"""
            SELECT j.area as etiqueta, COUNT(*) as cantidad
            FROM cortesias c LEFT JOIN jefaturas j ON c.usuario_id = j.usuario_id
            {chart_where} GROUP BY j.area ORDER BY cantidad DESC
        """, chart_params)
        chart_titulo = "Casos Coin In Cero por Sección"

    chart_labels = [r['etiqueta'] or 'Sin asignar' for r in chart_rows]
    chart_cantidades = [r['cantidad'] for r in chart_rows]

    # Detalle de productos por cliente/fecha (para expandir filas)
    productos_detalle = _exec(f"""
        SELECT c.cliente_id, c.fecha_jornada,
               c.descripcion_cat, c.descripcion_prod,
               COUNT(*) as cantidad, SUM(c.micros) as monto
        FROM cortesias c
        {where_clause}
        GROUP BY c.cliente_id, c.fecha_jornada, c.descripcion_cat, c.descripcion_prod
        ORDER BY c.descripcion_cat, monto DESC
    """, all_params)
    prods_por_caso = {}
    for r in productos_detalle:
        key = f"{r['cliente_id']}|{r['fecha_jornada']}"
        prods_por_caso.setdefault(key, []).append(r)

    return render_template('comps/auditoria_coinin_cero.html',
                           resultados=resultados,
                           prods_por_caso=prods_por_caso,
                           chart_labels=chart_labels,
                           chart_cantidades=chart_cantidades,
                           chart_titulo=chart_titulo,
                           anios=anios, meses_disp=meses_disp,
                           areas=areas, jefes_disp=jefes_disp,
                           anio_actual=anio, mes_actual=mes,
                           area_actual=area, jefe_actual=jefe)


# ─────────────────────────── Exportar ───────────────────────────

@comps_bp.route('/exportar')
@login_required
def exportar_reportes():
    anios, meses_disp = get_anios_meses()
    return render_template('comps/exportar.html', anios=anios, meses_disp=meses_disp)


def _autosize_cols(writer, sheet_name):
    """Ajusta el ancho de columnas en una hoja de Excel."""
    ws = writer.sheets[sheet_name]
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 45)


def _write_sheet(writer, rows, sheet_name, columns=None):
    """Escribe datos en una hoja de Excel y ajusta columnas."""
    df = pd.DataFrame(rows)
    if df.empty:
        return
    if columns:
        df.columns = columns
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    _autosize_cols(writer, sheet_name)


@comps_bp.route('/exportar/generar', methods=['POST'])
@login_required
def exportar_generar():
    anio = request.form.get('anio', '')
    mes = request.form.get('mes', '')
    secciones = request.form.getlist('secciones')

    if not secciones:
        flash('Selecciona al menos una sección.', 'error')
        return redirect(url_for('comps.exportar_reportes'))

    output = BytesIO()
    periodo = f"{anio or 'Todos'}-{MESES_NOMBRE.get(mes, mes) if mes else 'Todos'}"
    dias_semana = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']

    with pd.ExcelWriter(output, engine='openpyxl') as writer:

        # ── CORTESÍAS ──────────────────────────────────────────────
        if 'cortesias' in secciones:
            cw, cp = build_date_filter('c.fecha_jornada', anio, mes)
            cw_solo, cp_solo = build_date_filter('fecha_jornada', anio, mes)
            sw, sp = build_date_filter('gaming_date', anio, mes)
            mw_exp, mp_exp = build_date_filter('m.fecha_operacion', anio, mes)

            # Resumen por jugador
            rows = _exec(f"""
                SELECT c.cliente_id, c.nombre_cliente,
                       COALESCE(MAX(s.player_level), '-') as player_level,
                       COUNT(c.id) as total_cortesias, SUM(c.micros) as monto_cortesias,
                       COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0) as total_coin_in,
                       COALESCE(MAX(s.total_promo_in), 0) as total_promo_in,
                       COALESCE(MAX(s.total_games), 0) as total_games,
                       CASE WHEN (COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0)) > 0
                            THEN ROUND((SUM(c.micros) * 100.0 / (COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0)))::numeric, 4)
                            ELSE 0 END as pct_cortesia_coin_in
                FROM cortesias c
                LEFT JOIN (
                    SELECT player_id, SUM(coin_in) as total_coin_in, SUM(promo_in) as total_promo_in,
                           SUM(total_games) as total_games, MAX(player_level) as player_level
                    FROM srw_jugadores {sw} GROUP BY player_id
                ) s ON c.cliente_id = s.player_id
                LEFT JOIN (
                    SELECT cliente_id, SUM(coin_in_puntos) as coin_in_mesas
                    FROM mesas_puntos m {mw_exp} GROUP BY cliente_id
                ) m ON c.cliente_id = m.cliente_id
                {cw}
                GROUP BY c.cliente_id, c.nombre_cliente
                ORDER BY monto_cortesias DESC
            """, {**sp, **mp_exp, **cp})
            _write_sheet(writer, rows, 'Cort - Resumen',
                         ['ID', 'Nombre', 'Nivel', 'Cortesías', 'Monto Cortesías',
                          'Coin-In', 'Promo-In', 'Juegos', '% Cortesía/Coin-In'])

            # Por categoría
            rows_cat = _exec(f"""
                SELECT descripcion_cat, COUNT(*) as cantidad, SUM(micros) as monto_total
                FROM cortesias {cw_solo}
                GROUP BY descripcion_cat ORDER BY monto_total DESC
            """, cp_solo)
            _write_sheet(writer, rows_cat, 'Cort - Categorías',
                         ['Categoría', 'Cantidad', 'Monto Total'])

            # Por producto
            rows_prod = _exec(f"""
                SELECT descripcion_cat, descripcion_prod, COUNT(*) as cantidad,
                       SUM(micros) as monto_total
                FROM cortesias {cw_solo}
                GROUP BY descripcion_cat, descripcion_prod
                ORDER BY descripcion_cat, monto_total DESC
            """, cp_solo)
            _write_sheet(writer, rows_prod, 'Cort - Productos',
                         ['Categoría', 'Producto', 'Cantidad', 'Monto Total'])

            # Por día
            dia_where, dia_params = build_date_filter('fecha_jornada', anio, mes)
            if dia_where:
                dia_where += " AND fecha_jornada IS NOT NULL"
            else:
                dia_where = "WHERE fecha_jornada IS NOT NULL"
            por_dia_raw = _exec(f"""
                SELECT fecha_jornada, COUNT(*) as cantidad, SUM(micros) as monto_total
                FROM cortesias {dia_where}
                GROUP BY fecha_jornada ORDER BY fecha_jornada
            """, dia_params)

            sw_dia, sp_dia = build_date_filter('gaming_date', anio, mes)
            coin_mda_dia = _exec(f"""
                SELECT gaming_date as fecha, SUM(coin_in) as coin_in
                FROM srw_jugadores {sw_dia}
                {'AND' if sw_dia else 'WHERE'} gaming_date IS NOT NULL
                GROUP BY gaming_date
            """, sp_dia)
            mda_map = {r['fecha']: r['coin_in'] or 0 for r in coin_mda_dia}

            mw_dia, mp_dia = build_date_filter('fecha_operacion', anio, mes)
            coin_mdj_dia = _exec(f"""
                SELECT fecha_operacion as fecha, SUM(coin_in_puntos) as coin_in
                FROM mesas_puntos {mw_dia}
                {'AND' if mw_dia else 'WHERE'} fecha_operacion IS NOT NULL
                GROUP BY fecha_operacion
            """, mp_dia)
            mdj_map = {r['fecha']: r['coin_in'] or 0 for r in coin_mdj_dia}

            por_dia = []
            for d in por_dia_raw:
                fecha_str = str(d['fecha_jornada'])
                try:
                    dia_sem = dias_semana[datetime.strptime(fecha_str, '%Y-%m-%d').weekday()]
                except Exception:
                    dia_sem = ''
                por_dia.append({
                    'fecha': fecha_str, 'dia': dia_sem,
                    'cantidad': d['cantidad'], 'monto': d['monto_total'],
                    'coin_in_mda': mda_map.get(fecha_str, 0),
                    'coin_in_mdj': mdj_map.get(fecha_str, 0),
                })
            _write_sheet(writer, por_dia, 'Cort - Por Día',
                         ['Fecha', 'Día', 'Cantidad', 'Monto Total', 'Coin-In MDA', 'Coin-In MDJ'])

        # ── PREMIOS ────────────────────────────────────────────────
        if 'premios' in secciones:
            pw, pp = build_date_filter('p.fecha_jornada', anio, mes)
            pw_solo, pp_solo = build_date_filter('fecha_jornada', anio, mes)
            sw, sp = build_date_filter('gaming_date', anio, mes)
            mw_exp, mp_exp = build_date_filter('m.fecha_operacion', anio, mes)

            # Por jugador
            rows = _exec(f"""
                SELECT p.cliente_id,
                       COALESCE(MAX(s.full_name), MAX(m.cliente_nombre), '(Sin nombre)') as nombre,
                       COALESCE(MAX(s.player_level), '-') as player_level,
                       COUNT(p.id) as total_premios,
                       SUM(p.transferencia_final) as monto_total,
                       COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0) as total_coin_in,
                       COALESCE(MAX(s.total_promo_in), 0) as total_promo_in,
                       COALESCE(MAX(s.total_games), 0) as total_games,
                       CASE WHEN (COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0)) > 0
                            THEN ROUND((SUM(p.transferencia_final) * 100.0 / (COALESCE(MAX(s.total_coin_in), 0) + COALESCE(MAX(m.coin_in_mesas), 0)))::numeric, 4)
                            ELSE 0 END as pct_premio_coin_in
                FROM premios_comps p
                LEFT JOIN (
                    SELECT player_id, MAX(full_name) as full_name,
                           MAX(player_level) as player_level,
                           SUM(coin_in) as total_coin_in,
                           SUM(promo_in) as total_promo_in,
                           SUM(total_games) as total_games
                    FROM srw_jugadores {sw} GROUP BY player_id
                ) s ON p.cliente_id = s.player_id
                LEFT JOIN (
                    SELECT cliente_id, MAX(cliente_nombre) as cliente_nombre,
                           SUM(coin_in_puntos) as coin_in_mesas
                    FROM mesas_puntos m {mw_exp} GROUP BY cliente_id
                ) m ON p.cliente_id = m.cliente_id
                {pw}
                GROUP BY p.cliente_id
                ORDER BY monto_total DESC
            """, {**sp, **mp_exp, **pp})
            _write_sheet(writer, rows, 'Premios - Jugadores',
                         ['ID', 'Nombre', 'Nivel', 'Total Premios', 'Monto Premios',
                          'Coin-In', 'Promo-In', 'Juegos', '% Premio/Coin-In'])

            # Por tipo de pago
            rows_tipo = _exec(f"""
                SELECT tipo_pago, COUNT(*) as cantidad, SUM(transferencia_final) as monto_total
                FROM premios_comps {pw_solo}
                GROUP BY tipo_pago ORDER BY monto_total DESC
            """, pp_solo)
            _write_sheet(writer, rows_tipo, 'Premios - Tipo Pago',
                         ['Tipo Pago', 'Cantidad', 'Monto Total'])

            # Por día
            dia_where, dia_params = build_date_filter('fecha_jornada', anio, mes)
            if dia_where:
                dia_where += " AND fecha_jornada IS NOT NULL"
            else:
                dia_where = "WHERE fecha_jornada IS NOT NULL"
            rows_dia = _exec(f"""
                SELECT fecha_jornada, COUNT(*) as cantidad, SUM(transferencia_final) as monto_total
                FROM premios_comps {dia_where}
                GROUP BY fecha_jornada ORDER BY fecha_jornada
            """, dia_params)
            _write_sheet(writer, rows_dia, 'Premios - Por Día',
                         ['Fecha', 'Cantidad', 'Monto Total'])

        # ── RESUMEN GENERAL ────────────────────────────────────────
        if 'resumen' in secciones:
            sw, sp = build_date_filter('gaming_date', anio, mes)
            cw, cparam = build_date_filter('fecha_jornada', anio, mes)
            pw, pparam = build_date_filter('fecha_jornada', anio, mes)
            mw_res, mp_res = build_date_filter('m.fecha_operacion', anio, mes)
            mw_res_solo, mp_res_solo = build_date_filter('fecha_operacion', anio, mes)

            # KPIs
            kpis_srw = _exec_one(f"""
                SELECT COALESCE(SUM(coin_in), 0) as total_coin_in,
                       COALESCE(SUM(promo_in), 0) as total_promo_in,
                       COALESCE(SUM(total_games), 0) as total_games,
                       COUNT(DISTINCT player_id) as jugadores_srw
                FROM srw_jugadores {sw}
            """, sp)
            kpis_mesas = _exec_one(f"""
                SELECT COALESCE(SUM(coin_in_puntos), 0) as total_coin_in_mesas
                FROM mesas_puntos {mw_res_solo}
            """, mp_res_solo)
            kpis_cort = _exec_one(f"""
                SELECT COALESCE(SUM(micros), 0) as total_cortesias,
                       COUNT(DISTINCT cliente_id) as clientes_cortesias
                FROM cortesias {cw}
            """, cparam)
            kpis_prem = _exec_one(f"""
                SELECT COALESCE(SUM(transferencia_final), 0) as total_premios,
                       COUNT(DISTINCT cliente_id) as clientes_premios
                FROM premios_comps {pw}
            """, pparam)

            total_coin_in_global = (kpis_srw.get('total_coin_in') or 0) + (kpis_mesas.get('total_coin_in_mesas') or 0)
            kpis_data = [
                {'Indicador': 'Total Coin-In', 'Valor': total_coin_in_global},
                {'Indicador': 'Total Promo-In', 'Valor': kpis_srw.get('total_promo_in') or 0},
                {'Indicador': 'Total Juegos', 'Valor': kpis_srw.get('total_games') or 0},
                {'Indicador': 'Jugadores SRW', 'Valor': kpis_srw.get('jugadores_srw') or 0},
                {'Indicador': 'Total Cortesías ($)', 'Valor': kpis_cort.get('total_cortesias') or 0},
                {'Indicador': 'Clientes con Cortesías', 'Valor': kpis_cort.get('clientes_cortesias') or 0},
                {'Indicador': 'Total Premios ($)', 'Valor': kpis_prem.get('total_premios') or 0},
                {'Indicador': 'Clientes con Premios', 'Valor': kpis_prem.get('clientes_premios') or 0},
            ]
            _write_sheet(writer, kpis_data, 'Resumen - KPIs')

            # Jugadores
            all_params = {**sp, **cparam, **pparam, **mp_res}
            jugadores_raw = _exec(f"""
                SELECT s.player_id, s.full_name, s.player_level,
                       s.total_coin_in + COALESCE(m.coin_in_mesas, 0) as total_coin_in,
                       s.total_promo_in, s.total_games, s.dias_jugados,
                       COALESCE(c.total_cortesias, 0) as total_cortesias,
                       COALESCE(c.monto_cortesias, 0) as monto_cortesias,
                       COALESCE(p.total_premios, 0) as total_premios,
                       COALESCE(p.monto_premios, 0) as monto_premios
                FROM (
                    SELECT player_id, MAX(full_name) as full_name, MAX(player_level) as player_level,
                           SUM(coin_in) as total_coin_in, SUM(promo_in) as total_promo_in,
                           SUM(total_games) as total_games, COUNT(DISTINCT gaming_date) as dias_jugados
                    FROM srw_jugadores {sw} GROUP BY player_id
                ) s
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_cortesias
                    FROM cortesias {cw} GROUP BY cliente_id
                ) c ON s.player_id = c.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as total_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps {pw} GROUP BY cliente_id
                ) p ON s.player_id = p.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, SUM(coin_in_puntos) as coin_in_mesas
                    FROM mesas_puntos m {mw_res} GROUP BY cliente_id
                ) m ON s.player_id = m.cliente_id
                WHERE COALESCE(c.total_cortesias, 0) > 0 OR COALESCE(p.total_premios, 0) > 0
                ORDER BY total_coin_in DESC
            """, all_params)

            jugadores = []
            for j in jugadores_raw:
                row = dict(j)
                row['pct_total_coin_in'] = round((row['total_coin_in'] or 0) * 100.0 / total_coin_in_global, 3) if total_coin_in_global > 0 else 0
                jugadores.append({
                    'id': row['player_id'], 'nombre': row['full_name'], 'nivel': row['player_level'],
                    'coin_in': row['total_coin_in'], 'pct_coin_in': row['pct_total_coin_in'],
                    'promo_in': row['total_promo_in'], 'juegos': row['total_games'],
                    'dias': row['dias_jugados'],
                    'cortesias': row['total_cortesias'], 'monto_cort': row['monto_cortesias'],
                    'premios': row['total_premios'], 'monto_prem': row['monto_premios'],
                })
            _write_sheet(writer, jugadores, 'Resumen - Jugadores',
                         ['ID', 'Nombre', 'Nivel', 'Coin-In', '% Total Coin-In',
                          'Promo-In', 'Juegos', 'Días', 'Cortesías', 'Monto Cortesías',
                          'Premios', 'Monto Premios'])

        # ── CONTROL INVITACIONES — GENERAL ─────────────────────────
        if 'control_general' in secciones:
            sw, sp = build_date_filter('s.gaming_date', anio, mes)
            cw, cparam = build_date_filter('c.fecha_jornada', anio, mes)
            pw, pparam = build_date_filter('p.fecha_jornada', anio, mes)
            mw, mparam = build_date_filter('m.fecha_operacion', anio, mes)
            sw_solo, sp_solo = build_date_filter('gaming_date', anio, mes)
            mw_solo_dias, mp_solo_dias = build_date_filter('fecha_operacion', anio, mes)
            sw_plain, sp_plain = build_date_filter('gaming_date', anio, mes)
            mw_plain, mp_plain = build_date_filter('fecha_operacion', anio, mes)

            dias_totales = (_exec_one(f"""
                SELECT COUNT(DISTINCT fecha) as dias FROM (
                    SELECT gaming_date as fecha FROM srw_jugadores {sw_solo}
                    UNION
                    SELECT fecha_operacion as fecha FROM mesas_puntos {mw_solo_dias}
                ) sub
            """, {**sp_solo, **mp_solo_dias})).get('dias') or 1

            pct_primario, pct_categoria = _get_invitaciones_config()
            all_params = {**cparam, **pparam, **mparam, **sp_plain, **mp_plain, **sp}

            jugadores_srw = _exec(f"""
                SELECT s.player_id,
                       MAX(s.full_name) as nombre, MAX(s.player_level) as nivel,
                       SUM(s.coin_in) + COALESCE(MAX(m.coin_in_mesas), 0) as coin_in_mensual,
                       COALESCE(MAX(d.dias_combinados), COUNT(DISTINCT s.gaming_date)) as dias_asistidos,
                       COALESCE(MAX(c.total_cortesias), 0) as total_cortesias,
                       COALESCE(MAX(c.monto_micros), 0) as monto_micros,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios
                FROM srw_jugadores s
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_micros
                    FROM cortesias c {cw} GROUP BY cliente_id
                ) c ON s.player_id = c.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps p {pw} GROUP BY cliente_id
                ) p ON s.player_id = p.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, SUM(coin_in_puntos) as coin_in_mesas
                    FROM mesas_puntos m {mw} GROUP BY cliente_id
                ) m ON s.player_id = m.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, COUNT(DISTINCT fecha) as dias_combinados FROM (
                        SELECT player_id as cliente_id, gaming_date as fecha FROM srw_jugadores {sw_plain}
                        UNION
                        SELECT cliente_id, fecha_operacion as fecha FROM mesas_puntos {mw_plain}
                    ) sub GROUP BY cliente_id
                ) d ON s.player_id = d.cliente_id
                {sw}
                GROUP BY s.player_id
                HAVING COALESCE(MAX(c.total_cortesias), 0) > 0
                ORDER BY coin_in_mensual DESC
            """, all_params)

            # Mesas-only
            sw_excl, sp_excl = build_date_filter('gaming_date', anio, mes)
            mw_inner, mparam_inner = build_date_filter('mp.fecha_operacion', anio, mes)
            mesas_excl = f"mp.cliente_id NOT IN (SELECT DISTINCT player_id FROM srw_jugadores {sw_excl})"
            if mparam_inner:
                mw_conds = []
                if anio:
                    mw_conds.append(f"SUBSTR(mp.fecha_operacion, 1, 4) = :anio_mp_fecha_operacion")
                if mes:
                    mw_conds.append(f"SUBSTR(mp.fecha_operacion, 6, 2) = :mes_mp_fecha_operacion")
                mesas_where = "WHERE " + " AND ".join(mw_conds) + " AND " + mesas_excl
            else:
                mesas_where = "WHERE " + mesas_excl
            mesas_params = {**cparam, **pparam, **mparam_inner, **sp_excl}

            jugadores_mesas = _exec(f"""
                SELECT mp.cliente_id as player_id,
                       MAX(mp.cliente_nombre) as nombre, COALESCE(MAX(mp.categoria), 'Sin Categoria') as nivel,
                       SUM(mp.coin_in_puntos) as coin_in_mensual,
                       COUNT(DISTINCT mp.fecha_operacion) as dias_asistidos,
                       COALESCE(MAX(c.total_cortesias), 0) as total_cortesias,
                       COALESCE(MAX(c.monto_micros), 0) as monto_micros,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios
                FROM mesas_puntos mp
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_micros
                    FROM cortesias c {cw} GROUP BY cliente_id
                ) c ON mp.cliente_id = c.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps p {pw} GROUP BY cliente_id
                ) p ON mp.cliente_id = p.cliente_id
                {mesas_where}
                GROUP BY mp.cliente_id
                HAVING COALESCE(MAX(c.total_cortesias), 0) > 0
                ORDER BY coin_in_mensual DESC
            """, mesas_params)

            # Cort-only
            ids_ya = set(j['player_id'] for j in jugadores_srw) | set(j['player_id'] for j in jugadores_mesas)
            cort_only = _exec(f"""
                SELECT c.cliente_id as player_id, MAX(c.nombre_cliente) as nombre,
                       '-' as nivel, 0 as coin_in_mensual, 0 as dias_asistidos,
                       COUNT(c.id) as total_cortesias, SUM(c.micros) as monto_micros,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios
                FROM cortesias c
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps p {pw} GROUP BY cliente_id
                ) p ON c.cliente_id = p.cliente_id
                {cw}
                GROUP BY c.cliente_id ORDER BY monto_micros DESC
            """, {**cparam, **pparam})
            cort_only = [j for j in cort_only if j['player_id'] not in ids_ya]

            all_jug = list(jugadores_srw) + list(jugadores_mesas) + cort_only
            resultados = _calc_invitaciones(all_jug, pct_primario, pct_categoria, dias_totales)
            _write_sheet(writer, resultados, 'Ctrl - General',
                         ['Nombre', 'Nivel', 'Días Asist.', '% Asist.', 'Premios', 'Monto Premios',
                          'Coin-In Mensual', 'Cortesías', 'Monto Cortesías',
                          'Invitación Mensual', 'Saldo', '% Cat.'])

        # ── CONTROL INVITACIONES — MDA ─────────────────────────────
        if 'control_mda' in secciones:
            sw, sp = build_date_filter('s.gaming_date', anio, mes)
            cw, cparam = build_date_filter('c.fecha_jornada', anio, mes)
            pw, pparam = build_date_filter('p.fecha_jornada', anio, mes)
            sw_solo, sp_solo = build_date_filter('gaming_date', anio, mes)

            jefe_filter_mda = " AND c.usuario_id IN (SELECT usuario_id FROM jefaturas WHERE area = 'MDA')"
            cw_inner = (cw + jefe_filter_mda) if cw else ("WHERE 1=1" + jefe_filter_mda)

            dias_totales = (_exec_one(f"""
                SELECT COUNT(DISTINCT gaming_date) as dias FROM srw_jugadores {sw_solo}
            """, sp_solo)).get('dias') or 1

            pct_primario, pct_categoria = _get_invitaciones_config()
            all_params = {**cparam, **pparam, **sp}

            jugadores_raw = _exec(f"""
                SELECT s.player_id, MAX(s.full_name) as nombre,
                       MAX(s.player_level) as nivel,
                       SUM(s.coin_in) as coin_in_mensual,
                       COUNT(DISTINCT s.gaming_date) as dias_asistidos,
                       COALESCE(MAX(c.total_cortesias), 0) as total_cortesias,
                       COALESCE(MAX(c.monto_micros), 0) as monto_micros,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios
                FROM srw_jugadores s
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_micros
                    FROM cortesias c {cw_inner} GROUP BY cliente_id
                ) c ON s.player_id = c.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps p {pw} GROUP BY cliente_id
                ) p ON s.player_id = p.cliente_id
                {sw}
                GROUP BY s.player_id
                HAVING COALESCE(MAX(c.total_cortesias), 0) > 0
                ORDER BY coin_in_mensual DESC
            """, all_params)

            resultados_mda = _calc_invitaciones(jugadores_raw, pct_primario, pct_categoria, dias_totales)

            ids_ya_mda = set(j['player_id'] for j in jugadores_raw)
            cort_only_mda = _exec(f"""
                SELECT c.cliente_id as player_id, MAX(c.nombre_cliente) as nombre,
                       '-' as nivel, 0 as coin_in_mensual, 0 as dias_asistidos,
                       COUNT(c.id) as total_cortesias, SUM(c.micros) as monto_micros,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios
                FROM cortesias c
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps p {pw} GROUP BY cliente_id
                ) p ON c.cliente_id = p.cliente_id
                {cw_inner}
                GROUP BY c.cliente_id ORDER BY monto_micros DESC
            """, {**cparam, **pparam})
            cort_only_mda = [j for j in cort_only_mda if j['player_id'] not in ids_ya_mda]

            all_mda = list(resultados_mda) + _calc_invitaciones(cort_only_mda, pct_primario, pct_categoria, dias_totales)
            _write_sheet(writer, all_mda, 'Ctrl - MDA',
                         ['Nombre', 'Nivel', 'Días Asist.', '% Asist.', 'Premios', 'Monto Premios',
                          'Coin-In Mensual', 'Cortesías', 'Monto Cortesías',
                          'Invitación Mensual', 'Saldo', '% Cat.'])

        # ── CONTROL INVITACIONES — MDJ ─────────────────────────────
        if 'control_mdj' in secciones:
            cw, cparam = build_date_filter('c.fecha_jornada', anio, mes)
            pw, pparam = build_date_filter('p.fecha_jornada', anio, mes)
            mw, mparam = build_date_filter('mp.fecha_operacion', anio, mes)
            mw_solo, mp_solo = build_date_filter('fecha_operacion', anio, mes)

            jefe_filter_mdj = " AND c.usuario_id IN (SELECT usuario_id FROM jefaturas WHERE area = 'MDJ')"
            cw_inner = (cw + jefe_filter_mdj) if cw else ("WHERE 1=1" + jefe_filter_mdj)

            dias_totales = (_exec_one(f"""
                SELECT COUNT(DISTINCT fecha_operacion) as dias FROM mesas_puntos {mw_solo}
            """, mp_solo)).get('dias') or 1

            pct_primario, pct_categoria = _get_invitaciones_config()
            all_params = {**cparam, **pparam, **mparam}

            jugadores_raw = _exec(f"""
                SELECT mp.cliente_id as player_id, MAX(mp.cliente_nombre) as nombre,
                       COALESCE(MAX(mp.categoria), COALESCE(MAX(s.player_level), 'Sin Categoria')) as nivel,
                       SUM(mp.coin_in_puntos) as coin_in_mensual,
                       COUNT(DISTINCT mp.fecha_operacion) as dias_asistidos,
                       COALESCE(MAX(c.total_cortesias), 0) as total_cortesias,
                       COALESCE(MAX(c.monto_micros), 0) as monto_micros,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios
                FROM mesas_puntos mp
                LEFT JOIN (
                    SELECT player_id, MAX(player_level) as player_level FROM srw_jugadores GROUP BY player_id
                ) s ON mp.cliente_id = s.player_id
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_micros
                    FROM cortesias c {cw_inner} GROUP BY cliente_id
                ) c ON mp.cliente_id = c.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps p {pw} GROUP BY cliente_id
                ) p ON mp.cliente_id = p.cliente_id
                {mw}
                GROUP BY mp.cliente_id
                HAVING COALESCE(MAX(c.total_cortesias), 0) > 0
                ORDER BY coin_in_mensual DESC
            """, all_params)

            resultados_mdj = _calc_invitaciones(jugadores_raw, pct_primario, pct_categoria, dias_totales)

            ids_ya_mdj = set(j['player_id'] for j in jugadores_raw)
            cort_only_mdj = _exec(f"""
                SELECT c.cliente_id as player_id, MAX(c.nombre_cliente) as nombre,
                       '-' as nivel, 0 as coin_in_mensual, 0 as dias_asistidos,
                       COUNT(c.id) as total_cortesias, SUM(c.micros) as monto_micros,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios
                FROM cortesias c
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps p {pw} GROUP BY cliente_id
                ) p ON c.cliente_id = p.cliente_id
                {cw_inner}
                GROUP BY c.cliente_id ORDER BY monto_micros DESC
            """, {**cparam, **pparam})
            cort_only_mdj = [j for j in cort_only_mdj if j['player_id'] not in ids_ya_mdj]

            all_mdj = list(resultados_mdj) + _calc_invitaciones(cort_only_mdj, pct_primario, pct_categoria, dias_totales)
            _write_sheet(writer, all_mdj, 'Ctrl - MDJ',
                         ['Nombre', 'Nivel', 'Días Asist.', '% Asist.', 'Premios', 'Monto Premios',
                          'Coin-In Mensual', 'Cortesías', 'Monto Cortesías',
                          'Invitación Mensual', 'Saldo', '% Cat.'])

        # ── CONTROL INVITACIONES — MRK ─────────────────────────────
        if 'control_mrk' in secciones:
            sw, sp = build_date_filter('s.gaming_date', anio, mes)
            cw, cparam = build_date_filter('c.fecha_jornada', anio, mes)
            pw, pparam = build_date_filter('p.fecha_jornada', anio, mes)
            mw, mparam = build_date_filter('m.fecha_operacion', anio, mes)
            sw_solo, sp_solo = build_date_filter('gaming_date', anio, mes)
            mw_solo_dias, mp_solo_dias = build_date_filter('fecha_operacion', anio, mes)
            sw_plain, sp_plain = build_date_filter('gaming_date', anio, mes)
            mw_plain, mp_plain = build_date_filter('fecha_operacion', anio, mes)

            jefe_filter_mrk = " AND c.usuario_id IN (SELECT usuario_id FROM jefaturas WHERE area = 'MRK')"
            cw_inner = (cw + jefe_filter_mrk) if cw else ("WHERE 1=1" + jefe_filter_mrk)

            dias_totales = (_exec_one(f"""
                SELECT COUNT(DISTINCT fecha) as dias FROM (
                    SELECT gaming_date as fecha FROM srw_jugadores {sw_solo}
                    UNION
                    SELECT fecha_operacion as fecha FROM mesas_puntos {mw_solo_dias}
                ) sub
            """, {**sp_solo, **mp_solo_dias})).get('dias') or 1

            pct_primario, pct_categoria = _get_invitaciones_config()
            all_params = {**cparam, **pparam, **mparam, **sp_plain, **mp_plain, **sp}

            jugadores_srw = _exec(f"""
                SELECT s.player_id,
                       MAX(s.full_name) as nombre, MAX(s.player_level) as nivel,
                       SUM(s.coin_in) + COALESCE(MAX(m.coin_in_mesas), 0) as coin_in_mensual,
                       COALESCE(MAX(d.dias_combinados), COUNT(DISTINCT s.gaming_date)) as dias_asistidos,
                       COALESCE(MAX(c.total_cortesias), 0) as total_cortesias,
                       COALESCE(MAX(c.monto_micros), 0) as monto_micros,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios
                FROM srw_jugadores s
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_micros
                    FROM cortesias c {cw_inner} GROUP BY cliente_id
                ) c ON s.player_id = c.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps p {pw} GROUP BY cliente_id
                ) p ON s.player_id = p.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, SUM(coin_in_puntos) as coin_in_mesas
                    FROM mesas_puntos m {mw} GROUP BY cliente_id
                ) m ON s.player_id = m.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, COUNT(DISTINCT fecha) as dias_combinados FROM (
                        SELECT player_id as cliente_id, gaming_date as fecha FROM srw_jugadores {sw_plain}
                        UNION
                        SELECT cliente_id, fecha_operacion as fecha FROM mesas_puntos {mw_plain}
                    ) sub GROUP BY cliente_id
                ) d ON s.player_id = d.cliente_id
                {sw}
                GROUP BY s.player_id
                HAVING COALESCE(MAX(c.total_cortesias), 0) > 0
                ORDER BY coin_in_mensual DESC
            """, all_params)

            # Mesas-only (MRK)
            sw_excl, sp_excl = build_date_filter('gaming_date', anio, mes)
            mw_inner, mparam_inner = build_date_filter('mp.fecha_operacion', anio, mes)
            mesas_excl = f"mp.cliente_id NOT IN (SELECT DISTINCT player_id FROM srw_jugadores {sw_excl})"
            if mparam_inner:
                mw_conds = []
                if anio:
                    mw_conds.append(f"SUBSTR(mp.fecha_operacion, 1, 4) = :anio_mp_fecha_operacion")
                if mes:
                    mw_conds.append(f"SUBSTR(mp.fecha_operacion, 6, 2) = :mes_mp_fecha_operacion")
                mesas_where = "WHERE " + " AND ".join(mw_conds) + " AND " + mesas_excl
            else:
                mesas_where = "WHERE " + mesas_excl

            cw_inner_mesas = (cw + jefe_filter_mrk) if cw else ("WHERE 1=1" + jefe_filter_mrk)
            mesas_params = {**cparam, **pparam, **mparam_inner, **sp_excl}

            jugadores_mesas = _exec(f"""
                SELECT mp.cliente_id as player_id,
                       MAX(mp.cliente_nombre) as nombre, COALESCE(MAX(mp.categoria), 'Sin Categoria') as nivel,
                       SUM(mp.coin_in_puntos) as coin_in_mensual,
                       COUNT(DISTINCT mp.fecha_operacion) as dias_asistidos,
                       COALESCE(MAX(c.total_cortesias), 0) as total_cortesias,
                       COALESCE(MAX(c.monto_micros), 0) as monto_micros,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios
                FROM mesas_puntos mp
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as total_cortesias, SUM(micros) as monto_micros
                    FROM cortesias c {cw_inner_mesas} GROUP BY cliente_id
                ) c ON mp.cliente_id = c.cliente_id
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps p {pw} GROUP BY cliente_id
                ) p ON mp.cliente_id = p.cliente_id
                {mesas_where}
                GROUP BY mp.cliente_id
                HAVING COALESCE(MAX(c.total_cortesias), 0) > 0
                ORDER BY coin_in_mensual DESC
            """, mesas_params)

            # Cort-only (MRK)
            ids_ya_mrk = set(j['player_id'] for j in jugadores_srw) | set(j['player_id'] for j in jugadores_mesas)
            cort_only_mrk = _exec(f"""
                SELECT c.cliente_id as player_id, MAX(c.nombre_cliente) as nombre,
                       '-' as nivel, 0 as coin_in_mensual, 0 as dias_asistidos,
                       COUNT(c.id) as total_cortesias, SUM(c.micros) as monto_micros,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios
                FROM cortesias c
                LEFT JOIN (
                    SELECT cliente_id, COUNT(*) as cant_premios, SUM(transferencia_final) as monto_premios
                    FROM premios_comps p {pw} GROUP BY cliente_id
                ) p ON c.cliente_id = p.cliente_id
                {cw_inner}
                GROUP BY c.cliente_id ORDER BY monto_micros DESC
            """, {**cparam, **pparam})
            cort_only_mrk = [j for j in cort_only_mrk if j['player_id'] not in ids_ya_mrk]

            all_mrk = list(jugadores_srw) + list(jugadores_mesas) + cort_only_mrk
            resultados_mrk = _calc_invitaciones(all_mrk, pct_primario, pct_categoria, dias_totales)
            _write_sheet(writer, resultados_mrk, 'Ctrl - MRK',
                         ['Nombre', 'Nivel', 'Días Asist.', '% Asist.', 'Premios', 'Monto Premios',
                          'Coin-In Mensual', 'Cortesías', 'Monto Cortesías',
                          'Invitación Mensual', 'Saldo', '% Cat.'])

        # ── AUDITORÍA COIN-IN CERO ─────────────────────────────────
        if 'coinin_cero' in secciones:
            cw_audit, cp_audit = build_date_filter('c.fecha_jornada', anio, mes)

            where_parts = []
            all_params = {}
            if cw_audit:
                where_parts.append(cw_audit.replace("WHERE ", ""))
                all_params.update(cp_audit)
            where_parts.append("""(
                NOT EXISTS (
                    SELECT 1 FROM srw_jugadores s
                    WHERE s.player_id = c.cliente_id AND s.gaming_date = c.fecha_jornada AND s.coin_in > 0
                )
                AND NOT EXISTS (
                    SELECT 1 FROM mesas_puntos m
                    WHERE m.cliente_id = c.cliente_id AND m.fecha_operacion = c.fecha_jornada AND m.coin_in_puntos > 0
                )
            )""")
            where_clause = "WHERE " + " AND ".join(where_parts)

            rows = _exec(f"""
                SELECT c.fecha_jornada as jornada, c.cliente_id,
                       MAX(c.nombre_cliente) as nombre_cliente,
                       COALESCE(MAX(s.coin_in_dia), 0) + COALESCE(MAX(me.coin_in_mesas), 0) as coin_in,
                       COUNT(c.id) as cant_cortesias, SUM(c.micros) as monto_cortesias,
                       COALESCE(MAX(p.cant_premios), 0) as cant_premios,
                       COALESCE(MAX(p.monto_premios), 0) as monto_premios,
                       COALESCE(MAX(j.nombre), '') as jefe_nombre,
                       COALESCE(MAX(j.area), '') as jefe_area
                FROM cortesias c
                LEFT JOIN (
                    SELECT player_id, gaming_date, SUM(coin_in) as coin_in_dia
                    FROM srw_jugadores GROUP BY player_id, gaming_date
                ) s ON c.cliente_id = s.player_id AND c.fecha_jornada = s.gaming_date
                LEFT JOIN (
                    SELECT cliente_id, fecha_operacion, SUM(coin_in_puntos) as coin_in_mesas
                    FROM mesas_puntos GROUP BY cliente_id, fecha_operacion
                ) me ON c.cliente_id = me.cliente_id AND c.fecha_jornada = me.fecha_operacion
                LEFT JOIN (
                    SELECT cliente_id, fecha_jornada, COUNT(*) as cant_premios,
                           SUM(transferencia_final) as monto_premios
                    FROM premios_comps GROUP BY cliente_id, fecha_jornada
                ) p ON c.cliente_id = p.cliente_id AND c.fecha_jornada = p.fecha_jornada
                LEFT JOIN jefaturas j ON c.usuario_id = j.usuario_id
                {where_clause}
                GROUP BY c.fecha_jornada, c.cliente_id, c.usuario_id
                ORDER BY c.fecha_jornada DESC, monto_cortesias DESC
            """, all_params)
            _write_sheet(writer, rows, 'Coin-In Cero',
                         ['Jornada', 'ID Cliente', 'Nombre', 'Coin-In',
                          'Cortesías', 'Monto Cortesías', 'Premios', 'Monto Premios',
                          'Jefe', 'Área'])

    output.seek(0)
    filename = f"Reporte_COMPS_{periodo}.xlsx"

    # En modo desktop, guardar en Descargas directamente
    if os.environ.get("SGOS_DESKTOP") == "1":
        from pathlib import Path as _Path
        downloads = _Path.home() / "Downloads"
        downloads.mkdir(exist_ok=True)
        dest = downloads / filename
        counter = 1
        stem, suffix = _Path(filename).stem, _Path(filename).suffix
        while dest.exists():
            dest = downloads / f"{stem} ({counter}){suffix}"
            counter += 1
        dest.write_bytes(output.getvalue())
        return jsonify({"saved": True, "path": str(dest)})

    return send_file(output, download_name=filename,
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
