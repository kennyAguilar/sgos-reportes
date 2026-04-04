"""
comps_engine.py — ETL para módulo Auditoría COMPS.
Lee archivos Excel de SRW, Cortesías, Premios, Mesas/Puntos y Jefaturas,
y los carga en las tablas PostgreSQL correspondientes.
"""
import os
import pandas as pd
from datetime import datetime
from sqlalchemy import types as sa_types


def limpiar_player_id(val):
    if pd.isna(val):
        return None
    s = str(val).strip().strip('x')
    return s if s else None


def cargar_srw(filepath):
    df = pd.read_excel(filepath, header=None, skiprows=3)
    df = df.iloc[:, 1:]
    df.columns = [
        'gaming_date', 'player_id', 'full_name', 'player_level',
        'coin_in', 'rec_cin', 'coin_out', 'rec_cout',
        'jackpot_amount', 'promo_in', 'promo_out', 'prom_jugado',
        'win_loss_mda', 'win_loss_mda_rec', 'bill_in',
        'total_games', 'total_egm_points'
    ]
    df = df[['gaming_date', 'player_id', 'full_name', 'player_level',
             'coin_in', 'total_games', 'promo_in']]
    df = df.dropna(subset=['player_id'])
    df['player_id'] = df['player_id'].astype(str).str.strip()
    df['gaming_date'] = pd.to_datetime(df['gaming_date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df = df.dropna(subset=['gaming_date'])
    df['coin_in'] = pd.to_numeric(df['coin_in'], errors='coerce').fillna(0).astype(float)
    df['total_games'] = pd.to_numeric(df['total_games'], errors='coerce').fillna(0).astype(float)
    df['promo_in'] = pd.to_numeric(df['promo_in'], errors='coerce').fillna(0).astype(float)
    return df


def cargar_cortesias(filepath):
    df = pd.read_excel(filepath, header=None, skiprows=8)
    df = df.rename(columns={
        6: 'fecha_jornada', 7: 'cliente_id', 10: 'nombre_cliente',
        14: 'descripcion_cat', 16: 'descripcion_prod', 19: 'micros',
        22: 'estado', 28: 'usuario_id', 29: 'nombre_usuario'
    })
    cols = ['fecha_jornada', 'cliente_id', 'nombre_cliente',
            'descripcion_cat', 'descripcion_prod', 'micros',
            'estado', 'usuario_id', 'nombre_usuario']
    df = df[cols]
    df = df[df['estado'] == 'QUEMADO']
    df = df.dropna(subset=['cliente_id'])
    df['cliente_id'] = df['cliente_id'].astype(str).str.strip()
    df['fecha_jornada'] = pd.to_datetime(df['fecha_jornada'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['micros'] = pd.to_numeric(df['micros'], errors='coerce').fillna(0)
    df['usuario_id'] = df['usuario_id'].astype(str).str.replace(r'\.0$', '', regex=True)
    df['nombre_usuario'] = df['nombre_usuario'].fillna('')
    return df


def cargar_premios_comps(filepath):
    df = pd.read_excel(filepath, header=1)
    df.columns = [
        'fecha', 'maquina', 'id_mensaje', 'cliente_id',
        'monto_transferido', 'propina', 'transferencia_final',
        'slot_attendant', 'monto_slot_atten', 'validador',
        'monto_validador', 'tipo_pago', 'ingreso_cawa'
    ]
    df = df[df['tipo_pago'].isin(['Jackpot HP', 'Progressive Jackpot HP'])]
    df = df.dropna(subset=['cliente_id'])
    df['cliente_id'] = df['cliente_id'].astype(str).str.strip().str.strip('x')
    df['transferencia_final'] = pd.to_numeric(df['transferencia_final'], errors='coerce').fillna(0)
    df['fecha_dt'] = pd.to_datetime(df['fecha'], format='%d-%m-%Y %H:%M', errors='coerce')
    df = df.dropna(subset=['fecha_dt'])
    df['fecha_jornada'] = df['fecha_dt'].apply(
        lambda dt: (dt - pd.Timedelta(hours=9)).strftime('%Y-%m-%d')
    )
    return df[['fecha_jornada', 'cliente_id', 'transferencia_final', 'tipo_pago']]


def cargar_mesas_puntos(filepath):
    df = pd.read_excel(filepath, header=None, skiprows=2)
    df = df.iloc[:, 1:]
    df.columns = [
        'fecha_operacion', 'sesion_id', 'mesa_id', 'juego',
        'cliente_id', 'cliente_nombre',
        'hora_inicio', 'hora_fin',
        'tpo_jugado', 'ap_promedio', 'puntos'
    ]
    df = df[['fecha_operacion', 'cliente_id', 'cliente_nombre', 'puntos']]
    df = df.dropna(subset=['cliente_id'])
    df['cliente_id'] = df['cliente_id'].astype(str).str.strip()
    df['fecha_operacion'] = pd.to_datetime(df['fecha_operacion'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['puntos'] = pd.to_numeric(df['puntos'], errors='coerce').fillna(0)
    df['coin_in_puntos'] = df['puntos'] * 1000
    return df


def cargar_jefaturas(filepath):
    """Carga Jefatura.xlsx (Hoja1=jefaturas, Hoja2=categorias_nivel)."""
    result = {}
    # Hoja1: jefaturas
    df1 = pd.read_excel(filepath, sheet_name='Hoja1', header=0)
    df1.columns = ['usuario_id', 'nombre', 'area']
    df1['usuario_id'] = df1['usuario_id'].astype(str).str.strip()
    df1['nombre'] = df1['nombre'].fillna('')
    df1['area'] = df1['area'].fillna('')
    result['jefaturas'] = df1

    # Hoja2: categorias_nivel
    df2 = pd.read_excel(filepath, sheet_name='Hoja2', header=0)
    df2.columns = ['categoria', 'porcentaje']
    df2 = df2.dropna(subset=['categoria'])
    df2['porcentaje'] = pd.to_numeric(df2['porcentaje'], errors='coerce').fillna(0)
    result['categorias_nivel'] = df2

    return result


def guardar_comps_db(db, tabla, etl_fn, filepath, col_fecha, models_map):
    """
    Genérico: lee Excel con etl_fn, borra registros del rango de fechas, e inserta.
    Retorna (filas, tabla).
    models_map: dict con nombre de tabla → modelo SQLAlchemy para update post-carga.
    """
    df = etl_fn(filepath)
    if df.empty:
        return 0

    fechas = df[col_fecha].dropna()
    if not fechas.empty:
        fecha_min = fechas.min()
        fecha_max = fechas.max()
        from sqlalchemy import text
        db.session.execute(
            text(f"DELETE FROM {tabla} WHERE {col_fecha} BETWEEN :fmin AND :fmax"),
            {"fmin": fecha_min, "fmax": fecha_max}
        )

    # Insertar con pandas to_sql (append) – forzar tipos para evitar overflow
    # Todas las columnas string → Text, numéricas → Float para evitar Integer overflow
    dtype = {}
    for col in df.columns:
        if df[col].dtype == 'object':
            dtype[col] = sa_types.Text()
        else:
            dtype[col] = sa_types.Float()
    df.to_sql(tabla, db.engine, if_exists='append', index=False, dtype=dtype)
    db.session.commit()
    return len(df)


def guardar_jefaturas_db(db, filepath):
    """Carga Jefatura.xlsx y reemplaza jefaturas y categorias_nivel."""
    from sqlalchemy import text
    data = cargar_jefaturas(filepath)

    db.session.execute(text("DELETE FROM jefaturas"))
    data['jefaturas'].to_sql('jefaturas', db.engine, if_exists='append', index=False)

    db.session.execute(text("DELETE FROM categorias_nivel"))
    data['categorias_nivel'].to_sql('categorias_nivel', db.engine, if_exists='append', index=False)

    db.session.commit()


def actualizar_nombres_cortesias(db):
    """Actualiza nombre_cliente en cortesias desde SRW, los que falten marca como (Sin registro)."""
    from sqlalchemy import text
    db.session.execute(text("""
        UPDATE cortesias SET nombre_cliente = (
            SELECT s.full_name FROM srw_jugadores s
            WHERE s.player_id = cortesias.cliente_id LIMIT 1
        ) WHERE nombre_cliente IS NULL OR TRIM(nombre_cliente) = ''
    """))
    db.session.execute(text("""
        UPDATE cortesias SET nombre_cliente = '(Sin registro en SRW)'
        WHERE nombre_cliente IS NULL OR TRIM(nombre_cliente) = ''
    """))
    db.session.commit()
