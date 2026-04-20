import os
import gc
import time
import uuid
from datetime import datetime
from io import BytesIO
import pandas as pd
from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session, abort, jsonify
from sqlalchemy import select, func, distinct
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from flask_wtf.csrf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

load_dotenv()  # Carga las variables del archivo .env

from pathlib import Path

def desktop_save_response(output, filename):
    """En modo desktop, guarda el archivo en Descargas y devuelve JSON."""
    if os.environ.get("SGOS_DESKTOP") == "1":
        downloads = Path.home() / "Downloads"
        downloads.mkdir(exist_ok=True)
        dest = downloads / filename
        counter = 1
        stem, suffix = Path(filename).stem, Path(filename).suffix
        while dest.exists():
            dest = downloads / f"{stem} ({counter}){suffix}"
            counter += 1
        dest.write_bytes(output.getvalue())
        return jsonify({"saved": True, "path": str(dest)})
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

try:
    from sgos_web.engine import procesar_sgos, exportar_excel_bytes, obtener_asistentes, guardar_datos_db, generar_reportes, _cargar_df, generar_kpis
except ImportError:
    from engine import procesar_sgos, exportar_excel_bytes, obtener_asistentes, guardar_datos_db, generar_reportes, _cargar_df, generar_kpis

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or os.urandom(32).hex()

# Cookies de sesión seguras
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
# SESSION_COOKIE_SECURE solo en producción con HTTPS (no en desktop/localhost)
is_desktop = os.environ.get("SGOS_DESKTOP") == "1"
if not is_desktop and os.environ.get("FLASK_ENV") != "development":
    app.config['SESSION_COOKIE_SECURE'] = True

# Protección CSRF
csrf = CSRFProtect(app)

# Rate Limiting
limiter = Limiter(get_remote_address, app=app, storage_uri="memory://")

import re

MIN_PASSWORD_LENGTH = 9

def validar_password(password):
    """Retorna mensaje de error o None si es válida."""
    if len(password) < MIN_PASSWORD_LENGTH:
        return f"La contraseña debe tener al menos {MIN_PASSWORD_LENGTH} caracteres."
    if not re.search(r'[A-Z]', password):
        return "La contraseña debe contener al menos una letra mayúscula."
    if not re.search(r'[!@#$%^&*()_+\-=\[\]{}|;:\'",.<>?/`~]', password):
        return "La contraseña debe contener al menos un carácter especial."
    return None

# Configuración de Login
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"

# Configuración de Base de Datos (Neon PostgreSQL obligatorio)
db_url = os.environ.get("DATABASE_URL")
if not db_url:
    raise RuntimeError(
        "La variable de entorno DATABASE_URL es obligatoria. "
        "Configúrala con tu connection string de Neon PostgreSQL."
    )

app.config['SQLALCHEMY_DATABASE_URI'] = db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

try:
    from sgos_web.extensions import db
except ImportError:
    from extensions import db

db.init_app(app)

# --- MODELOS ---
class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))

class Operacion(db.Model):
    __tablename__ = 'operaciones'

    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.DateTime, nullable=False)
    jornada = db.Column(db.DateTime, nullable=False)
    id_cliente = db.Column(db.String(100))
    monto = db.Column(db.Float, default=0.0)
    voucher = db.Column(db.String(100))
    attendant = db.Column(db.String(100), nullable=False)
    validador = db.Column(db.String(100))
    forma_pago = db.Column(db.String(50))
    ingreso_cawa = db.Column(db.String(50))
    # tipo = db.Column(db.String(50), default='GETNET') # Eliminado, usaremos tabla separada
    
    # Campos calculados útiles para consultas rápidas
    mes = db.Column(db.String(7))  # YYYY-MM
    hora = db.Column(db.Integer)

    __table_args__ = (
        db.Index('ix_operaciones_mes', 'mes'),
        db.Index('ix_operaciones_jornada', 'jornada'),
        db.Index('ix_operaciones_attendant', 'attendant'),
        db.Index('ix_operaciones_mes_attendant', 'mes', 'attendant'),
    )

    def __repr__(self):
        return f"<Operacion {self.id} - {self.attendant} - {self.monto}>"

class Premio(db.Model):
    __tablename__ = 'premios'
    
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.DateTime, nullable=False)
    jornada = db.Column(db.DateTime, nullable=False)
    id_cliente = db.Column(db.String(100))
    monto = db.Column(db.Float, default=0.0) # Transferencia Final
    propina = db.Column(db.Float, default=0.0)
    maquina = db.Column(db.String(50))
    attendant = db.Column(db.String(100), nullable=False)
    validador = db.Column(db.String(100))
    forma_pago = db.Column(db.String(50))
    ingreso_cawa = db.Column(db.String(50))
    
    # Campos calculados
    mes = db.Column(db.String(7))
    hora = db.Column(db.Integer)

    __table_args__ = (
        db.Index('ix_premios_mes', 'mes'),
        db.Index('ix_premios_jornada', 'jornada'),
        db.Index('ix_premios_attendant', 'attendant'),
        db.Index('ix_premios_mes_attendant', 'mes', 'attendant'),
    )

    def __repr__(self):
        return f"<Premio {self.id} - {self.attendant} - {self.monto}>"

# --- MODELOS COMPS ---
class SrwJugador(db.Model):
    __tablename__ = 'srw_jugadores'
    id = db.Column(db.Integer, primary_key=True)
    gaming_date = db.Column(db.String(10))
    player_id = db.Column(db.String(50))
    full_name = db.Column(db.String(200))
    player_level = db.Column(db.String(50))
    coin_in = db.Column(db.Float, default=0)
    total_games = db.Column(db.Float, default=0)
    promo_in = db.Column(db.Float, default=0)

class Cortesia(db.Model):
    __tablename__ = 'cortesias'
    id = db.Column(db.Integer, primary_key=True)
    fecha_jornada = db.Column(db.String(10))
    cliente_id = db.Column(db.String(50))
    nombre_cliente = db.Column(db.String(200))
    descripcion_cat = db.Column(db.String(200))
    descripcion_prod = db.Column(db.String(200))
    micros = db.Column(db.Float, default=0)
    estado = db.Column(db.String(50))
    usuario_id = db.Column(db.String(50))
    nombre_usuario = db.Column(db.String(200))

class PremioComps(db.Model):
    __tablename__ = 'premios_comps'
    id = db.Column(db.Integer, primary_key=True)
    fecha_jornada = db.Column(db.String(10))
    cliente_id = db.Column(db.String(50))
    transferencia_final = db.Column(db.Float, default=0)
    tipo_pago = db.Column(db.String(100))

class MesaPuntos(db.Model):
    __tablename__ = 'mesas_puntos'
    id = db.Column(db.Integer, primary_key=True)
    fecha_operacion = db.Column(db.String(10))
    cliente_id = db.Column(db.String(50))
    cliente_nombre = db.Column(db.String(200))
    categoria = db.Column(db.String(100))
    puntos = db.Column(db.Float, default=0)
    coin_in_puntos = db.Column(db.Float, default=0)

class Jefatura(db.Model):
    __tablename__ = 'jefaturas'
    id = db.Column(db.Integer, primary_key=True)
    usuario_id = db.Column(db.String(50), unique=True)
    nombre = db.Column(db.String(200))
    area = db.Column(db.String(100))

class CategoriaNivel(db.Model):
    __tablename__ = 'categorias_nivel'
    id = db.Column(db.Integer, primary_key=True)
    categoria = db.Column(db.String(100), unique=True)
    porcentaje = db.Column(db.Float, default=0)

class CargaLog(db.Model):
    __tablename__ = 'carga_log'
    id = db.Column(db.Integer, primary_key=True)
    tabla = db.Column(db.String(50))
    archivo = db.Column(db.String(200))
    filas = db.Column(db.Integer)
    fecha_carga = db.Column(db.String(50))
    file_hash = db.Column(db.String(128), index=True)
    modo = db.Column(db.String(20), default='replace')
    usuario = db.Column(db.String(100))
    descartados = db.Column(db.Integer, default=0)
    meses = db.Column(db.String(200))

# Crear tablas si no existen (solo para desarrollo local/inicial)
with app.app_context():
    db.create_all()

    # Auto-migración ligera: añade columnas nuevas a carga_log si faltan (SQLite / Postgres)
    try:
        with db.engine.begin() as _conn:
            _existentes = {c["name"] for c in db.inspect(db.engine).get_columns("carga_log")}
            _nuevas = {
                "file_hash": "VARCHAR(128)",
                "modo": "VARCHAR(20)",
                "usuario": "VARCHAR(100)",
                "descartados": "INTEGER DEFAULT 0",
                "meses": "VARCHAR(200)",
            }
            for _col, _tipo in _nuevas.items():
                if _col not in _existentes:
                    _conn.exec_driver_sql(f"ALTER TABLE carga_log ADD COLUMN {_col} {_tipo}")
                    print(f"[migración] carga_log: columna '{_col}' añadida")
    except Exception as _mig_err:
        print(f"[migración] aviso: {_mig_err}")

    # Crear usuario admin por defecto si no existe
    if not User.query.filter_by(username="admin").first():
        admin = User(username="admin")
        default_pw = os.environ.get("ADMIN_DEFAULT_PASSWORD", "admin123")
        admin.set_password(default_pw)
        db.session.add(admin)
        db.session.commit()
        print("Usuario 'admin' creado. Cambia la contraseña desde Gestión de Usuarios.")

UPLOAD_FOLDER = os.path.join(os.path.abspath(os.getcwd()), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB (ajusta si quieres)

ALLOWED_EXT = {".xlsx", ".xls"}
TABLAS_NO_FILTRAR = {
    "Resumen Mensual", 
    "Operaciones por Hora", 
    "Conteo mensual de operaciones por MDA",
    "Conteo total de operaciones por MDA"
}


def allowed_file(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXT


def safe_file_path(file_id: str) -> str:
    """
    Evita path traversal: normaliza y obliga a estar dentro de uploads.
    """
    file_id = secure_filename(file_id)
    path = os.path.abspath(os.path.join(app.config["UPLOAD_FOLDER"], file_id))
    base = os.path.abspath(app.config["UPLOAD_FOLDER"])
    if not path.startswith(base + os.sep):
        abort(400, "file_id inválido.")
    return path


def aplicar_opciones(tablas: dict, opciones: list[str]) -> dict:
    if not opciones:
        return tablas
    return {k: v for k, v in tablas.items() if k in opciones}


def cleanup_old_uploads(folder, max_age_seconds=3600):
    """Elimina archivos subidos con más de max_age_seconds de antigüedad."""
    now = time.time()
    for fname in os.listdir(folder):
        fpath = os.path.join(folder, fname)
        if os.path.isfile(fpath) and now - os.path.getmtime(fpath) > max_age_seconds:
            try:
                os.remove(fpath)
            except OSError:
                pass


def preparar_tablas(path: str, opciones: list[str], asistentes_sel: list[str]) -> tuple[dict, list]:
    """
    Lee el Excel UNA sola vez. Retorna (tablas, asistentes_disponibles).
    """
    df = _cargar_df(path)
    asistentes_disponibles = sorted(df["Attendant"].dropna().unique().tolist())
    tablas_base = generar_reportes(df)

    if not asistentes_sel or set(asistentes_sel) == set(asistentes_disponibles):
        return aplicar_opciones(tablas_base, opciones), asistentes_disponibles

    tablas_filtradas = generar_reportes(df, asistentes_sel)
    for nombre in TABLAS_NO_FILTRAR:
        if nombre in tablas_base:
            tablas_filtradas[nombre] = tablas_base[nombre]

    return aplicar_opciones(tablas_filtradas, opciones), asistentes_disponibles


def tablas_a_html(tablas: dict) -> dict:
    return {
        k: v.to_html(index=False, classes="table table-sm table-striped w-auto mx-auto sortable")
        for k, v in tablas.items()
    }


@app.after_request
def set_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'geolocation=(), camera=(), microphone=()'
    if os.environ.get("FLASK_ENV") != "development":
        response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    return response


@app.route("/sw.js")
def service_worker():
    return app.send_static_file("js/sw.js"), 200, {"Content-Type": "application/javascript", "Service-Worker-Allowed": "/"}


@app.route("/health")
def health():
    return "OK", 200


@app.route("/login", methods=["GET", "POST"])
@limiter.limit("10 per minute")
def login():
    if current_user.is_authenticated:
        return redirect(url_for("home"))
        
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")
        user = User.query.filter_by(username=username).first()
        
        if user and user.check_password(password):
            login_user(user)
            return redirect(url_for("home"))
        else:
            flash("Usuario o contraseña incorrectos.", "danger")
            
    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def home():
    return render_template("home.html")


# ─────────────────────────── Gestión de Usuarios ───────────────────────────

@app.route("/usuarios")
@login_required
def usuarios():
    users = User.query.order_by(User.username).all()
    return render_template("usuarios.html", users=users)


@app.route("/usuarios/crear", methods=["POST"])
@login_required
def crear_usuario():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "").strip()
    if not username or not password:
        flash("Usuario y contraseña son obligatorios.", "warning")
        return redirect(url_for("usuarios"))
    error_pw = validar_password(password)
    if error_pw:
        flash(error_pw, "warning")
        return redirect(url_for("usuarios"))
    if User.query.filter_by(username=username).first():
        flash(f"El usuario '{username}' ya existe.", "warning")
        return redirect(url_for("usuarios"))
    user = User(username=username)
    user.set_password(password)
    db.session.add(user)
    db.session.commit()
    flash(f"Usuario '{username}' creado exitosamente.", "success")
    return redirect(url_for("usuarios"))


@app.route("/usuarios/eliminar/<int:user_id>", methods=["POST"])
@login_required
def eliminar_usuario(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("Usuario no encontrado.", "danger")
    elif user.id == current_user.id:
        flash("No puedes eliminarte a ti mismo.", "warning")
    else:
        flash(f"Usuario '{user.username}' eliminado.", "success")
        db.session.delete(user)
        db.session.commit()
    return redirect(url_for("usuarios"))


@app.route("/usuarios/cambiar-password/<int:user_id>", methods=["POST"])
@login_required
def cambiar_password(user_id):
    user = db.session.get(User, user_id)
    if not user:
        flash("Usuario no encontrado.", "danger")
        return redirect(url_for("usuarios"))
    new_password = request.form.get("password", "").strip()
    if not new_password:
        flash("La contraseña no puede estar vacía.", "warning")
        return redirect(url_for("usuarios"))
    error_pw = validar_password(new_password)
    if error_pw:
        flash(error_pw, "warning")
        return redirect(url_for("usuarios"))
    user.set_password(new_password)
    db.session.commit()
    flash(f"Contraseña de '{user.username}' actualizada.", "success")
    return redirect(url_for("usuarios"))


@app.route("/sgos", methods=["GET", "POST"])
@login_required
def index():
    if request.method == "POST":
        f = request.files.get("file")
        if not f or f.filename == "":
            flash("No se subió ningún archivo.", "warning")
            return redirect(url_for("index"))

        if not allowed_file(f.filename):
            flash("Formato no permitido. Sube un .xlsx o .xls", "warning")
            return redirect(url_for("index"))

        filename = secure_filename(f.filename)
        token = uuid.uuid4().hex
        saved_name = f"{token}__{filename}"
        path = os.path.join(app.config["UPLOAD_FOLDER"], saved_name)
        f.save(path)

        cleanup_old_uploads(app.config["UPLOAD_FOLDER"])

        # Modo de carga: 'replace' (default) borra el mes y recarga; 'append' solo inserta
        modo = request.form.get("modo_carga", "replace")
        if modo not in ("replace", "append"):
            modo = "replace"

        # --- Guardar en Base de Datos ---
        try:
            resultado = guardar_datos_db(path, db, Operacion, Premio, modo=modo)

            # Validación: faltan columnas requeridas
            if resultado.get("columnas_faltantes"):
                flash(
                    f"❌ Archivo inválido. Faltan columnas: {', '.join(resultado['columnas_faltantes'])}",
                    "danger",
                )
                return redirect(url_for("index"))

            # Error genérico del engine
            if resultado.get("error") and resultado["guardados"] == 0:
                flash(f"⚠️ {resultado['error']}", "warning")
                return redirect(url_for("index"))

            # Detectar duplicado por hash
            file_hash = resultado.get("file_hash")
            duplicado = None
            if file_hash:
                duplicado = CargaLog.query.filter_by(file_hash=file_hash).first()

            # Registrar en log de cargas
            desc = resultado["descartados"]
            total_descartados = desc["fecha"] + desc["attendant"] + desc["hora"]
            try:
                log = CargaLog(
                    tabla=resultado["tipo"].lower(),
                    archivo=filename,
                    filas=resultado["guardados"],
                    fecha_carga=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    file_hash=file_hash,
                    modo=modo,
                    usuario=current_user.username if current_user.is_authenticated else None,
                    descartados=total_descartados,
                    meses=",".join(resultado.get("meses", [])),
                )
                db.session.add(log)
                db.session.commit()
            except Exception as log_err:
                app.logger.warning(f"No se pudo registrar CargaLog: {log_err}")
                db.session.rollback()

            gc.collect()

            # Mensaje principal
            flash(
                f"✅ {resultado['guardados']} registros {resultado['tipo']} guardados "
                f"({'reemplazo' if modo == 'replace' else 'acumulativo'}) · "
                f"meses: {', '.join(resultado.get('meses', []))}",
                "success",
            )

            # Reporte de descartes (si hay)
            if total_descartados > 0:
                detalles = []
                if desc["fecha"]:
                    detalles.append(f"{desc['fecha']} sin fecha/jornada")
                if desc["attendant"]:
                    detalles.append(f"{desc['attendant']} sin attendant")
                if desc["hora"]:
                    detalles.append(f"{desc['hora']} fuera de jornada (09 h)")
                flash(
                    f"ℹ️ {total_descartados} filas descartadas de {resultado['total_leido']} leídas — "
                    + "; ".join(detalles),
                    "info",
                )

            # Aviso de montos negativos
            if desc["monto_neg"] > 0:
                flash(
                    f"⚠️ Se detectaron {desc['monto_neg']} registros con monto negativo. Revisa en el histórico.",
                    "warning",
                )

            # Aviso de duplicado
            if duplicado and duplicado.filas > 0:
                flash(
                    f"ℹ️ Este archivo ya había sido cargado el {duplicado.fecha_carga} "
                    f"({duplicado.filas} filas). Los datos se {'reemplazaron' if modo == 'replace' else 'acumularon'} igualmente.",
                    "info",
                )

        except Exception as e:
            app.logger.error(f"Error al guardar en base de datos: {e}")
            flash("Error al guardar en base de datos. Contacta al administrador.", "danger")

        opciones = request.form.getlist("opciones")  # lo que marcó en index
        session[f"tablas_{saved_name}"] = opciones

        # OJO: NO guardamos listas grandes en session.
        # Deja que dashboard recalculé 'asistentes_disponibles' desde el archivo.
        # Guardamos solo selección (por defecto: vacío => se interpreta como "todos").
        session[f"asistentes_sel_{saved_name}"] = []

        return redirect(url_for("dashboard", file_id=saved_name))

    return render_template("index.html")


@app.route("/dashboard/<file_id>", methods=["GET", "POST"])
@login_required
def dashboard(file_id):
    path = safe_file_path(file_id)
    if not os.path.exists(path):
        return "Archivo no encontrado.", 404

    try:
        if request.method == "POST":
            asistentes_sel = request.form.getlist("asistentes")
            session[f"asistentes_sel_{file_id}"] = asistentes_sel
            return redirect(url_for("dashboard", file_id=file_id))

        opciones = session.get(f"tablas_{file_id}", [])
        asistentes_sel = session.get(f"asistentes_sel_{file_id}", [])

        tablas, asistentes_disponibles = preparar_tablas(path, opciones, asistentes_sel)
        asistentes_seleccionados = asistentes_sel or asistentes_disponibles

        return render_template(
            "dashboard.html",
            file_id=file_id,
            tablas_html=tablas_a_html(tablas),
            asistentes_disponibles=asistentes_disponibles,
            asistentes_seleccionados=asistentes_seleccionados
        )
    except Exception as e:
        app.logger.error(f"Error en dashboard: {e}")
        flash("Error al procesar el archivo. Contacta al administrador.", "danger")
        return redirect(url_for("index"))


def get_available_dates(model):
    """Obtiene años y meses disponibles en la base de datos."""
    with db.engine.connect() as conn:
        result = conn.execute(select(model.mes).distinct().where(model.mes != None).order_by(model.mes.desc()))
        meses = [row[0] for row in result]
    
    years = sorted(list(set(m.split('-')[0] for m in meses)), reverse=True)
    return years, meses


def get_db_dataframe(year=None, month=None):
    """Consulta la base de datos y devuelve un DataFrame con el formato esperado por engine.py"""
    query = select(Operacion)
    
    if year and year != "all":
        if month and month != "all":
            query = query.filter(Operacion.mes == f"{year}-{month}")
        else:
            query = query.filter(Operacion.mes.like(f"{year}-%"))

    with db.engine.connect() as conn:
        df = pd.read_sql(query, conn)
    
    if df.empty:
        return df

    # Renombrar columnas para coincidir con engine.py
    df = df.rename(columns={
        "fecha": "Fecha",
        "jornada": "Jornada",
        "monto": "Monto",
        "attendant": "Attendant",
        "mes": "Mes",
        "hora": "Hora",
        "forma_pago": "FormaPago"
    })
    
    # Calcular JornadaDia
    df["JornadaDia"] = pd.to_datetime(df["Jornada"]).dt.normalize()
    df["Tipo"] = "GETNET"
    
    return df


def get_premios_dataframe(year=None, month=None):
    """Consulta la base de datos de PREMIOS y devuelve un DataFrame"""
    query = select(Premio)
    
    if year and year != "all":
        if month and month != "all":
            query = query.filter(Premio.mes == f"{year}-{month}")
        else:
            query = query.filter(Premio.mes.like(f"{year}-%"))

    with db.engine.connect() as conn:
        df = pd.read_sql(query, conn)
    
    if df.empty:
        return df

    # Renombrar columnas para coincidir con engine.py
    df = df.rename(columns={
        "fecha": "Fecha",
        "jornada": "Jornada",
        "monto": "Monto",
        "attendant": "Attendant",
        "mes": "Mes",
        "hora": "Hora",
        "forma_pago": "FormaPago",
        "maquina": "Maquina"
    })
    
    # Calcular JornadaDia
    df["JornadaDia"] = pd.to_datetime(df["Jornada"]).dt.normalize()
    df["Tipo"] = "PREMIOS"
    
    return df


def _get_df_mes_anterior(get_df_fn, year: str, month: str, asistentes_seleccionados=None):
    """Devuelve el DataFrame del mes inmediatamente anterior al (year, month).
    Si year o month son 'all'/None retorna None (no hay punto de comparación claro).
    """
    if not year or year == "all" or not month or month == "all":
        return None
    try:
        y, m = int(year), int(month)
        if m == 1:
            y_prev, m_prev = y - 1, 12
        else:
            y_prev, m_prev = y, m - 1
        df_prev = get_df_fn(str(y_prev), f"{m_prev:02d}")
        if df_prev is None or df_prev.empty:
            return None
        if asistentes_seleccionados:
            df_prev = df_prev[df_prev["Attendant"].isin(asistentes_seleccionados)]
        return df_prev
    except Exception:
        return None


@app.route("/dashboard_db", methods=["GET", "POST"])
@login_required
def dashboard_db():
    years, _ = get_available_dates(Operacion)
    
    if request.method == "POST":
        # Si viene del formulario de filtros
        if "year" in request.form:
            session["year_db"] = request.form.get("year")
            session["month_db"] = request.form.get("month")
        
        # Si viene del filtro de asistentes (puede venir junto o separado)
        # El form de asistentes en dashboard.html usa 'asistentes'
        if "asistentes" in request.form or "year" in request.form:
             # Si se posteó el form, actualizamos asistentes si están presentes
             # OJO: Si el form incluye todo, request.form.getlist("asistentes") estará vacío si desmarcó todo
             # Pero si es solo cambio de año, tal vez no envíe asistentes si están en otro form.
             # En dashboard.html pondremos todo en un mismo form o manejaremos la persistencia.
             # Asumiremos que el POST viene del dashboard y trae todo.
             session["asistentes_sel_db"] = request.form.getlist("asistentes")

        return redirect(url_for("dashboard_db"))

    selected_year = session.get("year_db", "all")
    selected_month = session.get("month_db", "all")

    df = get_db_dataframe(selected_year, selected_month)
    
    asistentes_disponibles = []
    if not df.empty:
        asistentes_disponibles = sorted(df["Attendant"].dropna().unique().tolist())
    else:
        flash("No hay datos para el periodo seleccionado.", "info")

    asistentes_sel = session.get("asistentes_sel_db", [])
    asistentes_seleccionados = asistentes_sel or asistentes_disponibles

    kpis = None
    if not df.empty:
        df_filtrado = df[df["Attendant"].isin(asistentes_seleccionados)] if asistentes_seleccionados else df
        df_prev = _get_df_mes_anterior(get_db_dataframe, selected_year, selected_month, asistentes_seleccionados)
        kpis = generar_kpis(df_filtrado, df_prev)
        tablas = generar_reportes(df, asistentes_seleccionados)
    else:
        tablas = {}
    
    return render_template(
        "dashboard.html",
        file_id="db",
        tablas_html=tablas_a_html(tablas),
        asistentes_disponibles=asistentes_disponibles,
        asistentes_seleccionados=asistentes_seleccionados,
        titulo_dashboard="Histórico Getnet",
        tipo_modulo="GETNET",
        kpis=kpis,
        years=years,
        selected_year=selected_year,
        selected_month=selected_month
    )


@app.route("/dashboard_premios", methods=["GET", "POST"])
@login_required
def dashboard_premios():
    years, _ = get_available_dates(Premio)
    
    if request.method == "POST":
        if "year" in request.form:
            session["year_premios"] = request.form.get("year")
            session["month_premios"] = request.form.get("month")
        
        if "asistentes" in request.form or "year" in request.form:
            session["asistentes_sel_premios"] = request.form.getlist("asistentes")
            
        return redirect(url_for("dashboard_premios"))

    selected_year = session.get("year_premios", "all")
    selected_month = session.get("month_premios", "all")

    df = get_premios_dataframe(selected_year, selected_month)
    
    asistentes_disponibles = []
    if not df.empty:
        asistentes_disponibles = sorted(df["Attendant"].dropna().unique().tolist())
    else:
        flash("No hay datos de Premios para el periodo seleccionado.", "info")

    asistentes_sel = session.get("asistentes_sel_premios", [])
    asistentes_seleccionados = asistentes_sel or asistentes_disponibles

    kpis = None
    if not df.empty:
        df_filtrado = df[df["Attendant"].isin(asistentes_seleccionados)] if asistentes_seleccionados else df
        df_prev = _get_df_mes_anterior(get_premios_dataframe, selected_year, selected_month, asistentes_seleccionados)
        kpis = generar_kpis(df_filtrado, df_prev)
        tablas = generar_reportes(df, asistentes_seleccionados)
    else:
        tablas = {}
    
    return render_template(
        "dashboard.html",
        file_id="premios_db",
        tablas_html=tablas_a_html(tablas),
        asistentes_disponibles=asistentes_disponibles,
        asistentes_seleccionados=asistentes_seleccionados,
        titulo_dashboard="Histórico Premios",
        tipo_modulo="PREMIOS",
        kpis=kpis,
        years=years,
        selected_year=selected_year,
        selected_month=selected_month
    )


@app.route("/download/<file_id>", methods=["GET"])
@login_required
def download(file_id):
    if file_id == "db":
        selected_year = session.get("year_db", "all")
        selected_month = session.get("month_db", "all")
        df = get_db_dataframe(selected_year, selected_month)
        download_name = "reporte_historico_getnet.xlsx"
    elif file_id == "premios_db":
        selected_year = session.get("year_premios", "all")
        selected_month = session.get("month_premios", "all")
        df = get_premios_dataframe(selected_year, selected_month)
        download_name = "reporte_historico_premios.xlsx"
    else:
        df = None

    if file_id in ["db", "premios_db"]:
        if df is None or df.empty:
            return "No hay datos para descargar.", 404
            
        asistentes_disponibles = sorted(df["Attendant"].dropna().unique().tolist())
        
        # Usar la sesión correcta según el tipo
        session_key = "asistentes_sel_db" if file_id == "db" else "asistentes_sel_premios"
        asistentes_sel = session.get(session_key, [])
        asistentes_seleccionados = asistentes_sel or asistentes_disponibles
        
        tablas = generar_reportes(df, asistentes_seleccionados)
        output: BytesIO = exportar_excel_bytes(tablas)
        
        return desktop_save_response(output, download_name)

    path = safe_file_path(file_id)
    if not os.path.exists(path):
        return "Archivo no encontrado.", 404

    opciones = session.get(f"tablas_{file_id}", [])

    if request.args.get("filtered") == "true":
        asistentes_sel = request.args.getlist("asistentes")
    else:
        asistentes_sel = session.get(f"asistentes_sel_{file_id}", [])

    tablas, _ = preparar_tablas(path, opciones, asistentes_sel)
    output: BytesIO = exportar_excel_bytes(tablas)

    return desktop_save_response(output, "reporte_operaciones.xlsx")


@app.route("/graphs")
@login_required
def graphs():
    year = request.args.get("year")
    years, _ = get_available_dates(Operacion)

    df = get_db_dataframe(year=year)
    if df.empty:
        # Si no hay datos, pasamos listas vacías para que no falle el JS
        return render_template("graphs.html", 
                               data_mes={"labels": [], "ops": [], "monto": []},
                               data_hora={"labels": [], "ops": [], "monto": []},
                               years=years,
                               selected_year=year,
                               title="Dashboard de Reportes",
                               endpoint="graphs")

    # Reutilizamos la lógica de engine para agrupar
    tablas = generar_reportes(df)
    
    df_mes = tablas["Resumen Mensual"]
    df_hora = tablas["Operaciones por Hora"]

    data_mes = {
        "labels": df_mes["Mes"].tolist(),
        "ops": df_mes["Operaciones"].tolist(),
        "monto": df_mes["Monto"].tolist()
    }

    data_hora = {
        "labels": df_hora["Hora"].tolist(),
        "ops": df_hora["Operaciones"].tolist(),
        "monto": df_hora["Monto"].tolist(),
        "ops_avg": df_hora["Operaciones Promedio"].tolist(),
        "monto_avg": df_hora["Monto Promedio"].tolist()
    }

    return render_template("graphs.html", 
                           data_mes=data_mes, 
                           data_hora=data_hora, 
                           years=years, 
                           selected_year=year,
                           title="Dashboard de Reportes",
                           endpoint="graphs")


@app.route("/graphs_premios")
@login_required
def graphs_premios():
    year = request.args.get("year")
    years, _ = get_available_dates(Premio)

    df = get_premios_dataframe(year=year)
    
    # Filtrar solo Premios y Premios Progresivos
    if not df.empty:
        tipos_validos = ["jackpot hp", "progresive jackpot hp", "progressive jackpot hp"]
        df = df[df["FormaPago"].astype(str).str.lower().str.strip().isin(tipos_validos)].copy()

    if df.empty:
        return render_template("graphs.html", 
                               data_mes={"labels": [], "ops": [], "monto": []},
                               data_hora={"labels": [], "ops": [], "monto": []},
                               years=years,
                               selected_year=year,
                               title="Dashboard de Premios",
                               endpoint="graphs_premios")

    tablas = generar_reportes(df)
    
    df_mes = tablas["Resumen Mensual"]
    df_hora = tablas["Operaciones por Hora"]

    data_mes = {
        "labels": df_mes["Mes"].tolist(),
        "ops": df_mes["Operaciones"].tolist(),
        "monto": df_mes["Monto"].tolist()
    }

    data_hora = {
        "labels": df_hora["Hora"].tolist(),
        "ops": df_hora["Operaciones"].tolist(),
        "monto": df_hora["Monto"].tolist(),
        "ops_avg": df_hora["Operaciones Promedio"].tolist(),
        "monto_avg": df_hora["Monto Promedio"].tolist()
    }

    return render_template("graphs.html", 
                           data_mes=data_mes, 
                           data_hora=data_hora, 
                           years=years, 
                           selected_year=year,
                           title="Dashboard de Premios",
                           endpoint="graphs_premios")


try:
    from sgos_web.comps_routes import comps_bp
except ImportError:
    from comps_routes import comps_bp

app.register_blueprint(comps_bp)


if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_ENV") == "development")
