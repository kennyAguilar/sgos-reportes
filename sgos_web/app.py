import os
import uuid
from io import BytesIO
import pandas as pd
from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session, abort
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import select
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user

load_dotenv()  # Carga las variables del archivo .env

try:
    from sgos_web.engine import procesar_sgos, exportar_excel_bytes, obtener_asistentes, guardar_datos_db, generar_reportes
except ImportError:
    from engine import procesar_sgos, exportar_excel_bytes, obtener_asistentes, guardar_datos_db, generar_reportes

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "sgos-secret")

# Configuración de Login
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"

# Configuración de Base de Datos
# Si no hay variable DATABASE_URL o está vacía, usa SQLite local por defecto
db_url = os.environ.get("DATABASE_URL")
if not db_url:
    db_url = "sqlite:///sgos_local.db"

app.config['SQLALCHEMY_DATABASE_URI'] = db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

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

    def __repr__(self):
        return f"<Premio {self.id} - {self.attendant} - {self.monto}>"

# Crear tablas si no existen (solo para desarrollo local/inicial)
with app.app_context():
    db.create_all()
    
    # Crear usuario admin por defecto si no existe
    if not User.query.filter_by(username="admin").first():
        admin = User(username="admin")
        admin.set_password("admin123")  # Contraseña por defecto
        db.session.add(admin)
        db.session.commit()
        print("Usuario 'admin' creado con contraseña 'admin123'")

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB (ajusta si quieres)

ALLOWED_EXT = {".xlsx", ".xls"}
TABLAS_NO_FILTRAR = {"Resumen Mensual", "Operaciones por Hora"}


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


def preparar_tablas(path: str, opciones: list[str], asistentes_sel: list[str]) -> dict:
    """
    Regla:
    - Si hay filtro de asistentes: filtra todas las tablas EXCEPTO las de TABLAS_NO_FILTRAR
    - Si NO hay filtro: todo sin filtrar
    - Aplica 'opciones' al final
    """
    tablas_base = procesar_sgos(path)  # 1 vez siempre

    # Si no hay selección o viene vacío, devolvemos base con opciones
    if not asistentes_sel:
        return aplicar_opciones(tablas_base, opciones)

    # Si seleccionaron TODOS, no hace falta reprocesar filtrado
    asistentes_disponibles = obtener_asistentes(path)
    if set(asistentes_sel) == set(asistentes_disponibles):
        return aplicar_opciones(tablas_base, opciones)

    tablas_filtradas = procesar_sgos(path, asistentes_filtro=asistentes_sel)  # 2da (solo si aplica)

    # Forzar que ciertas tablas queden sin filtro
    for nombre in TABLAS_NO_FILTRAR:
        if nombre in tablas_base:
            tablas_filtradas[nombre] = tablas_base[nombre]

    return aplicar_opciones(tablas_filtradas, opciones)


def tablas_a_html(tablas: dict) -> dict:
    return {
        k: v.to_html(index=False, classes="table table-sm table-striped w-auto mx-auto")
        for k, v in tablas.items()
    }


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("index"))
        
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")
        user = User.query.filter_by(username=username).first()
        
        if user and user.check_password(password):
            login_user(user)
            return redirect(url_for("index"))
        else:
            flash("Usuario o contraseña incorrectos.")
            
    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


@app.route("/", methods=["GET", "POST"])
@login_required
def index():
    if request.method == "POST":
        f = request.files.get("file")
        if not f or f.filename == "":
            flash("No se subió ningún archivo.")
            return redirect(url_for("index"))

        if not allowed_file(f.filename):
            flash("Formato no permitido. Sube un .xlsx o .xls")
            return redirect(url_for("index"))

        filename = secure_filename(f.filename)
        token = uuid.uuid4().hex
        saved_name = f"{token}__{filename}"
        path = os.path.join(app.config["UPLOAD_FOLDER"], saved_name)
        f.save(path)

        # --- NUEVO: Guardar en Base de Datos ---
        try:
            total_guardados, tipo_archivo = guardar_datos_db(path, db, Operacion, Premio)
            flash(f"¡Éxito! Se guardaron {total_guardados} registros de tipo {tipo_archivo} en la base de datos.")
        except Exception as e:
            flash(f"Error al guardar en base de datos: {str(e)}")

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

    asistentes_disponibles = obtener_asistentes(path)

    if request.method == "POST":
        asistentes_sel = request.form.getlist("asistentes")
        # Si el usuario no marca nada, lo tratamos como "todos" (vacío)
        # Si prefieres lo contrario, cámbialo.
        session[f"asistentes_sel_{file_id}"] = asistentes_sel
        return redirect(url_for("dashboard", file_id=file_id))

    opciones = session.get(f"tablas_{file_id}", [])
    asistentes_sel = session.get(f"asistentes_sel_{file_id}", [])

    # Si no hay selección guardada, mostramos todos marcados por defecto
    asistentes_seleccionados = asistentes_sel or asistentes_disponibles

    tablas = preparar_tablas(path, opciones, asistentes_seleccionados)
    return render_template(
        "dashboard.html",
        file_id=file_id,
        tablas_html=tablas_a_html(tablas),
        asistentes_disponibles=asistentes_disponibles,
        asistentes_seleccionados=asistentes_seleccionados
    )


def get_db_dataframe():
    """Consulta la base de datos y devuelve un DataFrame con el formato esperado por engine.py"""
    with db.engine.connect() as conn:
        df = pd.read_sql(select(Operacion), conn)
    
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
    
    return df


def get_premios_dataframe():
    """Consulta la base de datos de PREMIOS y devuelve un DataFrame"""
    with db.engine.connect() as conn:
        df = pd.read_sql(select(Premio), conn)
    
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
    
    return df


@app.route("/dashboard_db", methods=["GET", "POST"])
@login_required
def dashboard_db():
    df = get_db_dataframe()
    
    if df.empty:
        flash("No hay datos de Getnet en la base de datos.")
        return redirect(url_for("index"))

    asistentes_disponibles = sorted(df["Attendant"].dropna().unique().tolist())

    if request.method == "POST":
        asistentes_sel = request.form.getlist("asistentes")
        session["asistentes_sel_db"] = asistentes_sel
        return redirect(url_for("dashboard_db"))

    asistentes_sel = session.get("asistentes_sel_db", [])
    asistentes_seleccionados = asistentes_sel or asistentes_disponibles

    # Generar reportes usando el DataFrame directo
    tablas = generar_reportes(df, asistentes_seleccionados)
    
    return render_template(
        "dashboard.html",
        file_id="db",
        tablas_html=tablas_a_html(tablas),
        asistentes_disponibles=asistentes_disponibles,
        asistentes_seleccionados=asistentes_seleccionados,
        titulo_dashboard="Histórico Getnet"
    )


@app.route("/dashboard_premios", methods=["GET", "POST"])
@login_required
def dashboard_premios():
    df = get_premios_dataframe()
    
    if df.empty:
        flash("No hay datos de Premios en la base de datos.")
        return redirect(url_for("index"))

    asistentes_disponibles = sorted(df["Attendant"].dropna().unique().tolist())

    if request.method == "POST":
        asistentes_sel = request.form.getlist("asistentes")
        session["asistentes_sel_premios"] = asistentes_sel
        return redirect(url_for("dashboard_premios"))

    asistentes_sel = session.get("asistentes_sel_premios", [])
    asistentes_seleccionados = asistentes_sel or asistentes_disponibles

    # Generar reportes usando el DataFrame directo
    tablas = generar_reportes(df, asistentes_seleccionados)
    
    return render_template(
        "dashboard.html",
        file_id="premios_db",
        tablas_html=tablas_a_html(tablas),
        asistentes_disponibles=asistentes_disponibles,
        asistentes_seleccionados=asistentes_seleccionados,
        titulo_dashboard="Histórico Premios"
    )


@login_required
@app.route("/download/<file_id>", methods=["GET"])
def download(file_id):
    if file_id == "db":
        df = get_db_dataframe()
        download_name = "reporte_historico_getnet.xlsx"
    elif file_id == "premios_db":
        df = get_premios_dataframe()
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
        
        return send_file(
            output,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    path = safe_file_path(file_id)
    if not os.path.exists(path):
        return "Archivo no encontrado.", 404

    opciones = session.get(f"tablas_{file_id}", [])
    asistentes_disponibles = obtener_asistentes(path)

    # Priorizar filtro desde URL (si viene del botón con JS)
    if request.args.get("filtered") == "true":
        asistentes_sel = request.args.getlist("asistentes")
    else:
        asistentes_sel = session.get(f"asistentes_sel_{file_id}", [])

    asistentes_seleccionados = asistentes_sel or asistentes_disponibles

    tablas = preparar_tablas(path, opciones, asistentes_seleccionados)
    output: BytesIO = exportar_excel_bytes(tablas)

    return send_file(
        output,
        as_attachment=True,
        download_name="reporte_operaciones.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/premios")
@login_required
def premios():
    return render_template("construccion.html")


@app.route("/graphs")
@login_required
def graphs():
    df = get_db_dataframe()
    if df.empty:
        # Si no hay datos, pasamos listas vacías para que no falle el JS
        return render_template("graphs.html", 
                               data_mes={"labels": [], "ops": [], "monto": []},
                               data_hora={"labels": [], "ops": [], "monto": []})

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
        "monto": df_hora["Monto"].tolist()
    }

    return render_template("graphs.html", data_mes=data_mes, data_hora=data_hora)


if __name__ == "__main__":
    app.run(debug=True)
