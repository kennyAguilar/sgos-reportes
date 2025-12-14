import os
import uuid
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session, abort
from werkzeug.utils import secure_filename

try:
    from sgos_web.engine import procesar_sgos, exportar_excel_bytes, obtener_asistentes
except ImportError:
    from engine import procesar_sgos, exportar_excel_bytes, obtener_asistentes

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "sgos-secret")  # ideal: variable de entorno

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
        k: v.to_html(index=False, classes="table table-sm table-striped")
        for k, v in tablas.items()
    }


@app.route("/", methods=["GET", "POST"])
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

        opciones = request.form.getlist("opciones")  # lo que marcó en index
        session[f"tablas_{saved_name}"] = opciones

        # OJO: NO guardamos listas grandes en session.
        # Deja que dashboard recalculé 'asistentes_disponibles' desde el archivo.
        # Guardamos solo selección (por defecto: vacío => se interpreta como "todos").
        session[f"asistentes_sel_{saved_name}"] = []

        return redirect(url_for("dashboard", file_id=saved_name))

    return render_template("index.html")


@app.route("/dashboard/<file_id>", methods=["GET", "POST"])
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


@app.route("/download/<file_id>", methods=["GET"])
def download(file_id):
    path = safe_file_path(file_id)
    if not os.path.exists(path):
        return "Archivo no encontrado.", 404

    opciones = session.get(f"tablas_{file_id}", [])
    asistentes_disponibles = obtener_asistentes(path)
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


if __name__ == "__main__":
    app.run(debug=True)
