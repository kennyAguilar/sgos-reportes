import os
import uuid
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session
from werkzeug.utils import secure_filename

from sgos_web.motor import procesar_sgos, exportar_excel_bytes, obtener_asistentes

app = Flask(__name__)
app.secret_key = "sgos-secret"  # simple para flash messages

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
ALLOWED_EXT = {".xlsx", ".xls"}

def allowed_file(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXT

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        f = request.files.get("file")
        if not f or f.filename == "":
            flash("No se subió ningún archivo.")
            return redirect(url_for("index"))

        if not allowed_file(f.filename):
            flash("Formato no permitido. Sube un .xlsx")
            return redirect(url_for("index"))

        filename = secure_filename(f.filename)
        token = uuid.uuid4().hex
        saved_name = f"{token}__{filename}"
        path = os.path.join(app.config["UPLOAD_FOLDER"], saved_name)
        f.save(path)

        opciones = request.form.getlist("opciones")
        session[f"tablas_{saved_name}"] = opciones
        
        asistentes = obtener_asistentes(path)
        session[f"asistentes_disponibles_{saved_name}"] = asistentes
        session[f"asistentes_seleccionados_{saved_name}"] = asistentes
        
        return redirect(url_for("dashboard", file_id=saved_name))

    return render_template("index.html")

@app.route("/dashboard/<file_id>", methods=["GET", "POST"])
def dashboard(file_id):
    path = os.path.join(app.config["UPLOAD_FOLDER"], file_id)
    if not os.path.exists(path):
        return "Archivo no encontrado.", 404
    
    if request.method == "POST":
        asistentes_seleccionados = request.form.getlist("asistentes")
        session[f"asistentes_seleccionados_{file_id}"] = asistentes_seleccionados
        return redirect(url_for("dashboard", file_id=file_id))

    tablas = procesar_sgos(path)
    
    opciones = session.get(f"tablas_{file_id}", [])
    asistentes_seleccionados = session.get(f"asistentes_seleccionados_{file_id}", [])
    asistentes_disponibles = session.get(f"asistentes_disponibles_{file_id}", [])
    
    if opciones:
        tablas = {k: v for k, v in tablas.items() if k in opciones}
    
    # Filtrar por asistentes, PERO NO para "Resumen Mensual" ni "Operaciones por Hora"
    if asistentes_seleccionados:
        tablas_filtradas = procesar_sgos(path, asistentes_filtro=asistentes_seleccionados)
        if opciones:
            tablas_filtradas = {k: v for k, v in tablas_filtradas.items() if k in opciones}
        
        # Mantener las tablas no filtradas
        tablas_no_filtrar = ["Resumen Mensual", "Operaciones por Hora"]
        for tabla in tablas_no_filtrar:
            if tabla in tablas and tabla not in tablas_filtradas:
                tablas_filtradas[tabla] = tablas[tabla]
        
        tablas = tablas_filtradas

    # Convertimos a HTML para mostrar en pantalla
    tablas_html = {k: v.to_html(index=False, classes="table table-sm table-striped") for k, v in tablas.items()}

    return render_template(
        "dashboard.html",
        file_id=file_id,
        tablas_html=tablas_html,
        asistentes_disponibles=asistentes_disponibles,
        asistentes_seleccionados=asistentes_seleccionados
    )

@app.route("/download/<file_id>", methods=["GET"])
def download(file_id):
    path = os.path.join(app.config["UPLOAD_FOLDER"], file_id)
    if not os.path.exists(path):
        return "Archivo no encontrado.", 404

    opciones = session.get(f"tablas_{file_id}", [])
    asistentes_seleccionados = session.get(f"asistentes_seleccionados_{file_id}", [])
    
    # Procesar todas las tablas primero
    tablas = procesar_sgos(path)
    
    # Filtrar por asistentes, PERO NO para "Resumen Mensual" ni "Operaciones por Hora"
    if asistentes_seleccionados:
        tablas_filtradas = procesar_sgos(path, asistentes_filtro=asistentes_seleccionados)
        
        # Mantener las tablas no filtradas
        tablas_no_filtrar = ["Resumen Mensual", "Operaciones por Hora"]
        for tabla in tablas_no_filtrar:
            if tabla in tablas:
                tablas_filtradas[tabla] = tablas[tabla]
        
        tablas = tablas_filtradas
    
    if opciones:
        tablas = {k: v for k, v in tablas.items() if k in opciones}
    
    output = exportar_excel_bytes(tablas)

    return send_file(
        output,
        as_attachment=True,
        download_name="reporte_operaciones.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True)
