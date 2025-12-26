import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter

ORDEN_HORAS = list(range(10, 24)) + list(range(0, 9))
HORAS_VALIDAS = set(ORDEN_HORAS)

COLUMNAS_CLAVE_STD = {"Jornada", "Fecha", "Monto"}
COLUMNAS_CLAVE_PREMIOS = {"Monto Transferido", "Slot Attendant", "Transferencia Final"}

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

def _formatear_periodo(periodo_str: str) -> str:
    # Espera formato "YYYY-MM"
    try:
        anio, mes = periodo_str.split("-")
        nombre_mes = MESES_ES.get(int(mes), mes)
        return f"{nombre_mes} {anio}"
    except:
        return periodo_str

def _autosize_sheet(ws, df, max_width=None):
    for col_idx, col_name in enumerate(df.columns, start=1):
        # Detectar si es fecha para dar más margen
        is_date = pd.api.types.is_datetime64_any_dtype(df[col_name])
        
        col_vals = df[col_name].astype(str).fillna("")
        max_len = max(len(str(col_name)), *(len(v) for v in col_vals))
        
        # Las fechas suelen necesitar más espacio por el formato (dd/mm/yyyy, etc.)
        padding = 6 if is_date else 3
        
        adjusted_width = max_len + padding
        if max_width:
            adjusted_width = min(adjusted_width, max_width)
        ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

def _detectar_fila_header(path_xlsx: str, sheet_name: str):
    preview = pd.read_excel(path_xlsx, sheet_name=sheet_name, engine="openpyxl", header=None, nrows=30)
    for i in range(len(preview)):
        fila = set(preview.iloc[i].astype(str).str.strip().tolist())
        if COLUMNAS_CLAVE_STD.issubset(fila):
            return i
        if COLUMNAS_CLAVE_PREMIOS.issubset(fila):
            return i
    return 0  # fallback

def _cargar_df(path_xlsx: str, sheet_name: str | None = None) -> pd.DataFrame:
    # Si no pasan hoja, usa la primera
    if sheet_name is None:
        sheet_name = pd.ExcelFile(path_xlsx, engine="openpyxl").sheet_names[0]

    header_row = _detectar_fila_header(path_xlsx, sheet_name)

    df = pd.read_excel(path_xlsx, sheet_name=sheet_name, engine="openpyxl", header=header_row)

    # --- Lógica para PREMIOS ---
    if "Transferencia Final" in df.columns and "Slot Attendant" in df.columns:
        # Si no existe columna 'Fecha' explícita, asumimos la primera columna (A)
        if "Fecha" not in df.columns:
            df.rename(columns={df.columns[0]: "Fecha"}, inplace=True)
        
        # Si no existe 'Jornada', la calculamos (Fecha - 10h para ajustar día operativo)
        if "Jornada" not in df.columns:
            # Convertimos temporalmente para calcular
            fechas_dt = pd.to_datetime(df["Fecha"], dayfirst=True, errors='coerce')
            # Asumimos inicio de jornada a las 10:00 AM (restamos 10h)
            df["Jornada"] = (fechas_dt - pd.Timedelta(hours=10)).dt.normalize()

        df = df.rename(columns={
            "Cliente": "IdCliente",
            "Transferencia Final": "Monto",
            "Tipo de Pago": "FormaPago"
        })
        df["Tipo"] = "PREMIOS"
    else:
        df["Tipo"] = "GETNET"
    # ---------------------------

    df = df.rename(columns={
        "Id Cliente": "IdCliente",
        "Slot Attendant": "Attendant",
        "Forma Pago": "FormaPago",
        "Ingreso CAWA": "Ingreso",
    })

    # Usar format='mixed' para soportar tanto DD-MM-YYYY como YYYY-MM-DD correctamente
    df["Fecha"] = pd.to_datetime(df.get("Fecha"), errors="coerce", dayfirst=True, format="mixed")
    df["Jornada"] = pd.to_datetime(df.get("Jornada"), errors="coerce", dayfirst=True, format="mixed")
    df["Monto"] = pd.to_numeric(df.get("Monto"), errors="coerce").fillna(0)

    df = df.dropna(subset=["Fecha", "Jornada", "Attendant"]).copy()

    df["JornadaDia"] = df["Jornada"].dt.normalize()
    df["Hora"] = df["Fecha"].dt.hour

    # Excluir 09 y cualquier otra fuera de jornada
    df = df[df["Hora"].isin(HORAS_VALIDAS)].copy()

    df["Mes"] = df["JornadaDia"].dt.to_period("M").astype(str)
    return df

def guardar_datos_db(path_xlsx: str, db, OperacionModel, PremioModel, sheet_name: str | None = None):
    """
    Lee el Excel, detecta si es Getnet o Premios, y guarda en la tabla correspondiente.
    """
    df = _cargar_df(path_xlsx, sheet_name=sheet_name)
    
    if df.empty:
        return 0, "No data"

    # Detectar meses presentes en el archivo
    meses_en_archivo = df["Mes"].unique()
    
    # Detectar tipo de archivo
    tipo_archivo = df["Tipo"].iloc[0] if "Tipo" in df.columns else "GETNET"
    
    TargetModel = PremioModel if tipo_archivo == "PREMIOS" else OperacionModel

    # Iniciar transacción
    try:
        for mes in meses_en_archivo:
            # Borrar datos existentes de ese mes en la tabla correspondiente
            # Si el modelo tiene columna 'tipo', filtramos también por tipo para no borrar otros
            query = db.session.query(TargetModel).filter(TargetModel.mes == mes)
            if hasattr(TargetModel, 'tipo'):
                query = query.filter(TargetModel.tipo == tipo_archivo)
            query.delete()
        
        # Insertar nuevos datos
        registros = []
        for _, row in df.iterrows():
            if tipo_archivo == "PREMIOS":
                reg = PremioModel(
                    fecha=row["Fecha"],
                    jornada=row["Jornada"],
                    id_cliente=str(row.get("IdCliente", "")),
                    monto=row["Monto"],
                    propina=row.get("Propina", 0),
                    maquina=str(row.get("Máquina", "") or row.get("Maquina", "")),
                    attendant=row["Attendant"],
                    validador=str(row.get("Validador", "")),
                    forma_pago=str(row.get("FormaPago", "")),
                    ingreso_cawa=str(row.get("Ingreso", "")),
                    mes=row["Mes"],
                    hora=row["Hora"]
                )
            else:
                reg = OperacionModel(
                    fecha=row["Fecha"],
                    jornada=row["Jornada"],
                    id_cliente=str(row.get("IdCliente", "")),
                    monto=row["Monto"],
                    voucher=str(row.get("Voucher", "")),
                    attendant=row["Attendant"],
                    validador=str(row.get("Validador", "")),
                    forma_pago=str(row.get("FormaPago", "")),
                    ingreso_cawa=str(row.get("Ingreso", "")),
                    mes=row["Mes"],
                    hora=row["Hora"]
                )
            registros.append(reg)
        
        db.session.add_all(registros)
        db.session.commit()
        return len(registros), tipo_archivo
    except Exception as e:
        db.session.rollback()
        raise e

def generar_reportes(df: pd.DataFrame, asistentes_filtro: list = None) -> dict:
    """
    Genera los diccionarios de DataFrames (tablas) a partir de un DataFrame principal ya limpio.
    """
    if asistentes_filtro:
        df = df[df["Attendant"].isin(asistentes_filtro)].copy()

    tabla_mes = (
        df.groupby("Mes", as_index=False)
          .agg(Operaciones=("Monto", "count"), Monto=("Monto", "sum"))
          .sort_values("Mes")
    )
    tabla_mes["Mes"] = tabla_mes["Mes"].apply(_formatear_periodo)

    tabla_hora = (
        df.groupby("Hora", as_index=False)
          .agg(Operaciones=("Monto", "count"), Monto=("Monto", "sum"))
        .set_index("Hora")
        .reindex(ORDEN_HORAS, fill_value=0)
        .reset_index()
    )
    # Convertir a string para que Excel lo trate como categorías (texto) y no números
    tabla_hora["Hora"] = tabla_hora["Hora"].astype(str)

    ops_por_jornada = (
        df.groupby(["Attendant", "JornadaDia"], as_index=False)
          .size()
          .rename(columns={"size": "TotalOperaciones"})
    )

    if len(ops_por_jornada) > 0:
        idx_max = ops_por_jornada.groupby("Attendant")["TotalOperaciones"].idxmax()
        tabla_record = (
            ops_por_jornada.loc[idx_max]
              .sort_values("TotalOperaciones", ascending=False)
              .reset_index(drop=True)
        )
    else:
        tabla_record = ops_por_jornada

    # Configurar agregación: si es Premios, NO mostramos Monto
    es_premios = False
    if not df.empty and "Tipo" in df.columns:
        es_premios = (df["Tipo"].iloc[0] == "PREMIOS")

    agg_config = {"Operaciones": ("Monto", "count")}
    if not es_premios:
        agg_config["Monto"] = ("Monto", "sum")

    tabla_asistente_mes = (
        df.groupby(["Attendant", "Mes"], as_index=False)
          .agg(**agg_config)
          .sort_values(["Mes", "Operaciones"], ascending=[True, False])
    )
    tabla_asistente_mes["Mes"] = tabla_asistente_mes["Mes"].apply(_formatear_periodo)

    # Para Conteo Operaciones:
    if es_premios and "FormaPago" in df.columns:
        # Crear copia para no afectar el df original y normalizar
        df_p = df.copy()
        df_p['FormaPagoNorm'] = df_p['FormaPago'].astype(str).str.lower().str.strip()
        
        def classify_payment(val):
            if val in ["jackpot hp", "progresive jackpot hp", "progressive jackpot hp"]:
                return "Premios"
            elif val == "mdc purse clear":
                return "MDC purse clear"
            elif val == "cancel credit":
                return "Cancel Credit"
            elif val == "chip cash handpay":
                return "Chip Cash HandPay"
            return None

        df_p['Categoria'] = df_p['FormaPagoNorm'].apply(classify_payment)
        
        # Pivot table para contar por categoría
        # Usamos 'Monto' como columna dummy para contar (count)
        tabla_conteo_ops = pd.pivot_table(
            df_p[df_p['Categoria'].notna()],
            index=["Mes", "Attendant"],
            columns="Categoria",
            values="Monto", 
            aggfunc="count",
            fill_value=0
        ).reset_index()
        
        # Asegurar que existan todas las columnas deseadas (rellenar con 0 si no hay datos)
        cols_deseadas = ["Premios", "MDC purse clear", "Cancel Credit", "Chip Cash HandPay"]
        for col in cols_deseadas:
            if col not in tabla_conteo_ops.columns:
                tabla_conteo_ops[col] = 0
        
        # Ordenar columnas: Mes, Attendant, y luego las categorías en el orden pedido
        tabla_conteo_ops = tabla_conteo_ops[["Mes", "Attendant"] + cols_deseadas]
        
        # Ordenar filas: por Mes y luego por cantidad de Premios (descendente)
        tabla_conteo_ops = tabla_conteo_ops.sort_values(["Mes", "Premios"], ascending=[True, False])
        
        # Eliminar nombre del índice de columnas para limpieza
        tabla_conteo_ops.columns.name = None

        # --- NUEVA TABLA: Total de conteo anual por asistente (Detallado) ---
        tabla_conteo_anual = pd.pivot_table(
            df_p[df_p['Categoria'].notna()],
            index=["Attendant"],
            columns="Categoria",
            values="Monto", 
            aggfunc="count",
            fill_value=0
        ).reset_index()
        
        for col in cols_deseadas:
            if col not in tabla_conteo_anual.columns:
                tabla_conteo_anual[col] = 0
                
        tabla_conteo_anual = tabla_conteo_anual[["Attendant"] + cols_deseadas]
        tabla_conteo_anual = tabla_conteo_anual.sort_values(["Premios"], ascending=False)
        tabla_conteo_anual.columns.name = None

        # --- NUEVA TABLA: Conteo Total Anual (Simple) ---
        tabla_conteo_anual_total = (
            df_p[df_p['Categoria'].notna()]
            .groupby("Attendant", as_index=False)
            .agg(Operaciones=("Monto", "count"))
            .sort_values("Operaciones", ascending=False)
        )

        # --- NUEVA TABLA: Conteo de operaciones por MDA ---
        # Mes | Maquina | cantidad de premios (jackpot + progresive) | monto | cantidad de MDC Purse Clear | Cancel credit | Chip Cash HandPay
        
        # Usamos el mismo df_p que ya tiene 'Categoria' y 'FormaPagoNorm'
        # Pero necesitamos calcular Monto SOLO para Premios.
        # Creamos columna auxiliar MontoPremios
        df_p["MontoPremios"] = df_p.apply(lambda x: x["Monto"] if x["Categoria"] == "Premios" else 0, axis=1)
        
        # Agrupamos por Mes y Maquina
        # Calculamos conteos por categoría y suma de MontoPremios
        
        # Pivot para conteos
        pivot_counts = pd.pivot_table(
            df_p[df_p['Categoria'].notna()],
            index=["Mes", "Maquina"],
            columns="Categoria",
            values="Monto", # Dummy para count
            aggfunc="count",
            fill_value=0
        ).reset_index()
        
        # Suma de montos de premios
        monto_premios = df_p.groupby(["Mes", "Maquina"])["MontoPremios"].sum().reset_index()
        
        # Merge
        tabla_conteo_mda = pd.merge(pivot_counts, monto_premios, on=["Mes", "Maquina"], how="left")
        
        # Asegurar columnas
        for col in cols_deseadas:
            if col not in tabla_conteo_mda.columns:
                tabla_conteo_mda[col] = 0
                
        # Renombrar y ordenar
        # Queremos: Mes, Maquina, Premios, Monto (de premios), MDC, Cancel, Chip
        tabla_conteo_mda = tabla_conteo_mda.rename(columns={"MontoPremios": "Monto"})
        
        cols_finales_mda = ["Mes", "Maquina", "Premios", "Monto", "MDC purse clear", "Cancel Credit", "Chip Cash HandPay"]
        tabla_conteo_mda = tabla_conteo_mda[cols_finales_mda]
        
        # Ordenar
        tabla_conteo_mda = tabla_conteo_mda.sort_values(["Mes", "Premios"], ascending=[True, False])
        tabla_conteo_mda["Mes"] = tabla_conteo_mda["Mes"].apply(_formatear_periodo)
        tabla_conteo_mda.columns.name = None

        # --- NUEVA TABLA: Conteo total de operaciones por MDA (Acumulado) ---
        # Maquina | cantidad de premios (jackpot + progresive) | monto | cantidad de MDC Purse Clear | Cancel credit | Chip Cash HandPay
        
        # Pivot para conteos totales (sin agrupar por Mes)
        pivot_counts_total = pd.pivot_table(
            df_p[df_p['Categoria'].notna()],
            index=["Maquina"],
            columns="Categoria",
            values="Monto", 
            aggfunc="count",
            fill_value=0
        ).reset_index()
        
        # Suma de montos de premios totales
        monto_premios_total = df_p.groupby(["Maquina"])["MontoPremios"].sum().reset_index()
        
        # Merge total
        tabla_conteo_mda_total = pd.merge(pivot_counts_total, monto_premios_total, on=["Maquina"], how="left")
        
        # Asegurar columnas
        for col in cols_deseadas:
            if col not in tabla_conteo_mda_total.columns:
                tabla_conteo_mda_total[col] = 0
                
        # Renombrar y ordenar
        tabla_conteo_mda_total = tabla_conteo_mda_total.rename(columns={"MontoPremios": "Monto"})
        
        # Columnas finales (sin Mes)
        cols_finales_mda_total = ["Maquina", "Premios", "Monto", "MDC purse clear", "Cancel Credit", "Chip Cash HandPay"]
        tabla_conteo_mda_total = tabla_conteo_mda_total[cols_finales_mda_total]
        
        # Ordenar por Premios descendente
        tabla_conteo_mda_total = tabla_conteo_mda_total.sort_values(["Premios"], ascending=False)
        tabla_conteo_mda_total.columns.name = None

    else:
        # Lógica original para Getnet
        tabla_conteo_ops = (
            df.groupby(["Mes", "Attendant"], as_index=False)
              .agg(Operaciones=("Monto", "count"))
              .sort_values(["Mes", "Operaciones"], ascending=[True, False])
        )
        tabla_conteo_anual = (
            df.groupby(["Attendant"], as_index=False)
              .agg(Operaciones=("Monto", "count"))
              .sort_values(["Operaciones"], ascending=False)
        )
        tabla_conteo_anual_total = tabla_conteo_anual.copy()
        tabla_conteo_mda = pd.DataFrame() # Vacía para Getnet
        tabla_conteo_mda_total = pd.DataFrame() # Vacía para Getnet
    
    tabla_conteo_ops["Mes"] = tabla_conteo_ops["Mes"].apply(_formatear_periodo)

    qa_df = pd.DataFrame([
        ["filas_usadas", len(df)],
        ["min_fecha", str(df["Fecha"].min())],
        ["max_fecha", str(df["Fecha"].max())],
        ["horas_presentes", ", ".join(map(str, sorted(df["Hora"].unique())))],
    ], columns=["Metrica", "Valor"])

    reportes = {
        "Resumen Mensual": tabla_mes,
        "Operaciones por Hora": tabla_hora,
        "Record Asistentes": tabla_record,
        "Asistente por Mes": tabla_asistente_mes,
        "Conteo Operaciones": tabla_conteo_ops,
        "Total de conteo anual por asistente": tabla_conteo_anual,
        "Conteo Total Anual": tabla_conteo_anual_total,
    }
    
    if es_premios:
        reportes["Conteo mensual de operaciones por MDA"] = tabla_conteo_mda
        reportes["Conteo total de operaciones por MDA"] = tabla_conteo_mda_total
        
    reportes["QA"] = qa_df
    
    return reportes

def procesar_sgos(path_xlsx: str, sheet_name: str | None = None, asistentes_filtro: list = None):
    df = _cargar_df(path_xlsx, sheet_name=sheet_name)
    return generar_reportes(df, asistentes_filtro)

def obtener_asistentes(path_xlsx: str, sheet_name: str | None = None) -> list:
    # print(f"DEBUG: obtener_asistentes called with path={path_xlsx}, sheet_name={sheet_name}")
    df = _cargar_df(path_xlsx, sheet_name=sheet_name)
    return sorted(df["Attendant"].dropna().unique().tolist())

def exportar_excel_bytes(tablas: dict) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, df in tablas.items():
            sheet_name = str(sheet)[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.book[sheet_name]
            _autosize_sheet(ws, df)
    output.seek(0)
    return output
