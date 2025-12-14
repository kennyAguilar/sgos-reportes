import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter

ORDEN_HORAS = list(range(10, 24)) + list(range(0, 9))
HORAS_VALIDAS = set(ORDEN_HORAS)

COLUMNAS_CLAVE = {"Jornada", "Fecha", "Monto"}  # mínimo para validar header

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

def _autosize_sheet(ws, df, max_width=40):
    for col_idx, col_name in enumerate(df.columns, start=1):
        col_vals = df[col_name].astype(str).fillna("")
        max_len = max(len(str(col_name)), *(len(v) for v in col_vals))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, max_width)

def _detectar_fila_header(path_xlsx: str, sheet_name: str):
    preview = pd.read_excel(path_xlsx, sheet_name=sheet_name, engine="openpyxl", header=None, nrows=30)
    for i in range(len(preview)):
        fila = set(preview.iloc[i].astype(str).str.strip().tolist())
        if COLUMNAS_CLAVE.issubset(fila):
            return i
    return 0  # fallback

def _cargar_df(path_xlsx: str, sheet_name: str | None = None) -> pd.DataFrame:
    # Si no pasan hoja, usa la primera
    if sheet_name is None:
        sheet_name = pd.ExcelFile(path_xlsx, engine="openpyxl").sheet_names[0]

    header_row = _detectar_fila_header(path_xlsx, sheet_name)

    df = pd.read_excel(path_xlsx, sheet_name=sheet_name, engine="openpyxl", header=header_row)

    df = df.rename(columns={
        "Id Cliente": "IdCliente",
        "Slot Attendant": "Attendant",
        "Forma Pago": "FormaPago",
        "Ingreso CAWA": "Ingreso",
    })

    df["Fecha"] = pd.to_datetime(df.get("Fecha"), errors="coerce", dayfirst=True)
    df["Jornada"] = pd.to_datetime(df.get("Jornada"), errors="coerce", dayfirst=True)
    df["Monto"] = pd.to_numeric(df.get("Monto"), errors="coerce").fillna(0)

    df = df.dropna(subset=["Fecha", "Jornada", "Attendant"]).copy()

    df["JornadaDia"] = df["Jornada"].dt.normalize()
    df["Hora"] = df["Fecha"].dt.hour

    # Excluir 09 y cualquier otra fuera de jornada
    df = df[df["Hora"].isin(HORAS_VALIDAS)].copy()

    df["Mes"] = df["JornadaDia"].dt.to_period("M").astype(str)
    return df

def procesar_sgos(path_xlsx: str, sheet_name: str | None = None, asistentes_filtro: list = None):
    df = _cargar_df(path_xlsx, sheet_name=sheet_name)

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

    tabla_asistente_mes = (
        df.groupby(["Attendant", "Mes"], as_index=False)
          .agg(Operaciones=("Monto", "count"), Monto=("Monto", "sum"))
          .sort_values(["Mes", "Operaciones"], ascending=[True, False])
    )
    tabla_asistente_mes["Mes"] = tabla_asistente_mes["Mes"].apply(_formatear_periodo)

    qa_df = pd.DataFrame([
        ["filas_usadas", len(df)],
        ["min_fecha", str(df["Fecha"].min())],
        ["max_fecha", str(df["Fecha"].max())],
        ["horas_presentes", ", ".join(map(str, sorted(df["Hora"].unique())))],
    ], columns=["Metrica", "Valor"])

    return {
        "Resumen Mensual": tabla_mes,
        "Operaciones por Hora": tabla_hora,
        "Record Asistentes": tabla_record,
        "Asistente por Mes": tabla_asistente_mes,
        "QA": qa_df,
    }

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
