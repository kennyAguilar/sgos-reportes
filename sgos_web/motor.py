import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter

# Jornada válida Dreams: 10 -> 23 -> 00 -> 08 (NO 09)
ORDEN_HORAS = list(range(10, 24)) + list(range(0, 9))
HORAS_VALIDAS = set(ORDEN_HORAS)

def _autosize_sheet(ws, df, max_width=40):
    for col_idx, col_name in enumerate(df.columns, start=1):
        # cuidado con NaN
        col_vals = df[col_name].astype(str).fillna("")
        max_len = max(len(str(col_name)), *(len(v) for v in col_vals))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, max_width)

def procesar_sgos(path_xlsx: str, sheet_name: str = "Data anual", asistentes_filtro: list = None):
    raw = pd.read_excel(path_xlsx, sheet_name=sheet_name, engine="openpyxl")
    raw.columns = raw.iloc[0]          # primera fila = encabezados reales
    df = raw.iloc[1:].copy()

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

    # Excluir horas fuera de jornada (incluye 09)
    df = df[df["Hora"].isin(HORAS_VALIDAS)].copy()
    
    # Filtrar por asistentes si se especifican
    if asistentes_filtro:
        df = df[df["Attendant"].isin(asistentes_filtro)].copy()

    df["Mes"] = df["JornadaDia"].dt.to_period("M").astype(str)

    # TABLA 1: Mes
    tabla_mes = (
        df.groupby("Mes")
          .agg(Operaciones=("Monto", "count"), Monto=("Monto", "sum"))
          .reset_index()
          .sort_values("Mes")
    )

    # TABLA 2: Hora
    tabla_hora = (
        df.groupby("Hora")
          .agg(Operaciones=("Monto", "count"), Monto=("Monto", "sum"))
          .reindex(ORDEN_HORAS, fill_value=0)
          .reset_index()
    )

    # TABLA 3: Récord por asistente (máximo en una jornada)
    ops_por_jornada = (
        df.groupby(["Attendant", "JornadaDia"])
          .size()
          .reset_index(name="TotalOperaciones")
    )
    idx_max = ops_por_jornada.groupby("Attendant")["TotalOperaciones"].idxmax()
    tabla_record = (
        ops_por_jornada.loc[idx_max]
          .sort_values("TotalOperaciones", ascending=False)
          .reset_index(drop=True)
    )

    # TABLA 4: Asistente por mes
    tabla_asistente_mes = (
        df.groupby(["Attendant", "Mes"])
          .agg(Operaciones=("Monto", "count"), Monto=("Monto", "sum"))
          .reset_index()
          .sort_values(["Mes", "Operaciones"], ascending=[True, False])
    )

    # QA mínimo
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

def exportar_excel_bytes(tablas: dict) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, df in tablas.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
            ws = writer.book[sheet[:31]]
            _autosize_sheet(ws, df)
    output.seek(0)
    return output

def obtener_asistentes(path_xlsx: str, sheet_name: str = "Data anual") -> list:
    """Extrae la lista única de asistentes del archivo"""
    raw = pd.read_excel(path_xlsx, sheet_name=sheet_name, engine="openpyxl")
    raw.columns = raw.iloc[0]
    df = raw.iloc[1:].copy()
    
    df = df.rename(columns={"Slot Attendant": "Attendant"})
    asistentes = sorted(df["Attendant"].dropna().unique().tolist())
    return asistentes
