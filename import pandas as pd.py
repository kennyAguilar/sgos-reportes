import pandas as pd
from openpyxl.utils import get_column_letter

# ===============================
#   CONFIG
# ===============================
ARCHIVO = "SGOS - Dreams Punta Arenas.xlsx"
HOJA_DATOS = "Data anual"
OUTPUT_FILE = "reporte_operaciones_ok.xlsx"

# Jornada válida Dreams: 10 -> 23 -> 00 -> 08 (NO 09)
ORDEN_HORAS = list(range(10, 24)) + list(range(0, 9))
HORAS_VALIDAS = set(ORDEN_HORAS)

# ===============================
#   CARGA
# ===============================
raw = pd.read_excel(ARCHIVO, sheet_name=HOJA_DATOS, engine="openpyxl")
raw.columns = raw.iloc[0]          # primera fila = encabezados reales
df = raw.iloc[1:].copy()

df = df.rename(columns={
    "Id Cliente": "IdCliente",
    "Slot Attendant": "Attendant",
    "Forma Pago": "FormaPago",
    "Ingreso CAWA": "Ingreso",
})

# ===============================
#   LIMPIEZA BASE
# ===============================
df["Fecha"] = pd.to_datetime(df.get("Fecha"), errors="coerce", dayfirst=True)
df["Jornada"] = pd.to_datetime(df.get("Jornada"), errors="coerce", dayfirst=True)
df["Monto"] = pd.to_numeric(df.get("Monto"), errors="coerce").fillna(0)

df = df.dropna(subset=["Fecha", "Jornada", "Attendant"]).copy()

# Jornada manda (solo día)
df["JornadaDia"] = df["Jornada"].dt.normalize()

# Hora desde Fecha
df["Hora"] = df["Fecha"].dt.hour

# Excluir horas fuera de jornada (incluye 09)
df = df[df["Hora"].isin(HORAS_VALIDAS)].copy()

# Mes por Jornada
df["Mes"] = df["JornadaDia"].dt.to_period("M").astype(str)

# ===============================
#   TABLAS
# ===============================
# 1) Mes / Operaciones / Monto
tabla_mes = (
    df.groupby("Mes")
      .agg(Operaciones=("Monto", "count"),
           Monto=("Monto", "sum"))
      .reset_index()
      .sort_values("Mes")
)

# 2) Hora / Operaciones / Monto
tabla_hora = (
    df.groupby("Hora")
      .agg(Operaciones=("Monto", "count"),
           Monto=("Monto", "sum"))
      .reindex(ORDEN_HORAS, fill_value=0)
      .reset_index()
)

# 3) Récord por Asistente (máximo en una jornada)
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

# 4) Asistente / Mes / Operaciones / Monto
tabla_asistente_mes = (
    df.groupby(["Attendant", "Mes"])
      .agg(Operaciones=("Monto", "count"),
           Monto=("Monto", "sum"))
      .reset_index()
      .sort_values(["Mes", "Operaciones"], ascending=[True, False])
)

# ===============================
#   QA MINIMO
# ===============================
qa_df = pd.DataFrame([
    ["filas_usadas", len(df)],
    ["min_fecha", str(df["Fecha"].min())],
    ["max_fecha", str(df["Fecha"].max())],
    ["horas_presentes", ", ".join(map(str, sorted(df["Hora"].unique())))],
], columns=["Metrica", "Valor"])

# ===============================
#   EXPORTAR + AUTOSIZE
# ===============================
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    tablas = {
        "Resumen Mensual": tabla_mes,
        "Operaciones por Hora": tabla_hora,
        "Record Asistentes": tabla_record,
        "Asistente por Mes": tabla_asistente_mes,
        "QA": qa_df
    }

    for nombre_hoja, tabla in tablas.items():
        tabla.to_excel(writer, sheet_name=nombre_hoja, index=False)
        ws = writer.book[nombre_hoja]

        # Autoajuste de ancho de columnas
        for col_idx, col_name in enumerate(tabla.columns, start=1):
            max_len = max(
                len(str(col_name)),
                *(len(str(val)) for val in tabla[col_name].astype(str))
            )
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

print(f"✅ Listo: {OUTPUT_FILE}")
