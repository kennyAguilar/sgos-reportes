# Changelog

## [1.1.0] - 2026-02-16

### Agregado
- Implementación de cálculo de **Promedios por Hora** para Operaciones y Montos.

### Detalles de la Implementación (Reporte de Cambios)

Se han añadido dos nuevas columnas: "Operaciones Promedio" y "Monto Promedio".

#### 1. Cálculo en Backend (`engine.py`)
El cálculo se realiza de la siguiente manera:
- Se cuenta cuántas **"Jornadas"** (días operativos distintos) existen en el periodo seleccionado (histórico completo, año o mes).
  ```python
  total_dias = df_calc["JornadaDia"].nunique()
  ```
- **Operaciones Promedio:** Se calcula dividiendo el total de operaciones en esa hora entre el número de días operativos.
  - Fórmula: `Operaciones / total_dias`
- **Monto Promedio:** Se calcula dividiendo el monto total acumulado en esa hora entre el número de días operativos.
  - Fórmula: `Monto / total_dias`

#### 2. Reportes Excel
- Al descargar el reporte (tanto para Operaciones como para Premios), la hoja **"Operaciones por Hora"** incluye ahora las dos nuevas columnas de promedio al final.

#### 3. Gráficos Web (`graphs.html`)
- Se actualizaron los gráficos inferiores ("Operaciones por Hora" y "Montos por Hora").
- Ahora visualizan el **promedio calculado** en lugar de los totales acumulados.
- Se actualizaron títulos y etiquetas para reflejar claramente que los datos mostrados son promedios.

> **Nota:** Esta visualización permite entender el comportamiento "típico" de cada hora, sin importar si el filtro aplicado corresponde a un mes de 30 días o un año de 365 días.
