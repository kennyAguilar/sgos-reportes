# SGOS - Reportes de Operaciones

Aplicación web para procesar y visualizar reportes de operaciones SGOS desde archivos Excel.

## Características

✨ Selección de tablas a visualizar
- Resumen Mensual
- Operaciones por Hora
- Récord Asistentes
- Asistente por Mes
- QA

🎯 Filtro avanzado por asistentes
- Selecciona qué asistentes visualizar
- Los datos se filtran en tiempo real
- Exporta solo los datos seleccionados

💾 Descarga de reportes en Excel
- Descarga solo las tablas y asistentes seleccionados

## Requisitos

- Python 3.8+
- Flask
- Pandas
- OpenPyXL

## Instalación

1. Clona el repositorio:
```bash
git clone <tu-repo-url>
cd Registro\ de\ Getnet\ y\ Premios
```

2. Crea un entorno virtual:
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # Linux/Mac
```

3. Instala las dependencias:
```bash
pip install -r sgos_web/requirements.txt
```

4. Ejecuta la aplicación:
```bash
python sgos_web/app.py
```

5. Abre tu navegador en `http://localhost:5000`

## Uso

1. Sube un archivo Excel (.xlsx o .xls)
2. Selecciona qué tablas deseas ver
3. En el dashboard, filtra por asistentes (opcional)
4. Visualiza los reportes o descárgalos en Excel

## Estructura del Proyecto

```
├── sgos_web/
│   ├── app.py           # Aplicación Flask principal
│   ├── motor.py         # Lógica de procesamiento de datos
│   ├── requirements.txt  # Dependencias
│   ├── templates/       # Plantillas HTML
│   │   ├── index.html       # Página de carga
│   │   └── dashboard.html   # Dashboard de reportes
│   ├── uploads/         # Archivos subidos
│   └── __pycache__/     # Cache
└── README.md            # Este archivo
```

## Licencia

MIT
