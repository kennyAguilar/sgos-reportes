# SGOS Reportes — Migración a C# / .NET 8

> **Propósito**: Blueprint para recrear la app Flask actual (`sgos_web/`) en
> **C# con ASP.NET Core 8** como proyecto independiente en una carpeta nueva.
> Contiene arquitectura, skills, paquetes NuGet, estructura y plan por fases.
>
> **Repo fuente**: `kennyAguilar/sgos-reportes` (Python/Flask)
> **Repo destino**: nuevo, p. ej. `sgos-reportes-net`

---

## 1. Resumen ejecutivo

App interna de auditoría de operaciones de casino (Getnet, Premios, COMPS).

**Stack actual**: Flask + SQLAlchemy + Flask-Login + Flask-WTF + Flask-Limiter,
PostgreSQL (Neon), Jinja2 + Bootstrap 5 + Chart.js, pandas + openpyxl,
PyWebView + PyInstaller (desktop).

**Stack destino**: .NET 8 + ASP.NET Core Razor Pages, manteniendo **misma UI,
misma DB y misma funcionalidad**, ganando ejecutable nativo, mejor performance
en SQL grande (COMPS) y distribución sin Python.

---

## 2. Stack tecnológico destino

| Capa | Tecnología | Por qué |
|---|---|---|
| Runtime | **.NET 8 LTS** | Soporte largo plazo |
| Web | **ASP.NET Core 8 + Razor Pages** | Mapeo 1:1 con Flask + Jinja2 |
| ORM ligero | **Dapper** | Casi idéntico a `db.session.execute(text(...))` |
| ORM completo | **EF Core 8** | Solo modelos simples (User, Jefatura, etc.) |
| Driver DB | **Npgsql 8** | PostgreSQL nativo |
| Auth | **ASP.NET Core Identity** (cookies) | Reemplazo de Flask-Login |
| CSRF | **Antiforgery** built-in | Reemplazo de Flask-WTF |
| Rate limit | **Microsoft.AspNetCore.RateLimiting** | Reemplazo de Flask-Limiter |
| Excel | **ClosedXML** (preferido) o **EPPlus 7** | Reemplazo de openpyxl/pandas |
| LINQ analítico | `System.Linq` + GroupBy | Reemplazo de pandas groupby/pivot |
| Logging | **Serilog** | Estructurado |
| Frontend | **Bootstrap 5 + Bootstrap Icons + Chart.js** | Sin cambios |
| PWA | Manifest + Service Worker | Reutilizar `static/sw.js` |
| Desktop | **WebView2** + WPF host (o **MAUI Blazor Hybrid**) | Reemplazo de PyWebView |
| Empaquetado | `dotnet publish -r win-x64 --self-contained -p:PublishSingleFile=true` | Reemplazo de PyInstaller |

---

## 3. Skills necesarios (skills.sh)

### 3.1 Skills oficiales recomendados

| Skill | Cuándo usarlo |
|---|---|
| `azure-prepare` | Si despliegas a Azure App Service / Container Apps |
| `azure-deploy` | Ejecutar el despliegue una vez preparado |
| `azure-validate` | Validación pre-deploy (Bicep, RBAC, configuración) |
| `azure-rbac` | Si usas Managed Identity para Key Vault / Storage |
| `azure-diagnostics` | Debugging en producción (App Insights, KQL) |
| `appinsights-instrumentation` | Telemetría con OpenTelemetry .NET SDK |

### 3.2 Skills personalizados a crear

Crea estos `SKILL.md` propios bajo `~/.agents/skills/`:

#### `sgos-net-architecture/SKILL.md`
- **WHEN**: cualquier cambio estructural, nuevo módulo, refactor de capas
- **Contenido**: convenciones de carpetas (Pages/Services/Repositories/Models),
  patrón Repository + Dapper, DI, transacciones, DTOs vs entidades EF.

#### `sgos-net-dapper-patterns/SKILL.md`
- **WHEN**: convertir SQL crudo de `comps_routes.py` a Dapper
- **Contenido**: `text(":param")` → `@param`, `DynamicParameters`, helpers
  `_exec/_exec_one`, `decimal` vs `float`, paginación.

#### `sgos-net-excel-pipeline/SKILL.md`
- **WHEN**: portar `engine.py` / `comps_engine.py` o exportar Excel
- **Contenido**: lectura ClosedXML, normalización de columnas (alias,
  case-insensitive), conversión a `List<T>`, escritura multi-hoja, formatos.

#### `sgos-net-linq-aggregations/SKILL.md`
- **WHEN**: traducir agrupaciones pandas (groupby, pivot, resample)
- **Contenido**: equivalencias pandas → LINQ:
  - `df.groupby('mes').agg({'monto':'sum'})` → `GroupBy(x => x.Mes).Select(...)`
  - `pd.pivot_table(...)` → `GroupBy().Select(g => new {...})`
  - resample por hora/día/mes con `DateTime.Hour/Date`

#### `sgos-net-razor-migration/SKILL.md`
- **WHEN**: portar templates Jinja2 → Razor (.cshtml)
- **Contenido**: `{% extends %}` → `_Layout.cshtml`, `{{ var }}` → `@Model.Var`,
  `{% for %}` → `@foreach`, `{{ csrf_token() }}` → `@Html.AntiForgeryToken()`,
  flash → TempData, `url_for()` → `Url.Page()` / `asp-page` tag helper.

#### `sgos-net-webview2-desktop/SKILL.md`
- **WHEN**: empaquetar como app desktop Windows
- **Contenido**: bootstrap WPF + WebView2, Kestrel en background, puerto
  efímero, guardado en `~/Downloads`, splash, instalador Inno Setup / MSIX.

### 3.3 Plantilla mínima de skill

```markdown
---
name: sgos-net-dapper-patterns
description: Patrones para convertir SQL crudo de Flask/SQLAlchemy a Dapper
  en .NET 8. WHEN: portar queries de comps_routes.py, traducir text(:param)
  a @param, manejar DynamicParameters, transacciones Npgsql.
---

# Contenido del skill...
```

---

## 4. Estructura de carpetas del nuevo proyecto

```
sgos-reportes-net/
├── SgosReportes.sln
├── README.md
├── .editorconfig
├── .gitignore
├── global.json                       ← fija .NET 8 SDK
│
├── src/
│   ├── SgosReportes.Web/             ← ASP.NET Core principal
│   │   ├── Program.cs
│   │   ├── appsettings.json
│   │   ├── Pages/
│   │   │   ├── Shared/_Layout.cshtml         ← layout.html
│   │   │   ├── Shared/_LayoutComps.cshtml    ← comps/layout_comps.html
│   │   │   ├── Index.cshtml                  ← home.html
│   │   │   ├── Login.cshtml
│   │   │   ├── Sgos/{Index,Dashboard,Graphs}.cshtml
│   │   │   ├── Premios/{Dashboard,Graphs}.cshtml
│   │   │   ├── Usuarios/Index.cshtml
│   │   │   └── Comps/
│   │   │       ├── Index.cshtml
│   │   │       ├── AnalisisCortesias.cshtml
│   │   │       ├── AnalisisPremios.cshtml
│   │   │       ├── AnalisisResumen.cshtml
│   │   │       ├── ControlInvitaciones{,Mda,Mdj}.cshtml
│   │   │       ├── AuditoriaCoinInCero.cshtml
│   │   │       ├── Exportar.cshtml
│   │   │       └── Configuracion.cshtml      ← CRUD jefaturas/categorías
│   │   └── wwwroot/
│   │       ├── css/style.css
│   │       ├── js/{sortable.js, sw.js}
│   │       ├── icons/
│   │       └── manifest.webmanifest
│   │
│   ├── SgosReportes.Core/            ← lógica de dominio
│   │   ├── Models/                   ← Operacion, Premio, Cortesia, ...
│   │   ├── Services/
│   │   │   ├── ExcelImportService.cs        ← engine.py
│   │   │   ├── ReportService.cs             ← generar_reportes()
│   │   │   ├── CompsAnalyticsService.cs     ← consultas COMPS
│   │   │   └── ExcelExportService.cs        ← exportar_excel_bytes
│   │   ├── Repositories/
│   │   │   ├── OperacionRepository.cs (Dapper)
│   │   │   └── CompsRepository.cs (Dapper)
│   │   └── Helpers/
│   │       ├── DateFilterBuilder.cs         ← build_date_filter()
│   │       └── MoneyFormatter.cs
│   │
│   ├── SgosReportes.Data/            ← EF Core context + migrations
│   │   ├── AppDbContext.cs
│   │   └── Migrations/
│   │
│   └── SgosReportes.Desktop/         ← WebView2 (opcional)
│       ├── App.xaml(.cs)
│       ├── MainWindow.xaml(.cs)
│       └── KestrelHost.cs
│
├── tests/SgosReportes.Tests/
└── infra/main.bicep                  ← opcional, Azure
```

---

## 5. Paquetes NuGet por proyecto

### `SgosReportes.Web.csproj`
```xml
<PackageReference Include="Microsoft.AspNetCore.Identity.EntityFrameworkCore" Version="8.0.*" />
<PackageReference Include="Microsoft.EntityFrameworkCore.Design" Version="8.0.*" />
<PackageReference Include="Npgsql.EntityFrameworkCore.PostgreSQL" Version="8.0.*" />
<PackageReference Include="Serilog.AspNetCore" Version="8.0.*" />
<PackageReference Include="Serilog.Sinks.File" Version="6.0.*" />
```

### `SgosReportes.Core.csproj`
```xml
<PackageReference Include="Dapper" Version="2.1.*" />
<PackageReference Include="Npgsql" Version="8.0.*" />
<PackageReference Include="ClosedXML" Version="0.102.*" />
<PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="8.0.*" />
```

### `SgosReportes.Desktop.csproj` (opcional)
```xml
<PropertyGroup>
  <OutputType>WinExe</OutputType>
  <TargetFramework>net8.0-windows</TargetFramework>
  <UseWPF>true</UseWPF>
</PropertyGroup>
<ItemGroup>
  <PackageReference Include="Microsoft.Web.WebView2" Version="1.0.*" />
</ItemGroup>
```

---

## 6. Mapeo de archivos: Python → C#

| Origen (Flask) | Destino (.NET) |
|---|---|
| `sgos_web/app.py` | `Program.cs` + Pages individuales |
| `sgos_web/extensions.py` | DI container en `Program.cs` |
| `sgos_web/engine.py` | `Core/Services/ExcelImportService.cs` + `ReportService.cs` |
| `sgos_web/comps_engine.py` | `Core/Services/CompsImportService.cs` |
| `sgos_web/comps_routes.py` | `Web/Pages/Comps/*.cshtml.cs` (PageModels) |
| `templates/layout.html` | `Pages/Shared/_Layout.cshtml` |
| `templates/comps/layout_comps.html` | `Pages/Shared/_LayoutComps.cshtml` |
| `templates/*.html` | `Pages/**/*.cshtml` |
| `static/` | `wwwroot/` |
| `requirements.txt` | `*.csproj` (PackageReference) |
| `Procfile` | `Dockerfile` o `dotnet publish` |
| `desktop.py` | `Desktop/MainWindow.xaml.cs` |
| `SGOS.spec` | `dotnet publish -p:PublishSingleFile=true` |

---

## 7. Mapeo de funciones críticas

### 7.1 Helpers SQL de COMPS

```python
# Python
def _exec(sql, params): ...
def build_date_filter(col, anio, mes): ...
```

```csharp
// C#
public static class DapperExtensions {
    public static IEnumerable<T> Exec<T>(this IDbConnection c, string sql, object? p = null)
        => c.Query<T>(sql, p);
    public static T? ExecOne<T>(this IDbConnection c, string sql, object? p = null)
        => c.QueryFirstOrDefault<T>(sql, p);
}

public static class DateFilterBuilder {
    public static (string Where, DynamicParameters Params)
        Build(string column, int? anio, int? mes) {
        var p = new DynamicParameters();
        var conds = new List<string>();
        if (anio.HasValue) {
            conds.Add($"EXTRACT(YEAR FROM {column}::date) = @anio");
            p.Add("anio", anio.Value);
        }
        if (mes.HasValue) {
            conds.Add($"EXTRACT(MONTH FROM {column}::date) = @mes");
            p.Add("mes", mes.Value);
        }
        return (conds.Count > 0 ? "WHERE " + string.Join(" AND ", conds) : "", p);
    }
}
```

### 7.2 Reportes (pandas → LINQ)

```python
df.groupby('Mes').agg(Operaciones=('Monto','count'), Monto=('Monto','sum'))
```

```csharp
var resumen = operaciones
    .GroupBy(o => o.Mes)
    .Select(g => new ResumenMensual {
        Mes = g.Key,
        Operaciones = g.Count(),
        Monto = g.Sum(o => o.Monto)
    })
    .OrderBy(x => x.Mes)
    .ToList();
```

### 7.3 CSRF

```html
<!-- Jinja2 -->          <input type="hidden" name="csrf_token" value="{{ csrf_token() }}">
<!-- Razor -->           @Html.AntiForgeryToken()
```

### 7.4 Flash messages

```python
flash("Mensaje", "success")
```
```csharp
TempData["Success"] = "Mensaje";
// Leer en _Layout.cshtml: TempData["Success"], TempData["Warning"], etc.
```

---

## 8. Configuración (`appsettings.json`)

```json
{
  "ConnectionStrings": {
    "Postgres": "Host=ep-xxx.aws.neon.tech;Database=neondb;Username=...;Password=...;SslMode=Require"
  },
  "Auth": {
    "AdminDefaultPassword": "admin123",
    "MinPasswordLength": 9
  },
  "App": {
    "UploadFolder": "uploads",
    "MaxUploadMb": 20,
    "DesktopMode": false
  },
  "Serilog": {
    "MinimumLevel": "Information",
    "WriteTo": [
      { "Name": "Console" },
      { "Name": "File", "Args": { "path": "logs/app-.log", "rollingInterval": "Day" } }
    ]
  }
}
```

> Nunca commitear secretos. Usar `dotnet user-secrets` en dev y variables
> de entorno (o Azure Key Vault) en producción.

---

## 9. Plan de migración por fases

### Fase 0 — Setup (½ día)
1. Crear solución y proyectos vacíos
2. `global.json`, `.editorconfig`, `.gitignore` (.NET)
3. Conexión a Neon (misma DB)
4. Health check `/health`

### Fase 1 — Auth y layout base (1 día)
1. Migrar `users` con ASP.NET Core Identity (`IPasswordHasher` custom para hashes Werkzeug)
2. Páginas Login / Logout
3. `_Layout.cshtml` con navbar, flash messages, CSRF
4. Página Home con cards
5. CRUD de usuarios

### Fase 2 — Módulo Getnet/Premios (2-3 días)
1. Modelos `Operacion` y `Premio` con EF Core
2. `ExcelImportService` con ClosedXML
3. `ReportService.GenerarReportes()` con LINQ
4. Página de carga `/sgos`
5. `Dashboard` con tablas HTML
6. `Graphs` con Chart.js (datos vía `Json.Serialize`)
7. `/download/{id}` con ClosedXML

### Fase 3 — Módulo COMPS (3-5 días)
1. Modelos: `Cortesia`, `PremioComps`, `MesaPuntos`, `SrwJugador`, `Jefatura`, `CategoriaNivel`
2. `CompsImportService` (4-5 Excel distintos)
3. `CompsRepository` con Dapper (porting 1:1 del SQL)
4. Páginas: AnálisisCortesías, AnálisisPremios, Resumen, Control Invitaciones
   (General/MDA/MDJ), Coin-In Cero, Exportar
5. CRUD `Configuracion` (jefaturas + categorías)
6. Exportar Excel multi-hoja con ClosedXML

### Fase 4 — PWA + Service Worker (½ día)
1. Copiar manifest, iconos y `sw.js` a `wwwroot/`
2. Endpoint `/sw.js` con header `Service-Worker-Allowed: /`

### Fase 5 — Desktop WebView2 (1 día, opcional)
1. Proyecto WPF `SgosReportes.Desktop`
2. `KestrelHost.cs`: arrancar `IHost` en background, puerto efímero
3. `MainWindow` con `WebView2` apuntando a `http://127.0.0.1:{puerto}`
4. `DesktopMode = true` → guarda Excel directo en `~/Downloads`
5. Instalador con Inno Setup

### Fase 6 — Tests + CI (1 día)
1. xUnit para `Core/Services` con datasets fijos
2. GitHub Actions: build, test, publish artifact

### Fase 7 — Deploy (½-1 día)
- **Web**: Azure App Service / Container Apps / Render (Docker)
- **Desktop**: GitHub Releases con instalador `.exe`

**Total estimado**: 9-13 días de trabajo enfocado.

---

## 10. Comandos útiles

```bash
# Crear solución
dotnet new sln -n SgosReportes
dotnet new webapp -n SgosReportes.Web -o src/SgosReportes.Web --auth Individual
dotnet new classlib -n SgosReportes.Core -o src/SgosReportes.Core
dotnet new classlib -n SgosReportes.Data -o src/SgosReportes.Data
dotnet sln add (Get-ChildItem -Recurse -Filter *.csproj)

# Referencias entre proyectos
dotnet add src/SgosReportes.Web reference src/SgosReportes.Core src/SgosReportes.Data
dotnet add src/SgosReportes.Core reference src/SgosReportes.Data

# Paquetes
dotnet add src/SgosReportes.Core package Dapper
dotnet add src/SgosReportes.Core package Npgsql
dotnet add src/SgosReportes.Core package ClosedXML
dotnet add src/SgosReportes.Web package Npgsql.EntityFrameworkCore.PostgreSQL
dotnet add src/SgosReportes.Web package Serilog.AspNetCore

# Migraciones EF
dotnet ef migrations add Init -p src/SgosReportes.Data -s src/SgosReportes.Web
dotnet ef database update -p src/SgosReportes.Data -s src/SgosReportes.Web

# Ejecutar
dotnet run --project src/SgosReportes.Web

# Publicar single-file Windows
dotnet publish src/SgosReportes.Web -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

---

## 11. Compatibilidad de DB

La DB PostgreSQL **se mantiene exactamente igual**. La nueva app .NET conecta
al mismo Neon. Permite:

- Ejecutar ambas versiones en paralelo durante la migración
- Validar resultados Python vs .NET con los mismos datos
- Rollback inmediato si surge algún problema

**Tablas a respetar tal cual**: `users`, `operaciones`, `premios`,
`srw_jugadores`, `cortesias`, `premios_comps`, `mesas_puntos`, `jefaturas`,
`categorias_nivel`, `carga_log`.

> ⚠️ **Hashes de contraseña**: Werkzeug usa `pbkdf2:sha256`. .NET Identity
> usa otro formato. Opciones:
> 1. Implementar `IPasswordHasher<User>` custom que verifique formato
>    Werkzeug y migre transparente al primer login.
> 2. Forzar reset de contraseñas en la primera ejecución.

---

## 12. Riesgos y mitigaciones

| Riesgo | Mitigación |
|---|---|
| Diferencia parsing fechas | Tests con dataset real, normalizar a `DateOnly` |
| Floats vs decimals (pandas usa float64) | Usar `decimal` en C# para montos |
| Hashes Werkzeug incompatibles | `IPasswordHasher` custom (§11) |
| ClosedXML lento en Excel >100k filas | Streaming o EPPlus con `LoadFromCollection` |
| WebView2 requiere runtime instalado | Bootstrapper Microsoft "Evergreen" |
| SQL Dapper vs `text()` | Portar 1:1, validar contra Python |

---

## 13. Checklist antes de iniciar

- [ ] Crear repositorio nuevo `sgos-reportes-net` en GitHub
- [ ] Copiar este `README.md` en la raíz del nuevo repo
- [ ] Instalar .NET 8 SDK
- [ ] Instalar VS Code o Visual Studio 2022 con workload "ASP.NET and web development"
- [ ] Connection string de Neon en `dotnet user-secrets`
- [ ] Crear los `SKILL.md` personalizados de §3.2
- [ ] Backup de la DB antes de la primera ejecución
- [ ] Mantener la versión Flask corriendo en paralelo durante migración

---

## 14. Referencias rápidas al código actual

Orden de estudio recomendado para el agente que implemente la migración:

1. `sgos_web/app.py` — modelos, rutas Getnet/Premios, auth
2. `sgos_web/engine.py` — pipeline pandas + reportes
3. `sgos_web/comps_routes.py` — rutas COMPS y SQL grandes
4. `sgos_web/comps_engine.py` — carga de Excel COMPS
5. `sgos_web/templates/layout.html` — layout base
6. `sgos_web/templates/comps/layout_comps.html` — layout COMPS
7. `sgos_web/templates/graphs.html` — Chart.js (se reutiliza)
8. `sgos_web/static/css/style.css` — estilos (se reutiliza)
9. `desktop.py` — bootstrap PyWebView (referencia para WebView2)

---

**Fin del documento.** Listo para empezar el nuevo proyecto.
