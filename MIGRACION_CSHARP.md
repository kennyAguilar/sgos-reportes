# SGOS Reportes — Migración a C# / .NET 8 (Desktop Windows)

> **Propósito**: Blueprint para recrear la app Flask actual (`sgos_web/`) como
> **aplicación de escritorio nativa Windows en C# / .NET 8 con WPF**,
> manteniendo la misma base de datos PostgreSQL (Neon) y la misma lógica de
> negocio. **NO es una web app**: es un `.exe` nativo con ventanas XAML.
>
> **Repo fuente**: `kennyAguilar/sgos-reportes` (Python/Flask web)
> **Repo destino**: nuevo, p. ej. `sgos-reportes-desktop`

---

## 1. Resumen ejecutivo

App interna de auditoría de operaciones de casino (Getnet, Premios, COMPS).

**Stack actual (web)**: Flask + SQLAlchemy + Jinja2 + Bootstrap + Chart.js,
PostgreSQL (Neon), pandas + openpyxl, PyWebView + PyInstaller.

**Stack destino (desktop puro)**: **.NET 8 + WPF** con MVVM, conexión directa
a PostgreSQL vía Npgsql/Dapper, UI nativa XAML, gráficos con LiveCharts2,
Excel con ClosedXML, empaquetado como **ejecutable único Windows**.

**Sin servidor web, sin HTML/CSS/JS, sin Kestrel, sin navegador**. La app es
100% cliente grueso que habla directo con Neon por SSL.

---

## 2. Stack tecnológico destino

| Capa | Tecnología | Por qué |
|---|---|---|
| Runtime | **.NET 8 LTS** | Soporte largo plazo, self-contained |
| UI framework | **WPF** (recomendado) | Maduro, estable, gran ecosistema |
| Patrón UI | **MVVM** con `CommunityToolkit.Mvvm` | `[ObservableProperty]`, `[RelayCommand]` |
| DI / Host | **Microsoft.Extensions.Hosting** | DI estándar |
| ORM ligero | **Dapper** | SQL crudo (equivalente a `text(...)`) |
| ORM completo | **EF Core 8** (opcional) | Modelos simples |
| Driver DB | **Npgsql 8** | PostgreSQL nativo |
| Gráficos | **LiveChartsCore.SkiaSharpView.WPF** | Reemplazo de Chart.js nativo |
| Tablas | `DataGrid` nativo WPF | Sort, filtros, paginación |
| Excel | **ClosedXML** | Reemplazo de openpyxl/pandas |
| LINQ analítico | `System.Linq` + `GroupBy` | Reemplazo de pandas groupby |
| Hash contraseñas | **PBKDF2** y **Scrypt** compatibles Werkzeug | Leer hashes actuales sin reset |
| Logging | **Serilog** | Logs estructurados en `%LOCALAPPDATA%` |
| Tests | **xUnit** + **FluentAssertions** | Tests de servicios |
| Empaquetado | `dotnet publish -r win-x64 --self-contained -p:PublishSingleFile=true` | Reemplazo de PyInstaller |
| Instalador | **Inno Setup** | Distribución Windows |

---

## 3. Arquitectura MVVM

```
┌──────────────────────────────────────────────────────────┐
│  Views (XAML)                                            │
│  LoginWindow, MainWindow, DashboardView, GraphsView, ... │
└──────────────────────┬───────────────────────────────────┘
                       │ DataContext binding
┌──────────────────────▼───────────────────────────────────┐
│  ViewModels (MVVM)                                       │
│  [ObservableProperty] [RelayCommand]                     │
└──────────────────────┬───────────────────────────────────┘
                       │ DI
┌──────────────────────▼───────────────────────────────────┐
│  Services                                                │
│  AuthService, ExcelImportService, ReportService,         │
│  CompsAnalyticsService, ExcelExportService               │
└──────────────────────┬───────────────────────────────────┘
                       │
┌──────────────────────▼───────────────────────────────────┐
│  Repositories (Dapper + Npgsql)                          │
└──────────────────────┬───────────────────────────────────┘
                       │ SSL
┌──────────────────────▼───────────────────────────────────┐
│  PostgreSQL (Neon) — misma DB que la web Flask           │
└──────────────────────────────────────────────────────────┘
```

---

## 4. Estructura de carpetas del nuevo proyecto

```
sgos-reportes-desktop/
├── SgosReportes.sln
├── README.md
├── .editorconfig
├── .gitignore
├── global.json
│
├── src/
│   ├── SgosReportes.App/                 ← Proyecto WPF (.exe)
│   │   ├── App.xaml(.cs)                 ← bootstrap DI + Serilog
│   │   ├── appsettings.json
│   │   ├── Views/
│   │   │   ├── LoginWindow.xaml
│   │   │   ├── MainWindow.xaml
│   │   │   ├── HomeView.xaml
│   │   │   ├── CargarArchivoView.xaml
│   │   │   ├── DashboardGetnetView.xaml
│   │   │   ├── GraphsGetnetView.xaml
│   │   │   ├── DashboardPremiosView.xaml
│   │   │   ├── GraphsPremiosView.xaml
│   │   │   ├── DashboardCombinadoView.xaml
│   │   │   ├── HistorialCargasView.xaml
│   │   │   ├── UsuariosView.xaml
│   │   │   └── Comps/
│   │   │       ├── CompsIndexView.xaml
│   │   │       ├── AnalisisCortesiasView.xaml
│   │   │       ├── AnalisisPremiosView.xaml
│   │   │       ├── AnalisisResumenView.xaml
│   │   │       ├── AuditoriaCoinInCeroView.xaml
│   │   │       ├── ControlInvitacionesView.xaml
│   │   │       ├── ControlInvitacionesMdaView.xaml
│   │   │       ├── ControlInvitacionesMdjView.xaml
│   │   │       ├── ControlInvitacionesMktView.xaml
│   │   │       ├── ConfiguracionView.xaml
│   │   │       └── ExportarView.xaml
│   │   ├── ViewModels/
│   │   │   ├── LoginViewModel.cs
│   │   │   ├── MainViewModel.cs
│   │   │   ├── DashboardGetnetViewModel.cs
│   │   │   ├── GraphsGetnetViewModel.cs
│   │   │   ├── UsuariosViewModel.cs
│   │   │   └── Comps/*.cs
│   │   ├── Controls/
│   │   │   ├── LoadingOverlay.xaml
│   │   │   ├── DataGridPaged.xaml
│   │   │   └── KpiCard.xaml
│   │   ├── Styles/
│   │   │   ├── Colors.xaml                ← paleta Casino Royale
│   │   │   ├── Buttons.xaml
│   │   │   ├── DataGrid.xaml
│   │   │   └── Typography.xaml
│   │   ├── Converters/
│   │   │   ├── MoneyConverter.cs
│   │   │   └── DateFormatConverter.cs
│   │   └── Assets/
│   │       ├── app.ico
│   │       └── images/
│   │
│   └── SgosReportes.Core/                 ← Lógica de dominio (net8.0)
│       ├── Models/
│       │   ├── User.cs
│       │   ├── Operacion.cs
│       │   ├── Premio.cs
│       │   ├── Cortesia.cs
│       │   ├── PremioComps.cs
│       │   ├── MesaPuntos.cs
│       │   ├── SrwJugador.cs
│       │   ├── Jefatura.cs
│       │   └── CategoriaNivel.cs
│       ├── Services/
│       │   ├── IAuthService.cs / AuthService.cs
│       │   ├── IExcelImportService.cs / ExcelImportService.cs
│       │   ├── IReportService.cs / ReportService.cs
│       │   ├── ICompsAnalyticsService.cs / CompsAnalyticsService.cs
│       │   └── IExcelExportService.cs / ExcelExportService.cs
│       ├── Repositories/
│       │   ├── IUserRepository.cs / UserRepository.cs
│       │   ├── IOperacionRepository.cs / OperacionRepository.cs
│       │   ├── IPremioRepository.cs / PremioRepository.cs
│       │   └── ICompsRepository.cs / CompsRepository.cs
│       ├── Security/
│       │   └── WerkzeugCompatibleHasher.cs  ← pbkdf2:sha256 + scrypt + sgos custom
│       └── Helpers/
│           ├── DateFilterBuilder.cs
│           ├── MoneyFormatter.cs
│           └── DapperExtensions.cs
│
├── tests/
│   └── SgosReportes.Core.Tests/
│       ├── ReportServiceTests.cs
│       ├── CompsAnalyticsServiceTests.cs
│       └── WerkzeugCompatibleHasherTests.cs
│
└── installer/
    └── sgos-reportes.iss                  ← Inno Setup
```

---

## 5. Paquetes NuGet por proyecto

### `SgosReportes.App.csproj` (WPF)
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net8.0-windows</TargetFramework>
    <UseWPF>true</UseWPF>
    <Nullable>enable</Nullable>
    <ApplicationIcon>Assets\app.ico</ApplicationIcon>
    <AssemblyName>SGOS</AssemblyName>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="CommunityToolkit.Mvvm" Version="8.2.*" />
    <PackageReference Include="Microsoft.Extensions.Hosting" Version="8.0.*" />
    <PackageReference Include="Microsoft.Extensions.Configuration.Json" Version="8.0.*" />
    <PackageReference Include="Microsoft.Extensions.DependencyInjection" Version="8.0.*" />
    <PackageReference Include="LiveChartsCore.SkiaSharpView.WPF" Version="2.0.0-rc5" />
    <PackageReference Include="Serilog.Extensions.Hosting" Version="8.0.*" />
    <PackageReference Include="Serilog.Sinks.File" Version="6.0.*" />
    <PackageReference Include="Serilog.Sinks.Debug" Version="2.0.*" />
  </ItemGroup>
</Project>
```

### `SgosReportes.Core.csproj`
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Dapper" Version="2.1.*" />
    <PackageReference Include="Npgsql" Version="8.0.*" />
    <PackageReference Include="ClosedXML" Version="0.102.*" />
    <PackageReference Include="Konscious.Security.Cryptography.Scrypt" Version="1.3.*" />
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="8.0.*" />
  </ItemGroup>
</Project>
```

---

## 6. Mapeo de archivos: Python Flask → C# WPF

| Origen (Flask web) | Destino (WPF desktop) |
|---|---|
| `sgos_web/app.py` (rutas) | `MainWindow.xaml` + navegación en `MainViewModel` |
| `sgos_web/app.py` (User + auth) | `Core/Models/User.cs` + `AuthService` + `LoginWindow` |
| `sgos_web/extensions.py` | DI container en `App.xaml.cs` |
| `sgos_web/engine.py` | `Core/Services/ExcelImportService.cs` + `ReportService.cs` |
| `sgos_web/comps_engine.py` | `Core/Services/CompsImportService.cs` |
| `sgos_web/comps_routes.py` | ViewModels bajo `ViewModels/Comps/` |
| `templates/layout.html` | `MainWindow.xaml` (shell con navegación lateral) |
| `templates/login.html` | `LoginWindow.xaml` (modal) |
| `templates/home.html` | `HomeView.xaml` (UserControl con cards) |
| `templates/dashboard.html` | `DashboardGetnetView.xaml` con `DataGrid` |
| `templates/graphs.html` (Chart.js) | `GraphsGetnetView.xaml` con `CartesianChart` |
| `templates/comps/*.html` | `Views/Comps/*.xaml` |
| `static/css/style.css` | `Styles/Colors.xaml` + `Styles/*.xaml` |
| `static/js/sortable.js` | `DataGrid.CanUserSortColumns="True"` |
| `static/js/sgos-charts.js` | Código en cada ViewModel de gráficos |
| `static/manifest.webmanifest` + `sw.js` | (no aplica — app nativa) |
| `desktop.py` (PyWebView) | (no aplica — WPF es nativo) |
| `SGOS.spec` (PyInstaller) | `dotnet publish -p:PublishSingleFile=true` |
| `requirements.txt` | `*.csproj` |

---

## 7. Patrones clave de conversión

### 7.1 Login (Flask-Login → ventana WPF modal)

```python
# Flask
@app.route('/login', methods=['GET','POST'])
def login():
    user = User.query.filter_by(username=form.username.data).first()
    if user and user.check_password(form.password.data):
        login_user(user)
        return redirect(url_for('home'))
```

```csharp
// LoginViewModel.cs
public partial class LoginViewModel : ObservableObject {
    [ObservableProperty] private string _username = "";
    [ObservableProperty] private string _password = "";
    [ObservableProperty] private string _error = "";

    private readonly IAuthService _auth;
    public User? AuthenticatedUser { get; private set; }

    public LoginViewModel(IAuthService auth) => _auth = auth;

    [RelayCommand]
    private async Task LoginAsync(Window window) {
        var user = await _auth.AuthenticateAsync(Username, Password);
        if (user is null) { Error = "Usuario o contraseña inválidos"; return; }
        AuthenticatedUser = user;
        await _auth.UpdateLastLoginAsync(user.Id);
        window.DialogResult = true;
        window.Close();
    }
}
```

### 7.2 Dashboard (tabla paginada)

```xml
<!-- DashboardGetnetView.xaml -->
<UserControl>
  <Grid>
    <DataGrid ItemsSource="{Binding Operaciones}"
              AutoGenerateColumns="False"
              CanUserSortColumns="True"
              IsReadOnly="True">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Fecha"   Binding="{Binding Fecha, StringFormat=dd/MM/yyyy}" />
        <DataGridTextColumn Header="Cliente" Binding="{Binding IdCliente}" />
        <DataGridTextColumn Header="Monto"
                            Binding="{Binding Monto, Converter={StaticResource MoneyConverter}}" />
      </DataGrid.Columns>
    </DataGrid>
  </Grid>
</UserControl>
```

```csharp
public partial class DashboardGetnetViewModel : ObservableObject {
    [ObservableProperty] private ObservableCollection<Operacion> _operaciones = new();

    public async Task LoadAsync() {
        var data = await _repo.GetPageAsync(page: 1, size: 100);
        Operaciones = new ObservableCollection<Operacion>(data);
    }
}
```

### 7.3 Gráficos (Chart.js → LiveCharts2)

```xml
<lvc:CartesianChart Series="{Binding Series}"
                    XAxes="{Binding XAxes}"
                    YAxes="{Binding YAxes}"
                    xmlns:lvc="clr-namespace:LiveChartsCore.SkiaSharpView.WPF;assembly=LiveChartsCore.SkiaSharpView.WPF" />
```

```csharp
public ISeries[] Series { get; set; } = new ISeries[] {
    new ColumnSeries<decimal> {
        Name = "Monto mensual",
        Values = new[] { 1200m, 1800m, 2100m, 1950m }
    }
};
```

### 7.4 Helper SQL (Dapper)

```python
# Python
def _exec(sql, params): return db.session.execute(text(sql), params).fetchall()
```

```csharp
public static class DapperExtensions {
    public static IEnumerable<T> Exec<T>(this IDbConnection c, string sql, object? p = null)
        => c.Query<T>(sql, p);
    public static T? ExecOne<T>(this IDbConnection c, string sql, object? p = null)
        => c.QueryFirstOrDefault<T>(sql, p);
}
```

### 7.5 `build_date_filter` → `DateFilterBuilder`

```csharp
public static class DateFilterBuilder {
    public static (string Where, DynamicParameters Params)
        Build(string column, int? anio, int? mes) {
        var p = new DynamicParameters();
        var conds = new List<string>();
        if (anio.HasValue) { conds.Add($"EXTRACT(YEAR FROM {column}::date) = @anio"); p.Add("anio", anio.Value); }
        if (mes.HasValue)  { conds.Add($"EXTRACT(MONTH FROM {column}::date) = @mes"); p.Add("mes", mes.Value); }
        return (conds.Count > 0 ? "WHERE " + string.Join(" AND ", conds) : "", p);
    }
}
```

### 7.6 pandas groupby → LINQ

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

### 7.7 Hash compatible con Werkzeug (triple soporte)

La tabla `users` actual tiene tres formatos coexistiendo:
- `scrypt:32768:8:1$salt$hash` (Werkzeug scrypt)
- `pbkdf2:sha256:600000$salt$hash` (Werkzeug PBKDF2)
- `sgos.pbkdf2.sha256.v1$iter$salt$hash` (formato custom del proyecto)

```csharp
public class WerkzeugCompatibleHasher {
    public bool Verify(string password, string stored) {
        if (stored.StartsWith("pbkdf2:sha256:"))         return VerifyPbkdf2Sha256(password, stored);
        if (stored.StartsWith("scrypt:"))                return VerifyScrypt(password, stored);
        if (stored.StartsWith("sgos.pbkdf2.sha256.v1$")) return VerifySgosCustom(password, stored);
        return false;
    }

    private bool VerifyPbkdf2Sha256(string password, string stored) {
        // Formato: pbkdf2:sha256:ITER$SALT$HEXHASH
        var parts = stored.Split('$');
        var header = parts[0].Split(':');
        var iter = int.Parse(header[2]);
        var salt = Encoding.UTF8.GetBytes(parts[1]);
        var expected = parts[2];
        using var pbkdf2 = new Rfc2898DeriveBytes(password, salt, iter, HashAlgorithmName.SHA256);
        var actual = Convert.ToHexString(pbkdf2.GetBytes(32)).ToLowerInvariant();
        return CryptographicOperations.FixedTimeEquals(
            Encoding.ASCII.GetBytes(actual),
            Encoding.ASCII.GetBytes(expected));
    }

    private bool VerifyScrypt(string password, string stored) {
        // Formato: scrypt:N:R:P$SALT$HEXHASH
        // Usar Konscious.Security.Cryptography.Scrypt
        // ...
    }

    // VerifySgosCustom según el formato exacto usado por el proyecto
}
```

> ⚠️ Cuando un usuario logue correctamente con formato legacy, re-hashear con
> el formato canónico (`pbkdf2:sha256` o el propio custom) y guardar, para ir
> unificando la base con el tiempo.

---

## 8. Configuración (`appsettings.json`)

```json
{
  "ConnectionStrings": {
    "Postgres": "Host=ep-xxx.aws.neon.tech;Database=neondb;Username=...;Password=...;SslMode=Require;Trust Server Certificate=true"
  },
  "App": {
    "UploadFolder": "uploads",
    "MaxUploadMb": 20,
    "DefaultDownloadsFolder": "%USERPROFILE%\\Downloads"
  },
  "Serilog": {
    "MinimumLevel": "Information",
    "WriteTo": [
      { "Name": "Debug" },
      { "Name": "File", "Args": { "path": "%LOCALAPPDATA%\\SGOS\\logs\\app-.log", "rollingInterval": "Day" } }
    ]
  }
}
```

**Secretos**: la connection string con password NO va en el repo. Opciones:
- **DPAPI** (`ProtectedData.Protect`) cifrado en `%APPDATA%\SGOS\secrets.dat`
- Pedir al usuario la primera vez y guardar cifrado por usuario Windows
- En dev: `dotnet user-secrets`

---

## 9. Paleta de colores Casino Royale (`Styles/Colors.xaml`)

```xml
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <!-- Superficies -->
  <Color x:Key="BgBaseColor">#FF0A0E1A</Color>
  <Color x:Key="Surface1Color">#FF141B2D</Color>
  <Color x:Key="Surface2Color">#FF1C2540</Color>
  <!-- Acentos -->
  <Color x:Key="GoldColor">#FFD4AF37</Color>
  <Color x:Key="EmeraldColor">#FF10B981</Color>
  <Color x:Key="RubyColor">#FFDC2626</Color>
  <!-- Texto -->
  <Color x:Key="TextPrimaryColor">#FFF5F5F7</Color>
  <Color x:Key="TextSecondaryColor">#FFA8B2CE</Color>

  <SolidColorBrush x:Key="BgBaseBrush"   Color="{StaticResource BgBaseColor}" />
  <SolidColorBrush x:Key="Surface1Brush" Color="{StaticResource Surface1Color}" />
  <SolidColorBrush x:Key="GoldBrush"     Color="{StaticResource GoldColor}" />
  <SolidColorBrush x:Key="EmeraldBrush"  Color="{StaticResource EmeraldColor}" />
</ResourceDictionary>
```

Los tokens de `DESIGN.md` se traducen 1:1 a recursos XAML. Tipografía:
`FontFamily="Plus Jakarta Sans"` si está instalada en el sistema, con fallback
`Segoe UI`.

---

## 10. Plan de migración por fases

### Fase 0 — Setup (½ día)
1. Crear repo `sgos-reportes-desktop` en GitHub
2. `dotnet new sln` + proyectos `SgosReportes.App` (WPF) y `SgosReportes.Core` (classlib)
3. `global.json`, `.editorconfig`, `.gitignore` .NET
4. DI + Serilog en `App.xaml.cs`
5. Health check de conexión a Neon en startup

### Fase 1 — Login + Shell (1 día)
1. `LoginWindow.xaml` + `LoginViewModel` + `AuthService`
2. `WerkzeugCompatibleHasher` con tests (pbkdf2, scrypt, sgos custom)
3. `MainWindow.xaml` con navegación lateral (ListBox + ContentControl)
4. `Styles/Colors.xaml` con paleta Casino Royale
5. `HomeView.xaml` con cards de módulos

### Fase 2 — CRUD Usuarios (½ día)
1. `UserRepository` (Dapper) respetando `is_admin`, `created_at`, `last_login_at`
2. `UsuariosView.xaml` con `DataGrid`
3. Diálogos nuevo/editar/eliminar

### Fase 3 — Módulo Getnet/Premios (2-3 días)
1. Modelos `Operacion`, `Premio` + repos Dapper
2. `ExcelImportService` con ClosedXML
3. `ReportService.GenerarReportes()` en LINQ puro
4. `CargarArchivoView` con `OpenFileDialog` + barra de progreso
5. `DashboardGetnetView`, `DashboardPremiosView` con DataGrid paginado
6. `GraphsGetnetView`, `GraphsPremiosView` con LiveCharts2
7. `DashboardCombinadoView`, `HistorialCargasView`

### Fase 4 — Módulo COMPS (3-5 días)
1. Modelos: `Cortesia`, `PremioComps`, `MesaPuntos`, `SrwJugador`, `Jefatura`, `CategoriaNivel`
2. `CompsImportService` para los 4-5 Excel distintos
3. `CompsRepository` con Dapper (porting 1:1 del SQL de `comps_routes.py`)
4. Views bajo `Views/Comps/`: Resumen, Cortesías, Premios, Control Invitaciones
   (General/MDA/MDJ/MKT), Coin-In Cero, Exportar, Configuración

### Fase 5 — Exportación Excel (½ día)
1. `ExcelExportService` con ClosedXML multi-hoja
2. Guardar en `%USERPROFILE%\Downloads` por defecto
3. Botón "Abrir archivo" tras guardar

### Fase 6 — Tests (1 día)
1. xUnit sobre `Core/Services` con datasets fijos
2. Validación cruzada Python vs .NET (mismo dataset → mismos resultados)

### Fase 7 — Empaquetado (½ día)
1. `dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true`
2. Firma de código (opcional)
3. Inno Setup → `SGOS-Setup.exe`
4. Publicar en GitHub Releases

**Total estimado**: 9-12 días de trabajo enfocado.

---

## 11. Compatibilidad de DB

PostgreSQL en Neon **se mantiene exactamente igual**. La app desktop conecta
directo por SSL desde el PC del usuario. Esto permite:

- Correr la web Flask y la app desktop **en paralelo** durante la migración
- Validar resultados Python vs .NET con los mismos datos en la misma DB
- Rollback inmediato si algo falla

**Tablas a respetar tal cual**:
`users`, `operaciones`, `premios`, `srw_jugadores`, `cortesias`,
`premios_comps`, `mesas_puntos`, `jefaturas`, `categorias_nivel`, `carga_log`.

> ⚠️ **Hashes coexistentes en `users`** (observados en Neon):
> - `scrypt:32768:8:1$...` — 5 usuarios (Werkzeug scrypt)
> - `sgos.pbkdf2.sha256.v1$...` — 1 usuario (admin, formato custom)
>
> `WerkzeugCompatibleHasher` (§7.7) debe soportar **los tres formatos** para
> que cualquier usuario pueda loguearse sin reset.

---

## 12. Consideraciones específicas de desktop

### Conexión SSL a Neon desde `.exe`
Neon requiere TLS. Connection string:
```
Host=...;Database=...;Username=...;Password=...;SslMode=Require;Trust Server Certificate=true
```

### Reconexión / red inestable
Con Polly:
```csharp
services.AddSingleton<IAsyncPolicy>(Policy
    .Handle<NpgsqlException>()
    .WaitAndRetryAsync(3, i => TimeSpan.FromSeconds(Math.Pow(2, i))));
```

### Threading WPF
Operaciones DB/Excel → `async Task.Run`. UI thread solo binding.
`BindingOperations.EnableCollectionSynchronization` para colecciones compartidas.

### Ventana de login modal (App.xaml.cs)
```csharp
protected override void OnStartup(StartupEventArgs e) {
    base.OnStartup(e);
    var login = _services.GetRequiredService<LoginWindow>();
    if (login.ShowDialog() != true) { Shutdown(); return; }
    var main = _services.GetRequiredService<MainWindow>();
    main.Show();
}
```

### Preferencias de usuario
`%APPDATA%\SGOS\user.json` (última carpeta, tema, filtros guardados, etc.).

### Logs
Siempre en `%LOCALAPPDATA%\SGOS\logs\` con rotación diaria (Serilog).

---

## 13. Ventajas vs la versión web Flask

| Criterio | Flask web | C# desktop WPF |
|---|---|---|
| Distribución | Servidor o Python local | `.exe` único |
| Performance SQL grande (COMPS) | Bueno | Mejor (LINQ sobre struct, sin GIL) |
| UI | HTML emulada en navegador | Nativa Windows |
| Instalación | Python + deps + PyInstaller | Inno Setup o copiar `.exe` |
| Debugging | PDB + print | Visual Studio debugger paso a paso |
| Consumo memoria | ~150MB | ~80MB |
| Exportar Excel | Descarga navegador | Guardado directo a `Downloads` |
| Offline | Requiere Flask corriendo | Requiere red para DB (igual) |

---

## 14. Riesgos y mitigaciones

| Riesgo | Mitigación |
|---|---|
| Parsing fechas (pandas vs DateTime) | Tests con dataset real, usar `DateOnly` |
| Floats vs decimals (pandas float64) | Usar `decimal` en C# para todos los montos |
| Hashes Werkzeug incompatibles | `WerkzeugCompatibleHasher` triple (§7.7) |
| ClosedXML lento >100k filas | Streaming con `OpenXmlReader` |
| Conexión Neon cae | Polly retry + banner "Sin conexión" |
| UI bloqueada por query | Todo async + `LoadingOverlay` |
| Red lenta en oficina | Cache local SQLite + sync diferido (fase futura) |

---

## 15. Checklist antes de iniciar

- [ ] Crear repo `sgos-reportes-desktop` en GitHub
- [ ] Copiar este `MIGRACION_CSHARP.md` en la raíz del nuevo repo
- [ ] Instalar **.NET 8 SDK**
- [ ] Instalar **Visual Studio 2022** (workload ".NET Desktop Development")
      o **VS Code** con C# Dev Kit + XAML Tools
- [ ] Guardar connection string Neon en `dotnet user-secrets`
- [ ] Instalar **Inno Setup** para el instalador final
- [ ] Backup de la DB antes de la primera ejecución desktop
- [ ] Mantener la Flask corriendo en paralelo durante la migración

---

## 16. Referencias al código fuente actual

Orden de estudio para el agente que implemente la migración:

1. `sgos_web/app.py` — modelo `User`, rutas Getnet/Premios, auth
2. `sgos_web/engine.py` — pipeline pandas + reportes
3. `sgos_web/comps_routes.py` — SQL grande de COMPS (porting 1:1 a Dapper)
4. `sgos_web/comps_engine.py` — carga de Excel COMPS
5. `sgos_web/templates/*.html` — extraer estructura de tablas y campos (la UI se rehace en XAML)
6. `DESIGN.md` — paleta y tokens (se traducen a `Styles/Colors.xaml`)

---

**Fin del documento.** Desktop Windows puro, `.exe` nativo, sin web,
sin navegador, sin Kestrel. Misma DB, misma lógica, mejor performance
y distribución.
