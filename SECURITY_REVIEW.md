# SGOS – Revisión Técnica y de Seguridad

> Revisión del proyecto **SGOS – Reportes de Operaciones** (Flask + PostgreSQL/Neon).
> Fecha: 2026-04-17.

---

## 1. Stack / Tecnologías detectadas

### Backend
- **Python 3** + **Flask** (monolito en `sgos_web/app.py` + blueprint `sgos_web/comps_routes.py`).
- **Flask-Login** – autenticación de sesión.
- **Flask-WTF / CSRFProtect** – tokens CSRF globales.
- **Flask-Limiter** – rate limiting (backend `memory://`).
- **SQLAlchemy** + **psycopg** sobre **PostgreSQL (Neon)**.
- **Werkzeug** (`generate_password_hash`, `check_password_hash`, `secure_filename`).
- **Pandas / OpenPyXL / XlsxWriter** – procesamiento de Excel.
- **python-dotenv** – variables de entorno.

### Frontend
- **Bootstrap 5.3.3** + **Bootstrap Icons** vía CDN (`cdn.jsdelivr.net`).
- JS propio (`sortable.js`), **Service Worker** (`sw.js`) y `manifest.json` (PWA).
- **Chart.js** (usado en `graphs.html`).

### Empaquetado / despliegue
- **PyInstaller** (`SGOS.spec`, carpeta `build/`) para modo escritorio (`SGOS_DESKTOP=1`).
- `Procfile` + `wsgi.py` → despliegue estilo **Heroku/Render/Railway** con **Gunicorn**.

### Datos
- **PostgreSQL** obligatorio vía `DATABASE_URL`.
- Tablas: `users`, `operaciones`, `premios`, `srw_jugadores`, `cortesias`, `premios_comps`,
  `mesas_puntos`, `jefaturas`, `categorias_nivel`, `carga_log`.

---

## 2. Lo que ya está bien hecho

- Contraseñas con `werkzeug.security` (hash + salt), no texto plano.
- Política de contraseñas razonable (`validar_password`: 9+ chars, mayúscula, símbolo).
- **CSRF** global activo (`CSRFProtect`).
- Cookies: `HttpOnly`, `SameSite=Lax`, `Secure` condicional en producción.
- Cabeceras de seguridad en `@app.after_request`: `X-Content-Type-Options`,
  `X-Frame-Options`, `Referrer-Policy`, `Permissions-Policy`, `HSTS` en producción.
- Validación de extensiones (`.xlsx / .xls`) y `MAX_CONTENT_LENGTH = 20 MB`.
- Prevención de **path traversal** en `safe_file_path` (`secure_filename` + chequeo de prefijo).
- Nombres de archivo aleatorios (`uuid.hex`) y limpieza periódica (`cleanup_old_uploads`).
- **Rate limit** en `/login` (10/min).
- SQL parametrizado con `text(... :param)` en `comps_routes.py` (sin f-strings con input).
- `DATABASE_URL` obligatorio: falla rápido si no está configurado.

---

## 3. Problemas de seguridad / mejoras recomendadas

### 3.1 Críticos

1. **Usuario admin con contraseña por defecto embebida.**
   En `app.py` se crea `admin / admin123` si no existe y solo se imprime un aviso.
   Si alguien despliega sin fijar `ADMIN_DEFAULT_PASSWORD` queda expuesto.
   - Forzar que `ADMIN_DEFAULT_PASSWORD` sea obligatorio (igual que `DATABASE_URL`).
   - Marcar al admin con `must_change_password=True` y obligar cambio en el primer login.

2. **Falta Content-Security-Policy (CSP).**
   Se cargan Bootstrap, Bootstrap-Icons y Chart.js desde `cdn.jsdelivr.net`.
   Sin CSP, cualquier XSS permitiría inyectar scripts externos.
   - Añadir CSP estricta, ej.:
     ```
     default-src 'self';
     script-src 'self' https://cdn.jsdelivr.net;
     style-src  'self' https://cdn.jsdelivr.net 'unsafe-inline';
     img-src    'self' data:;
     object-src 'none';
     base-uri   'self';
     frame-ancestors 'self';
     ```
   - Mejor aún: **descargar los assets** y servirlos desde `/static/` (elimina dependencia del CDN).

3. **XSS potencial por `to_html(...)`.**
   En `tablas_a_html` se genera HTML a partir de Excel subido por usuarios y luego
   se renderiza en los templates. Pandas escapa por defecto, pero hay que confirmar que
   **ningún** template use `|safe` sobre contenido de celdas.
   - Verificar `to_html(escape=True)` explícito.
   - Auditar los templates (`dashboard.html`, `comps/*.html`) por `|safe` sobre datos dinámicos.

4. **Gestión de usuarios sin control de rol.**
   Cualquier usuario autenticado accede a `/usuarios` y puede crear/eliminar/cambiar
   contraseñas de otros (`crear_usuario`, `eliminar_usuario`, `cambiar_password`).
   No existe campo `is_admin`.
   - Añadir `is_admin = db.Column(db.Boolean, default=False)`.
   - Crear decorador `@admin_required`.
   - Proteger rutas de `/usuarios/*` y ocultar el menú a no-admins.

5. **Confirmar CSRF en todos los formularios.**
   `CSRFProtect` cubre globalmente, pero todo `<form method="post">`
   (incluido `login.html`, `usuarios.html`, filtros del dashboard) debe contener
   `{{ csrf_token() }}` o un `<input name="csrf_token" ...>`.

### 3.2 Importantes

6. **Rate limit solo en `/login`.**
   Endpoints sensibles como `/usuarios/crear`, `/usuarios/cambiar-password/<id>`,
   subida a `/sgos` y descargas deberían tener límites (p.ej. `@limiter.limit("30/minute")`).

7. **`Flask-Limiter` con storage `memory://`.**
   No sirve con múltiples workers de Gunicorn: cada proceso cuenta aparte
   → límite efectivo = N × límite. Usar Redis (`storage_uri="redis://..."`).

8. **Logs sin configuración estructurada.**
   `app.logger.error(...)` funciona, pero no hay `logging.config.dictConfig`,
   rotación ni formato consistente. Configurar handlers (stderr/fichero) con nivel y formato.

9. **Validación superficial del contenido de Excel.**
   Solo se valida extensión y tamaño, no **MIME real** ni estructura.
   - Validar con `openpyxl.load_workbook(..., read_only=True)` dentro de try/except.
   - Limitar número de filas procesadas.
   - Considerar chequeo de magic bytes (`PK\x03\x04` para xlsx).

10. **`MAX_CONTENT_LENGTH = 20 MB` a nivel Flask.**
    Bien, pero confirmar que el proxy/host (Nginx, Render, Railway) también limita
    el tamaño del body para evitar **DoS por ancho de banda**.

11. **`FLASK_SECRET_KEY` opcional.**
    Si no está definida se genera `os.urandom(32).hex()` en cada arranque:
    invalida sesiones al reiniciar y **no coincide entre workers**.
    Hacerla obligatoria en producción (igual que `DATABASE_URL`).

12. **Path de `uploads/` relativo al `cwd`.**
    `os.path.abspath(os.getcwd())` depende desde dónde se ejecute la app.
    Usar `os.path.dirname(os.path.abspath(__file__))` para anclaje estable.

13. **`f.save(path)` sin verificar tipo real.**
    Combinar con chequeo de magic bytes o `python-magic` para evitar que se suban
    archivos con extensión .xlsx pero contenido arbitrario.

14. **`/health` sin auth.**
    Correcto que sea público y devuelva solo `"OK"`. No añadir detalles
    (versión, estado DB) sin auth.

15. **`comps_routes.build_date_filter` – nombres de parámetros.**
    Genera nombres con `col.replace('.', '_')`. Si se llama dos veces sobre la misma
    columna en una misma query (UNION), los parámetros colisionan.
    Añadir un sufijo único (contador o uuid corto) por invocación.

16. **Service Worker `/sw.js`.**
    Revisar la estrategia de caché para **excluir** respuestas autenticadas
    (Set-Cookie, rutas privadas). Evitar que queden datos residuales en el cliente.

### 3.3 Menores / Calidad

17. `app.run(debug=...)` solo en `__main__` – bien. Garantizar `FLASK_ENV != development` en prod.
18. Añadir `Cross-Origin-Opener-Policy` y `Cross-Origin-Resource-Policy`.
19. Sin logout por inactividad: configurar `PERMANENT_SESSION_LIFETIME`.
20. Sin bloqueo tras N intentos fallidos ni rotación de contraseñas.
21. No hay suite de tests (`tests/`). Añadir `pytest` cubriendo permisos y rutas críticas.
22. **`requirements.txt` aparenta estar en UTF-16** (bytes `fe ff 62 00 6c 00 ...`).
    Puede romper `pip install` en algunos entornos. Guardar como **UTF-8 sin BOM**.
23. Activar **Dependabot** (GitHub) o correr `pip-audit` en CI.
24. Sustituir `print("Usuario 'admin' creado...")` por `app.logger.warning(...)`.
25. Añadir linter/formatter (**ruff**, **black**) y **pre-commit**.
26. Los campos `mes`/`hora` precalculados en `Operacion`/`Premio` son redundantes:
    se pueden derivar con `date_trunc('month', fecha)` y quedan susceptibles a desincronización.

---

## 4. Plan de acción sugerido (orden de prioridad)

1. Añadir `is_admin` + `@admin_required` y proteger todo `/usuarios/*`.
2. Forzar `FLASK_SECRET_KEY` y `ADMIN_DEFAULT_PASSWORD` como obligatorios en producción.
3. Agregar **CSP** y `COOP/CORP` en `set_security_headers`.
4. Mover assets del CDN a `/static/` (o al menos usar `integrity=` SRI).
5. Cambiar `storage_uri` de Flask-Limiter a **Redis** en producción.
6. Ampliar rate limit a rutas de administración y de subida.
7. Validar contenido real de los Excel (magic bytes + límite de filas).
8. Revisar los templates por `|safe` y confirmar `to_html(escape=True)`.
9. Logs estructurados y rotación.
10. Re-guardar `requirements.txt` en UTF-8 y fijar versiones exactas.
11. Suite mínima de tests (`pytest`) y CI con `pip-audit` / Dependabot.

---

## 5. Opinión general

El proyecto está **por encima del promedio de un Flask interno**: ya tiene CSRF,
hashing, rate limit básico, cabeceras de seguridad, mitigación de path traversal
y CSP parcial vía headers.

Los riesgos reales no son tanto técnicos de bajo nivel, sino:

- **Control de acceso**: cualquier usuario autenticado puede administrar usuarios.
- **Endurecimiento de producción**: CSP, secretos obligatorios, limiter en Redis,
  assets con SRI o auto-hospedados.

Corrigiendo los puntos **1–5 de la sección 4** se obtiene el mayor salto de seguridad
con poco esfuerzo.

### Sugerencia de arquitectura

`app.py` está creciendo (modelos + rutas + utilidades + bootstrap).
Cuando sea posible, separar en:

- `models.py` – modelos SQLAlchemy.
- `auth.py` – login, usuarios, decoradores de rol.
- `uploads.py` – carga/validación de Excel.
- `reports.py` – dashboards y descargas.

Siguiendo el patrón Blueprint que ya se usa en `comps_routes.py`.
Esto facilita **testing**, **revisión de permisos** y mantenimiento a largo plazo.
