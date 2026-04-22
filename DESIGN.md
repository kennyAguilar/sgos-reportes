# SGOS — Guía de Diseño de Interfaz

Sistema de Gestión de Operaciones de Slots (SGOS) · Casino Royale Theme

---

## 1. Identidad Visual

### Concepto
Tema oscuro de alto contraste inspirado en casino de lujo. Fondo profundo azul-noche con acentos dorados y detalles en esmeralda/rojo para estados.

### Paleta de Colores

#### Superficies (fondo en capas)
| Token | Hex | Uso |
|---|---|---|
| `--bg-base` | `#0A0E1A` | Fondo de página principal |
| `--surface-1` | `#141B2D` | Cards, paneles |
| `--surface-2` | `#1C2540` | Cards elevadas, modales |
| `--surface-3` | `#252F50` | Hover sobre cards |

#### Bordes
| Token | Hex | Uso |
|---|---|---|
| `--border-subtle` | `#2A3654` | Bordes por defecto |
| `--border-strong` | `#3A4770` | Bordes en foco/hover |

#### Texto
| Token | Hex | Uso |
|---|---|---|
| `--text-primary` | `#F5F5F7` | Texto principal |
| `--text-secondary` | `#A8B2CE` | Texto secundario, labels |
| `--text-muted` | `#6B7394` | Texto deshabilitado, hints |

#### Acentos de Marca
| Token | Hex | Uso |
|---|---|---|
| `--gold` | `#D4AF37` | Primario, CTAs, links |
| `--gold-hover` | `#E8C252` | Estado hover del dorado |
| `--gold-soft` | `rgba(212,175,55,0.12)` | Fondos de badges dorados |
| `--emerald` | `#10B981` | Éxito, datos positivos |
| `--emerald-hover` | `#34D399` | Hover de éxito |

#### Estados
| Token | Hex | Uso |
|---|---|---|
| `--ruby` | `#DC2626` | Error, eliminación |
| `--ruby-hover` | `#EF4444` | Hover de error |
| `--amber` | `#F59E0B` | Advertencia |
| `--info` | `#60A5FA` | Información |

---

## 2. Tipografía

**Familia principal:** `Plus Jakarta Sans` (Google Fonts)  
**Fallback:** `system-ui, -apple-system, Segoe UI, Roboto, sans-serif`

| Rol | Peso | Tamaño base |
|---|---|---|
| Body | 400 | 15px |
| Labels / UI | 500 | — |
| Subtítulos | 600 | — |
| Títulos | 700 | — |
| Hero / Display | 800 | — |

- `line-height`: 1.55
- `letter-spacing` en headings: `-0.01em`
- Anti-aliasing: `-webkit-font-smoothing: antialiased`

---

## 3. Espaciado (escala de 4px)

| Token | Valor | Uso típico |
|---|---|---|
| `--s-1` | 4px | Micro gaps |
| `--s-2` | 8px | Gaps internos |
| `--s-3` | 12px | Padding pequeño |
| `--s-4` | 16px | Padding estándar |
| `--s-5` | 24px | Secciones pequeñas |
| `--s-6` | 32px | Secciones medianas |
| `--s-7` | 48px | Secciones grandes |
| `--s-8` | 64px | Hero / bloques principales |

---

## 4. Radios y Sombras

### Border Radius
| Token | Valor | Uso |
|---|---|---|
| `--r-sm` | 6px | Badges, inputs pequeños |
| `--r-md` | 10px | Botones, inputs |
| `--r-lg` | 14px | Cards |
| `--r-xl` | 20px | Modales, panels grandes |

### Sombras
| Token | Uso |
|---|---|
| `--shadow-sm` | Elementos en superficie plana |
| `--shadow-md` | Cards flotantes |
| `--shadow-lg` | Modales, dropdowns |
| `--shadow-gold` | Cards destacadas con acento dorado |

---

## 5. Animaciones y Transiciones

**Easing:** `cubic-bezier(0.4, 0, 0.2, 1)` (Material Design standard)

| Token | Valor | Uso |
|---|---|---|
| `--t-fast` | 120ms | Hovers, toggles |
| `--t-base` | 180ms | Transiciones estándar |
| `--t-slow` | 260ms | Entradas de paneles |

**View Transitions API:** activas en Chromium para navegación entre páginas (`same-origin`).

---

## 6. Dependencias de UI

| Librería | Versión | Uso |
|---|---|---|
| Bootstrap | 5.3.3 | Grid, componentes base |
| Bootstrap Icons | 1.11.3 | Iconografía |
| Plus Jakarta Sans | — | Tipografía |
| Chart.js | — | Gráficos de dashboards |

---

## 7. Estructura de Layouts

### `layout.html` — Layout principal
```
<navbar sticky-top>
  Logo "SGOS" + links de navegación + dropdown de usuario
</navbar>

<main class="container py-4">
  {% block content %}
</main>

<footer>
```

### `layout_comps.html` — Layout del módulo COMPS
Variante del layout principal con navbar adaptado al módulo de auditoría.

### `login.html` — Página de autenticación
Layout standalone (sin navbar), centrado en pantalla con `login-page` / `login-card`.

---

## 8. Componentes Principales

### Navbar
- Clase: `.navbar-custom`
- Sticky top, fondo transparente sobre `--bg-base`
- Brand: texto `SGOS` con enlace a home
- Links activos detectados via `request.endpoint`
- Dropdown de usuario con opción de cerrar sesión

### Cards de Módulo (Home)
- Clase: `.module-card`
- Icono grande centrado (Bootstrap Icon)
- Hover con efecto elevación y borde dorado
- Dos módulos: **SGOS Reportes** (azul) y **Auditoría COMPS** (verde)

### Hero Section
- Clase: `.hero`
- Título H1 con ícono + subtítulo descriptivo
- Fondo con gradiente sutil

### Tablas de Datos
- Bootstrap `.table` con clase `.table-dark` o custom dark styling
- Ordenamiento client-side via `sortable.js`
- Filas con hover highlight

### Badges de Estado
Usando los colores de estado del sistema:
- `--emerald` → operaciones exitosas, montos altos
- `--ruby` → errores, rechazos
- `--amber` → advertencias, pendientes
- `--gold` → destacados, primarios

### Formularios
- Inputs con fondo `--surface-2`, borde `--border-subtle`
- Focus ring en `--gold-ring`
- Botón primario: `.btn-primary` → fondo dorado
- `.btn-primary-full` → ancho completo (login)

---

## 9. Módulos y Páginas

### Módulo SGOS Reportes
| Página | Ruta | Descripción |
|---|---|---|
| Home | `/home` | Panel de selección de módulos |
| Cargar Archivo | `/` | Upload de Excel Getnet o Premios |
| Histórico Getnet | `/dashboard` | Tabla paginada de operaciones |
| Dashboard de Reportes | `/graphs` | Gráficos de Getnet |
| Históricos de Premios | `/dashboard-premios` | Tabla paginada de premios |
| Dashboard de Premios | `/graphs-premios` | Gráficos de premios |
| Getnet vs Premios | `/dashboard-combinado` | Comparativo combinado |
| Historial de Cargas | `/cargas` | Log de archivos procesados |
| Usuarios | `/usuarios` | Gestión de usuarios (admin) |

### Módulo Auditoría COMPS (`/comps`)
| Página | Descripción |
|---|---|
| Index | Resumen de COMPS |
| Análisis de Resumen | KPIs generales |
| Análisis de Cortesías | Cortesías por jugador/fecha |
| Análisis de Premios | Premios COMPS |
| Auditoría Coin-In Cero | Jugadores sin actividad |
| Control Invitaciones MDA | Invitaciones mesas de dados |
| Control Invitaciones MDJ | Invitaciones mesa de juego |
| Control Invitaciones MKT | Invitaciones marketing |
| Control Invitaciones | Vista general |
| Configuración | Settings del módulo |
| Exportar | Exportación a Excel |

---

## 10. Fondo Decorativo

El `body::before` aplica un gradiente radial sutil de dos puntos:
- Esquina superior derecha: destello dorado tenue (`rgba(212,175,55,0.06)`)
- Esquina inferior izquierda: destello esmeralda tenue (`rgba(16,185,129,0.05)`)

Esto da profundidad sin distraer del contenido.

---

## 11. PWA (Progressive Web App)

- Manifest: `/static/manifest.webmanifest`
- Service Worker: `/static/js/sw.js`
- Theme color: `#0A0E1A`
- Soporte iOS: `apple-mobile-web-app-capable`, icono 192px

---

## 12. Archivos de Diseño

| Archivo | Ruta |
|---|---|
| CSS principal | `sgos_web/static/css/style.css` |
| Layout base | `sgos_web/templates/layout.html` |
| Layout COMPS | `sgos_web/templates/comps/layout_comps.html` |
| Charts JS | `sgos_web/static/js/sgos-charts.js` |
| UI JS | `sgos_web/static/js/sgos-ui.js` |
| Sortable tables | `sgos_web/static/js/sortable.js` |
