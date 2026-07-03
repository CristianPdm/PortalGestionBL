# Changelog — Gestión Stock Fiscal | GS

Todos los cambios relevantes del proyecto ordenados cronológicamente.
Formato basado en [Keep a Changelog](https://keepachangelog.com/es/).

---

## [1.3] — 2026-07-03

### Agregado
- `crear_db.py` — script independiente para inicializar la base de datos sin necesidad de levantar el portal. Soporta `--ruta` para crear la DB en una ruta distinta. No pisa datos si la DB ya existe.
- `cambios.bat` — script para registrar y subir cambios a GitHub con un solo comando (`cambios.bat "descripción"`).
- `setup_venv.bat` — script para crear el entorno virtual Python, instalar dependencias e inicializar la base de datos en una PC nueva.
- `.gitignore` — excluye `venv/`, `*.db`, exportaciones Excel y archivos temporales de Python.
- `CHANGELOG.md` — este archivo.

### Infraestructura
- Repositorio Git iniciado en `PortalGestionBL` y publicado en GitHub (`CristianPdm/PortalGestionBL`).
- Carpeta limpia `PortalGestionBL` separada de la carpeta de desarrollo, pensada para sincronizar con el servidor.

---

## [1.2] — 2026-07-03

### Cambiado
- Nombre del portal renombrado de **"GestionStockBL Portal"** a **"Gestión Stock Fiscal | GS"** en todos los archivos (`app.py`, `index.html`, `login.html`, `install.bat`, `start.bat`).

### Agregado
- `Instalacion_Linux_Produccion.md` — guía paso a paso para desplegar en servidor Linux: usuario de servicio, entorno virtual, montaje del recurso SMB de ARCA vía CIFS (`/etc/fstab`), servicio systemd con reinicio automático, apertura de firewall, y nginx como proxy reverso opcional.
- Paquete de despliegue `Gestion_Stock_Fiscal_GS_Deploy.zip` reconstruido con código y nombre actualizados.

---

## [1.1] — 2026-06-26

### Corregido — Compatibilidad SAP en exportación Excel
- **Problema:** el archivo `.xlsx` generado por el portal era rechazado por el importador de SAP. El mismo archivo abierto y re-guardado en Excel sí funcionaba.
- **Causa raíz:** `openpyxl >= 3.x` escribe celdas de texto como "inline strings" (`t="inlineStr"`) en vez de la tabla `sharedStrings.xml` estándar de Office, y deja `[Content_Types].xml` fuera del orden esperado dentro del ZIP. El lector de SAP es más estricto que Excel al respecto.
- **Solución:** migración de `openpyxl` a `xlsxwriter` en la función `_generar_excel()`. `xlsxwriter` genera la estructura "clásica" de Excel compatible con SAP sin intervención manual.
- Actualizado `requirements.txt`: `xlsxwriter>=3.2` reemplaza a `openpyxl>=3.1`.
- Actualizado `install.bat` con la nueva dependencia.

---

## [1.0] — 2026-06-22

### Agregado — Portal Flask + SQLite (reemplazo de la app HTML standalone)
- `app.py` — backend Flask completo:
  - Autenticación con sesiones y roles (`admin`, `comex`, `deposito`, `lectura`).
  - CRUD de BLs con mercaderías y contenedores asociados (cascade delete).
  - Gestión de contenedores vacíos independiente.
  - Locks por sección para edición concurrente sin conflictos.
  - Audit log de todas las operaciones.
  - Generación de Excel con 5 hojas en el formato ARCA.
  - Exportación automática diaria vía APScheduler (hora configurable).
  - Log de exportaciones con historial.
  - Endpoint `/api/maestros` que sirve países, embalajes y aduanas desde `DB_DatosMaestros.csv`.
  - Importación de datos históricos desde `DB_Comex.csv` (migración única).
  - Servidor de producción con `waitress` (8 threads), fallback al servidor de desarrollo de Flask.
- `templates/index.html` — interfaz principal:
  - Grilla de BLs con filtros por estado y validación.
  - Wizard de carga de 4 pasos: Datos BL → Mercadería → Contenedores → Validar.
  - Vista de lectura expandible por BL sin abrir el wizard.
  - Semáforo de completitud por BL (verde / naranja / rojo).
  - Reporte PDF con jsPDF (todas las hojas del Excel en formato visual).
  - Gestión de usuarios (admin).
  - Configuración de carpeta de ARCA y hora de envío automático.
- `templates/login.html` — pantalla de login.
- `portal_config.json` — configuración runtime (carpeta ARCA, hora de envío, aduana por defecto).
- `requirements.txt` — dependencias: `flask`, `xlsxwriter`, `apscheduler`, `waitress`.
- `install.bat` / `start.bat` — instalación y arranque en Windows.

### Corregido durante desarrollo inicial
- Estructura HTML: eliminado `</div>` huérfano, corregida clase `impDescGroup`.
- JS: referencias a `f-cond_merc2` corregidas a `f-cond-merc-tab`.
- Validaciones: CUIT (dígito verificador), fecha no futura, campos requeridos, validación cruzada embalaje 05 ↔ contenedor.
- `cond_merc` movida de nivel BL a nivel línea de mercadería.
- Carpeta de exportación SAP: se crea automáticamente si no existe.
- Campo `estado` incluido correctamente en la exportación Excel.
- Vista de lectura (full-view) para rol `lectura` y `deposito`.
- Rol `deposito` puede editar contenedores sin modificar datos del BL.
- Permisos por rol alineados al spec: `comex` carga BL, `deposito` gestiona contenedores y ubicación, `lectura` solo visualiza.

---

## Contexto del proyecto

**Antes del Portal:** la gestión de stock de BL se hacía con una aplicación HTML standalone (`GestionStockBL_v2` a `v6`), sin servidor, persistiendo datos en CSV mediante la File System Access API del browser. Limitaciones: un solo usuario simultáneo (lock por `_LOCK.json`), dependencia de unidades de red mapeadas (el browser no acepta rutas UNC), y fragilidad ante cambios de PC o browser.

**El Portal** reemplaza esa arquitectura por un backend centralizado con base de datos real, accesible desde cualquier PC de la red sin configuración local.
