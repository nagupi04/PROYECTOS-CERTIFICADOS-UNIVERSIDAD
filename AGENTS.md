# AGENTS.md

Respuesta en español: el asistente SIEMPRE debe responder en español, aunque las instrucciones o el código estén en inglés.

Proyecto Academia Horizonte: aplicación web Flask que emite certificados por estudiante a partir de los Excel de `INSUMOS/` (que NO se modifican, solo se leen). Arquitectura: la lógica vive SOLO en `app.py` (la "cocina"); las plantillas en `templates/` (el "salón") solo muestran datos. Los Excel son datos de muestra (`@ejemplo.cr`).

## Decisiones técnicas (ver PLAN.md para el plan por etapas)

- **Lectura de Excel: openpyxl** (ya instalada, 3.1.5). Descartadas: pandas (pesada, convierte IDs numéricas en números), xlrd (solo `.xls` antiguo).
- **Servidor web: Flask** (ya instalada, 3.1.3, trae Jinja2). Descartados: FastAPI (hecha para APIs JSON) y Django (overkill para 2-3 páginas).
- **Estructura**: `app.py` (cocina: lectura + cálculo + rutas) · `run_server.py` (arranque sin debug) · `iniciar_app.bat` (doble clic) · `templates/index.html` y `templates/certificado.html` (salón) · `static/style.css` (azul marino/dorado) · `PLAN.md`.
- **Orden de construcción**: 1) validar cruce de datos → 2) backend Flask → 3) interfaz → 4) prueba final. Cada etapa se valida antes de seguir.
- **Criterio de aceptación**: el cálculo reproduce 19 Aprobación, 3 Participación, 1 Sin certificado (de 23 estudiantes con evaluaciones; 22 certificados emitidos). Los 3 estudiantes en la sección de INCONSISTENCIAS (sin evaluaciones, huérfanos, módulos incompletos) NO cuentan como "sin certificado".

## Entorno y comandos

- Python 3.12.10 instalado en `C:\Users\Natasha\AppData\Local\Programs\Python\Python312\` — NO está en el PATH persistente; prepende esa ruta (y `...\Scripts`) en cada sesión de PowerShell antes de usar `python`/`pip`.
- Dependencias ya instaladas: Flask 3.1.3 y openpyxl 3.1.5.
- Arrancar: `python app.py` → http://127.0.0.1:5000 (Flask dev server, debug activo). Para probar sin servidor: `python -c "import app; ..."` ejecutando las funciones de `app.py`.
- No hay git ni tests. Verificación manual: `Invoke-WebRequest` a `/` y `/certificado/<id>`.
- `.xlsx` son ZIP; para inspeccionarlos sin Python usar `System.IO.Compression` y leer `xl/worksheets/sheet1.xml` (valores en inline strings, no hay `sharedStrings.xml`).

## Reglas de negocio (implementadas en `app.py`)

- Promedio = suma de Notas / cantidad de módulos cursados. Asistencia = promedio de `Asistencia_Pct`. Límites INCLUYEN el valor (`>=`/`<`).
- Se agrupa por `Identificacion` + `Programa`; un certificado por estudiante y programa.
- Aprobado: Promedio >= 70 Y Asistencia >= 80. Participación: Promedio < 70 Y Asistencia >= 80. Sin certificado: Asistencia < 80.
- Casos de borde (ya manejados): estudiante del maestro sin evaluaciones → Sin certificado con 0 módulos; registros de evaluaciones sin estudiante en el maestro → se ignoran.

## Esquema de datos

`INSUMOS/Maestro_Estudiantes.xlsx` — hoja `Estudiantes`, 1 encabezado + 24 filas:
- `Identificacion` — ID de 9 dígitos como **texto** (nunca número). Llave de unión.
- `Nombre_Completo`, `Correo`, `Programa`, `Cohorte` (formato `YYYY-X`, ej. `2026-A`).

`INSUMOS/Registro_Evaluaciones.xlsx` — hoja `Evaluaciones`, 1 encabezado + 88 filas (una por estudiante × módulo):
- `Identificacion` (texto, une con el maestro), `Programa`, `Modulo` (`Módulo 1`…`Módulo 4`, con `ó` acentuada).
- `Nota` y `Asistencia_Pct` numéricos 0–100; `Nota` incluye valores `0` que NO deben descartarse.
- `Fecha_Cierre` — **texto ISO** (`2026-03-13`), no es celda de fecha Excel. Una fecha por módulo (2026-03-13, 2026-04-17, 2026-05-22, 2026-06-26).

Encabezados en español y acentuados (`Identificacion`, `Módulo`). Coincidir IDs exactamente al unir.

## Convenciones de estilo

- Código simple y comentado en español, marcando los "ladrillos": variables, tipos, condicionales, bucles y funciones.
- Diseño web azul marino y dorado (`static/style.css`, variables `--azul-marino`, `--dorado`).