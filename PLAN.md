# PLAN.md — Minerma Global (app de certificados)

Plan de implementación por etapas. Modo: **Plan** — este documento describe el orden de trabajo, no es código.

## Contexto

- Aplicación web que emite certificados por estudiante y programa.
- Arquitectura de 3 piezas: Bodega (`INSUMOS/`, solo lectura) → Cocina (`app.py`, toda la lógica) → Salón (`templates/`, solo muestra).
- Datos reales ya auditados: **24 estudiantes** en el maestro, **88 filas** de evaluaciones.
- Nota: en una sesión anterior quedó un borrador funcional preliminar (`app.py` + `templates/`). Este plan consiste en **validarlo/refactorizarlo por etapas**: cada fase confirma su parte antes de seguir.

## Orden de construcción

| Etapa | Qué se construye | Cómo se valida |
|---|---|---|
| **0. Verificación del entorno** | Confirmar que Python, Flask y openpyxl están disponibles y que los Excel existen. | Comando `python --version` y `python -c "import flask, openpyxl"` → sin errores. |
| **1. Validación del cruce de datos** | Script `validar_datos.py` (solo lectura) que audita ambos Excel y reporta inconsistencias. | Comparar su reporte con las cifras esperadas (tabla debajo). |
| **2. Backend (cocina) en Flask** | `app.py`: `cargar_maestro()`, `cargar_evaluaciones()`, `calcular_resultados()`, `obtener_modulos()` y rutas mínimas (JSON sin interfaz aún). | `python -c "import app; ..."` → 19 aprobados, 3 participación, 2 sin certificado; casos borde correctos. Rutas HTTP devuelven 200/404. |
| **3. Interfaz (salón)** | `index.html`, `certificado.html`, `static/style.css` (baby blue / baby yellow). Plantillas SOLO muestran variables que calcula `app.py`. | Servidor corriendo; `Invoke-WebRequest` a `/`, `/certificado/<id>`, buscador `?q=` y una ID inexistente (404). Revisión visual. |
| **4. Prueba final y cierre** | Revisión integral + actualización de `AGENTS.md` y `PLAN.md`. | Recorrido completo: listar → abrir varios certificados (aprobado, participación, sin módulos) → buscar → 404. |

## Fase 1 en detalle — validación del cruce de datos

`validar_datos.py` (herramienta, no se tocan los Excel) debe reportar y comprobar:

1. **Totales**: 24 estudiantes en el maestro; 88 filas de evaluaciones.
2. **Unión por `Identificacion`** (texto, coincidencia exacta):
   - IDs en evaluaciones que NO existen en el maestro → esperado: `999880777` (4 filas). Se ignorarán en la emisión.
   - Estudiantes del maestro SIN evaluaciones → esperado: `304560321` `Espinoza Leon Javier`. Queda como "Sin certificado" con 0 módulos.
3. **Grupos `(id, programa)`**: distribución de módulos cursados por grupo → esperado: 8 grupos con 3 módulos, 16 con 4. Sin módulos duplicados dentro de un grupo.
4. **Rangos**: `Nota` y `Asistencia_Pct` dentro de 0–100 en todas las filas → esperado: 0 fuera de rango. Notas contiene `0` en 2 filas y no deben descartarse.
5. **Formato**: `Identificacion` llega como texto; `Fecha_Cierre` es texto ISO (`YYYY-MM-DD`), no fecha Excel.

Criterio de aceptación: el script imprime cada comprobación con `OK`/`FALLO` y termina sin modificar ningún archivo.

## Fase 3 en detalle — interfaces

- `index.html` → buscar: tarjetas resumen (total, aprobados, participación, sin certificado), campos buscador `?q=` (envía a Flask), tabla con estado coloreado y enlace "Ver certificado".
- `certificado.html` → título de certificado según estado, datos de rendimiento, tabla de módulos (o mensaje "sin módulos registrados"), botón volver.
- `style.css` → variables `--baby-blue` y `--baby-yellow`, tarjetas, tablas, bordes decorativos según estado.

## Reglas de negocio fijadas en AGENTS.md

Toda la lógica vive en la cocina (`app.py`); ninguna plantilla decide sola. Criterio de aceptación: el cálculo debe reproducir exactamente **19 Aprobado**, **3 Participación**, **2 Sin certificado** (22 certificados emitidos).

## Ruta de verificación manual (usamos)

```
python app.py                  # arranca en http://127.0.0.1:5000
Invoke-WebRequest http://127.0.0.1:5000/
Invoke-WebRequest http://127.0.0.1:5000/certificado/101230456
Invoke-WebRequest http://127.0.0.1:5000/certificado/304560321   # 0 módulos
Invoke-WebRequest "http://127.0.0.1:5000/?q=mariana"
# ID inexistente debe dar 404
```