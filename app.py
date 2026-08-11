# ============================================================
# ACADEMIA HORIZONTE - Sistema de certificados
# ============================================================
# "COCINA" (backend): aqui vive TODA la logica de negocio.
# El frontend (HTML) solo muestra lo que esta funcion calcula.
# ============================================================

# ---------- IMPORTACIONES ----------
# Flask: crea el servidor web. openpyxl: lee los archivos Excel.
# pathlib: rutas que funcionan desde cualquier carpeta (produccion).
from flask import Flask, render_template, request, abort
from openpyxl import load_workbook
from pathlib import Path

# ---------- VARIABLES GLOBALES (ladrillos: datos fijos) ----------
# BASE_DIR: carpeta donde vive este archivo, sin importar desde donde
# se ejecute la app (local o servidor de produccion).
BASE_DIR = Path(__file__).resolve().parent

# Rutas de los archivos Excel de entrada (la "bodega").
# IMPORTANTE: estos archivos NO se modifican, solo se leen.
ARCHIVO_MAESTRO = BASE_DIR / "INSUMOS" / "Maestro_Estudiantes.xlsx"
ARCHIVO_EVALUACIONES = BASE_DIR / "INSUMOS" / "Registro_Evaluaciones.xlsx"

# Modulos maximos esperados por programa (para detectar incompletos)
CANTIDAD_MODULOS_ESPERADOS = 4

# Limites de aprobacion (los limites INCLUYEN el valor: >= o <)
NOTA_MINIMA = 70          # Promedio minimo para aprobar
ASISTENCIA_MINIMA = 80    # Asistencia minima para recibir certificado

# Nombres de los estados (clave interna sin acento) y sus etiquetas
ESTADO_APROBADO = "Aprobado"
ESTADO_PARTICIPACION = "Participacion"
ESTADO_SIN_CERTIFICADO = "Sin certificado"

# Etiqueta con acentos para mostrar en la pagina
ETIQUETAS = {
    ESTADO_APROBADO: "Aprobacion",
    ESTADO_PARTICIPACION: "Participacion",
    ESTADO_SIN_CERTIFICADO: "Sin certificado",
}

# Creamos la aplicacion Flask
app = Flask(__name__)


# ---------- FUNCIONES DE LA COCINA (ladrillos: funciones) ----------

def cargar_maestro():
    """Lee el Excel de estudiantes y devuelve un diccionario.

    Estructura de salida (tipo: dict):
      { "101230456": { "nombre": "...", "correo": "...",
                       "programa": "...", "cohorte": "2026-A" }, ... }
    """
    # Abrimos el libro y elegimos la hoja "Estudiantes"
    libro = load_workbook(ARCHIVO_MAESTRO, read_only=True, data_only=True)
    hoja = libro["Estudiantes"]

    # Variable de salida (tipo: dict) que iremos llenando
    estudiantes = {}

    # BUCLE: recorremos las filas, la fila 1 es el encabezado
    primera = True
    for fila in hoja.iter_rows(values_only=True):
        if primera:
            primera = False
            continue

        # Cada fila viene como una tupla:
        # (Identificacion, Nombre_Completo, Correo, Programa, Cohorte)
        identificacion = str(fila[0])  # la ID es TEXTO, nunca numero
        if not identificacion or identificacion == "None":
            continue

        # VARIABLES locales: cada columna de la fila
        nombre = fila[1]
        correo = fila[2]
        programa = fila[3]
        cohorte = fila[4]

        # Guardamos al estudiante en el diccionario, usando la ID como llave
        estudiantes[identificacion] = {
            "nombre": nombre,
            "correo": correo,
            "programa": programa,
            "cohorte": cohorte,
        }

    libro.close()
    return estudiantes


def cargar_evaluaciones():
    """Lee el Excel de evaluaciones y devuelve una lista.

    Estructura de salida (tipo: list de dicts), una entrada por
    estudiante y modulo:
      [ { "identificacion": "101230456", "programa": "...",
          "modulo": "Modulo 1", "nota": 95.0, "asistencia": 100.0,
          "fecha_cierre": "2026-03-13" }, ... ]
    """
    libro = load_workbook(ARCHIVO_EVALUACIONES, read_only=True, data_only=True)
    hoja = libro["Evaluaciones"]

    # Variable de salida (tipo: list) que iremos llenando
    evaluaciones = []

    primera = True
    for fila in hoja.iter_rows(values_only=True):
        if primera:
            primera = False
            continue

        # Tupla: (Identificacion, Programa, Modulo, Nota, Asistencia_Pct,
        #         Fecha_Cierre)
        identificacion = str(fila[0])
        if not identificacion or identificacion == "None":
            continue

        evaluaciones.append({
            "identificacion": identificacion,
            "programa": fila[1],
            "modulo": fila[2],
            "nota": float(fila[3]),          # numerico 0-100
            "asistencia": float(fila[4]),    # numerico 0-100
            "fecha_cierre": fila[5],         # texto ISO "2026-03-13"
        })

    libro.close()
    return evaluaciones


def calcular_resultados(estudiantes, evaluaciones):
    """Agrupa por Identificacion + Programa y calcula el estado.

    Reglas de negocio:
      - Promedio  = suma de Notas / cantidad de modulos cursados
      - Asistencia = promedio de Asistencia_Pct
      - Aprobado:         Promedio >= 70  Y  Asistencia >= 80
      - Participacion:    Promedio < 70   Y  Asistencia >= 80
      - Sin certificado:  Asistencia < 80

    Solo participan quienes TIENEN evaluaciones. Los estudiantes del
    maestro sin datos se reportan como inconsistencia (no aqui).
    Salida: list de dicts con el resumen por estudiante y programa.
    """
    # Diccionario intermedio: agrupamos todas las evaluaciones
    # por la llave (identificacion, programa)
    grupos = {}
    for eva in evaluaciones:
        llave = (eva["identificacion"], eva["programa"])
        # CONJUNTO de datos si es la primera vez que vemos esta llave
        if llave not in grupos:
            grupos[llave] = {"notas": [], "asistencias": []}
        # Añadimos la nota y la asistencia de esta fila al grupo
        grupos[llave]["notas"].append(eva["nota"])
        grupos[llave]["asistencias"].append(eva["asistencia"])

    # Variable de salida (tipo: list)
    resultados = []

    # BUCLE: procesamos cada grupo (estudiante + programa)
    for (identificacion, programa), datos in grupos.items():
        # CONDICIONAL: el estudiante debe existir en el maestro
        # (los registros huerfanos de evaluaciones se ignoran)
        if identificacion not in estudiantes:
            continue
        estudiante = estudiantes[identificacion]

        # ---------- CALCULOS (reglas de negocio) ----------
        cantidad_modulos = len(datos["notas"])  # cuantos modulos curso

        # Promedio = suma de Notas / cantidad de modulos cursados
        promedio = sum(datos["notas"]) / cantidad_modulos

        # Asistencia = promedio de Asistencia_Pct
        asistencia = sum(datos["asistencias"]) / cantidad_modulos

        # ---------- DECISION (condicionales encadenados) ----------
        # Los limites INCLUYEN el valor, por eso usamos >= y <
        if promedio >= NOTA_MINIMA and asistencia >= ASISTENCIA_MINIMA:
            estado = ESTADO_APROBADO
        elif promedio < NOTA_MINIMA and asistencia >= ASISTENCIA_MINIMA:
            estado = ESTADO_PARTICIPACION
        else:
            estado = ESTADO_SIN_CERTIFICADO  # asistencia < 80

        # Redondeamos a 2 decimales y añadimos la etiqueta para mostrar
        resultados.append({
            "identificacion": identificacion,
            "nombre": estudiante["nombre"],
            "correo": estudiante["correo"],
            "programa": programa,
            "cohorte": estudiante["cohorte"],
            "modulos_cursados": cantidad_modulos,
            "promedio": round(promedio, 2),
            "asistencia": round(asistencia, 2),
            "estado": estado,
            "etiqueta": ETIQUETAS[estado],
        })

    # Ordenamos por nombre para la lista
    resultados.sort(key=lambda r: r["nombre"])
    return resultados


def detectar_inconsistencias(estudiantes, evaluaciones):
    """Audita el cruce de datos (bodega) y detecta problemas.

    Solo lectura: no modifica los Excel. Devuelve una lista de
    dicts, uno por cada tipo de inconsistencia encontrada:
      { "titulo": str, "explicacion": str, "items": [str, ...] }
    """
    # Variable de salida (tipo: list)
    inconsistencias = []

    # --- 1) Estudiantes del maestro SIN evaluaciones registradas ---
    # CONJUNTO de ID que al menos tienen una fila de evaluaciones
    ids_con_evaluaciones = {e["identificacion"] for e in evaluaciones}
    sin_evaluaciones = []
    for identificacion, estudiante in estudiantes.items():
        if identificacion not in ids_con_evaluaciones:
            sin_evaluaciones.append(
                estudiante["nombre"] + " (" + str(identificacion) + ")")
    if sin_evaluaciones:
        inconsistencias.append({
            "titulo": "Estudiantes sin evaluaciones registradas",
            "explicacion": "Estan en el maestro pero no cursaron modulos: "
                           "no reciben certificado.",
            "detalle": sin_evaluaciones,
        })

    # --- 2) Evaluaciones sin estudiante en el maestro ---
    huerfanas = {}
    for e in evaluaciones:
        if e["identificacion"] not in estudiantes:
            # CONJUNTO: contamos cuantas filas tiene cada ID desconocida
            huerfanas[e["identificacion"]] = huerfanas.get(
                e["identificacion"], 0) + 1
    if huerfanas:
        inconsistencias.append({
            "titulo": "Evaluaciones sin estudiante en el maestro",
            "explicacion": "Las filas de estas ID se ignoran: no se emite "
                           "certificado a ID desconocidas.",
            "detalle": ["ID " + i + " con " + str(n) + " fila(s)"
                      for i, n in sorted(huerfanas.items())],
        })

    # --- 3) Grupos con modulos incompletos (menos de lo esperado) ---
    grupos = {}
    for e in evaluaciones:
        llave = (e["identificacion"], e["programa"])
        if llave not in grupos:
            grupos[llave] = []
        grupos[llave].append(e["modulo"])

    incompletos = []
    for (identificacion, programa), modulos in grupos.items():
        if len(modulos) < CANTIDAD_MODULOS_ESPERADOS:
            nombre = estudiantes.get(identificacion, {}).get("nombre",
                                                             "ID desconocida")
            incompletos.append(
                nombre + " (" + str(identificacion) + "): "
                + str(len(modulos)) + " de " + str(CANTIDAD_MODULOS_ESPERADOS)
                + " modulos")
    if incompletos:
        inconsistencias.append({
            "titulo": "Estudiantes con modulos incompletos",
            "explicacion": "El promedio se calcula SOLO con los modulos "
                           "cursados, no con el total esperado.",
            "detalle": incompletos,
        })

    return inconsistencias


def obtener_modulos(identificacion, programa, evaluaciones):
    """Devuelve las evaluaciones individuales de un estudiante y programa.

    Se usa en el certificado para mostrar la tabla de modulos.
    """
    modulos = []
    for eva in evaluaciones:
        if (eva["identificacion"] == identificacion
                and eva["programa"] == programa):
            modulos.append(eva)
    # Ordenamos por modulo (Modulo 1, Modulo 2, ...)
    modulos.sort(key=lambda m: m["modulo"])
    return modulos


def listar_programas(estudiantes):
    """Devuelve los programas del maestro, ordenados (para el filtro)."""
    # VARIABLES (tipo: set) de programas unicos
    programas = {e["programa"] for e in estudiantes.values()}
    # Devolvemos una lista ordenada
    return sorted(programas)


# ---------- RUTAS DEL SERVIDOR (lo que ve el usuario) ----------

@app.route("/")
def index():
    """Pagina principal: contadores, filtro, tabla e inconsistencias."""
    estudiantes = cargar_maestro()
    evaluaciones = cargar_evaluaciones()
    resultados = calcular_resultados(estudiantes, evaluaciones)
    inconsistencias = detectar_inconsistencias(estudiantes, evaluaciones)

    # CONTADORES GLOBALES (se calculan con TODOS, antes del filtro)
    total = len(resultados)
    aprobados = sum(1 for r in resultados
                    if r["estado"] == ESTADO_APROBADO)
    participacion = sum(1 for r in resultados
                        if r["estado"] == ESTADO_PARTICIPACION)
    sin_certificado = sum(1 for r in resultados
                          if r["estado"] == ESTADO_SIN_CERTIFICADO)

    # Filtro por programa: ?programa=...  (opcional)
    # BUCLE que deja pasar solo los resultados del programa elegido
    programa_filtro = request.args.get("programa", "")
    if programa_filtro:
        resultados = [r for r in resultados
                      if r["programa"] == programa_filtro]

    # Lista de programas para el desplegable (los del maestro)
    programas = listar_programas(estudiantes)

    return render_template(
        "index.html",
        resultados=resultados,
        programas=programas,
        programa_filtro=programa_filtro,
        inconsistencias=inconsistencias,
        total=total,
        aprobados=aprobados,
        participacion=participacion,
        sin_certificado=sin_certificado,
    )


@app.route("/certificado/<identificacion>")
def certificado(identificacion):
    """Pagina de certificado individual de un estudiante."""
    estudiantes = cargar_maestro()
    evaluaciones = cargar_evaluaciones()
    resultados = calcular_resultados(estudiantes, evaluaciones)

    # BUCLE: buscamos el estudiante que coincida con la URL
    resumen = None
    for r in resultados:
        if r["identificacion"] == identificacion:
            resumen = r
            break

    # CONDICIONAL: si la ID no existe, pagina 404
    if resumen is None:
        abort(404, "Estudiante no encontrado")

    # CONJUNTO de modulos con sus notas para la tabla del certificado
    modulos = obtener_modulos(identificacion, resumen["programa"],
                              evaluaciones)

    return render_template(
        "certificado.html",
        resumen=resumen,
        modulos=modulos,
    )


# ---------- ARRANQUE DEL PROGRAMA ----------
# Solo se ejecuta si corremos "python app.py" directamente
if __name__ == "__main__":
    app.run(debug=True)