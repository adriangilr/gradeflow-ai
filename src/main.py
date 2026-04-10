from __future__ import annotations

import csv
import io
import os
import re
import shutil
from datetime import datetime, timedelta, timezone
from typing import Any

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

from src.auth import get_credentials
from src.config import ensure_directories, get_settings


# ==========================================================
# Configuración general
# ==========================================================

# Si quieres considerar "recientes" como 30 días, deja 30.
DIAS_RECIENTES = 30


# ==========================================================
# Utilidades generales
# ==========================================================

def limpiar_nombre_archivo(texto: str) -> str:
    """
    Limpia nombres de archivos o carpetas para evitar errores
    por caracteres problemáticos.
    """
    if not texto:
        return "sin_nombre"

    texto = texto.strip()
    reemplazos = {
        "/": "-",
        "\\": "-",
        ":": "-",
        "*": "-",
        "?": "",
        '"': "",
        "<": "(",
        ">": ")",
        "|": "-",
    }

    for viejo, nuevo in reemplazos.items():
        texto = texto.replace(viejo, nuevo)

    texto = re.sub(r"\s+", " ", texto).strip()
    return texto or "sin_nombre"


def asegurar_directorio(path: str) -> None:
    """
    Crea el directorio si no existe.
    """
    os.makedirs(path, exist_ok=True)


def construir_nombre_carpeta_entrega(
    submission: dict[str, Any],
    perfil: dict[str, str],
) -> str:
    """
    Construye un nombre de carpeta único por entrega usando:
    apellido_nombre_userId_submissionId
    """
    nombre = limpiar_nombre_archivo(perfil.get("nombre", "") or "sin_nombre")
    apellido = limpiar_nombre_archivo(perfil.get("apellido", "") or "sin_apellido")
    user_id = limpiar_nombre_archivo(str(submission.get("userId", "sin_userId")))
    submission_id = limpiar_nombre_archivo(str(submission.get("id", "sin_submissionId")))

    partes = [apellido, nombre, user_id, submission_id]
    carpeta = "_".join([p for p in partes if p])

    return carpeta or f"entrega_{submission_id}"


def seleccionar_opcion(lista: list[dict[str, Any]], tipo: str) -> dict[str, Any]:
    """
    Menú interactivo genérico para terminal.
    """
    print(f"\nSelecciona {tipo}:\n")

    for i, item in enumerate(lista, start=1):
        print(f"{i}. {item['display_name']}")

    while True:
        entrada = input("\nIngresa número: ").strip()

        try:
            indice = int(entrada) - 1
            if 0 <= indice < len(lista):
                return lista[indice]
        except ValueError:
            pass

        print("❌ Opción inválida. Intenta de nuevo.")


def parse_google_datetime(value: str | None) -> datetime | None:
    """
    Convierte timestamps de Google tipo:
    2026-04-09T18:25:43.123Z
    """
    if not value:
        return None

    try:
        if value.endswith("Z"):
            value = value.replace("Z", "+00:00")
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def ahora_utc() -> datetime:
    """
    Regresa la fecha actual en UTC.
    """
    return datetime.now(timezone.utc)


# ==========================================================
# Nombres visibles para menús
# ==========================================================

def obtener_nombre_curso_visible(course: dict[str, Any]) -> str:
    """
    Construye un nombre visible único para evitar confusiones
    cuando existen dos cursos con el mismo nombre.
    """
    name = course.get("name", "Curso sin nombre")
    section = course.get("section", "").strip()
    course_id = course.get("id", "").strip()
    room = course.get("room", "").strip()

    extras = []
    if section:
        extras.append(section)
    if room:
        extras.append(f"Aula: {room}")
    if course_id:
        extras.append(f"id={course_id}")

    return f"{name} | {' | '.join(extras)}" if extras else name


def obtener_nombre_actividad_visible(coursework: dict[str, Any]) -> str:
    """
    Construye un nombre más legible para la actividad.
    """
    title = coursework.get("title", "Actividad sin título")
    coursework_id = coursework.get("id", "")
    work_type = coursework.get("workType", "")
    max_points = coursework.get("maxPoints")

    extras = []
    if work_type:
        extras.append(work_type)
    if max_points is not None:
        extras.append(f"{max_points} pts")
    if coursework_id:
        extras.append(f"id={coursework_id}")

    return f"{title} | {' | '.join(extras)}" if extras else title


# ==========================================================
# Menús
# ==========================================================

def seleccionar_alcance_descarga() -> str:
    """
    Permite elegir si se descargará una actividad
    o todas las actividades del curso.
    """
    opciones = [
        {
            "id": "single_coursework",
            "display_name": "Descargar una sola actividad",
        },
        {
            "id": "all_courseworks",
            "display_name": "Descargar todas las actividades del curso",
        },
    ]
    return seleccionar_opcion(opciones, "el alcance de la descarga")["id"]


def seleccionar_modo_descarga() -> str:
    """
    Define qué entregas se descargarán.
    """
    opciones = [
        {"id": "all", "display_name": "Bajar todas las entregas activas (TURNED_IN)"},
        {
            "id": "late_ungraded",
            "display_name": "Bajar solo tardías y no evaluadas",
        },
    ]
    return seleccionar_opcion(opciones, "un modo de descarga")["id"]


def seleccionar_formato_salida() -> str:
    """
    Define si solo se guarda en carpeta o también en zip.
    """
    opciones = [
        {"id": "folder_only", "display_name": "Guardar solo en carpeta"},
        {"id": "zip_and_folder", "display_name": "Guardar en carpeta y generar .zip"},
    ]
    return seleccionar_opcion(opciones, "un formato de salida")["id"]


def seleccionar_filtro_actividades() -> str:
    """
    Permite aplicar un filtro a las actividades del curso
    antes de elegir una o antes de procesarlas todas.
    """
    opciones = [
        {"id": "all", "display_name": "Todas las actividades"},
        {"id": "with_submissions", "display_name": "Solo actividades con entregas"},
    ]
    return seleccionar_opcion(opciones, "un filtro de actividades")["id"]


def describir_modo_descarga(modo_descarga: str) -> str:
    descripciones = {
        "all": "todas las entregas activas",
        "resubmitted": "solo reentregadas y pendientes de reevaluación",
        "ungraded": "solo no evaluadas",
        "late": "solo tardías",
        "resubmitted_ungraded": "solo reentregadas y no evaluadas",
        "late_ungraded": "solo tardías y no evaluadas",
    }
    return descripciones.get(modo_descarga, "filtro desconocido")


# ==========================================================
# Descarga de Drive
# ==========================================================

def download_file(drive_service, file_id: str, file_name: str, folder: str) -> str | None:
    """
    Descarga un archivo de Drive al folder indicado.
    """
    asegurar_directorio(folder)

    safe_name = limpiar_nombre_archivo(file_name)
    file_path = os.path.join(folder, safe_name)

    try:
        request = drive_service.files().get_media(fileId=file_id)
        with io.FileIO(file_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request)

            done = False
            while not done:
                _, done = downloader.next_chunk()

        print(f"      ✅ Descargado: {safe_name}")
        return file_path

    except HttpError as err:
        print(f"      ❌ Error al descargar '{safe_name}': {err}")
        return None


# ==========================================================
# Lectura de Classroom
# ==========================================================

def obtener_todos_los_cursos(classroom_service) -> list[dict[str, Any]]:
    """
    Recupera todos los cursos paginando.
    Si existen cursos activos, prioriza esos.
    """
    courses: list[dict[str, Any]] = []
    page_token = None

    while True:
        response = (
            classroom_service.courses()
            .list(pageSize=100, pageToken=page_token)
            .execute()
        )

        courses.extend(response.get("courses", []))
        page_token = response.get("nextPageToken")

        if not page_token:
            break

    cursos_activos = [c for c in courses if c.get("courseState", "ACTIVE") == "ACTIVE"]
    cursos_finales = cursos_activos if cursos_activos else courses

    for course in cursos_finales:
        course["display_name"] = obtener_nombre_curso_visible(course)

    cursos_finales.sort(key=lambda x: x.get("display_name", "").lower())
    return cursos_finales


def obtener_todas_las_actividades(classroom_service, course_id: str) -> list[dict[str, Any]]:
    """
    Recupera todas las actividades de un curso.
    """
    courseworks: list[dict[str, Any]] = []
    page_token = None

    while True:
        response = (
            classroom_service.courses()
            .courseWork()
            .list(courseId=course_id, pageSize=100, pageToken=page_token)
            .execute()
        )

        courseworks.extend(response.get("courseWork", []))
        page_token = response.get("nextPageToken")

        if not page_token:
            break

    for coursework in courseworks:
        coursework["display_name"] = obtener_nombre_actividad_visible(coursework)

    courseworks.sort(key=lambda x: x.get("display_name", "").lower())
    return courseworks


def obtener_todas_las_entregas(
    classroom_service,
    course_id: str,
    coursework_id: str,
) -> list[dict[str, Any]]:
    """
    Recupera todas las entregas de una actividad.
    """
    submissions: list[dict[str, Any]] = []
    page_token = None

    while True:
        response = (
            classroom_service.courses()
            .courseWork()
            .studentSubmissions()
            .list(
                courseId=course_id,
                courseWorkId=coursework_id,
                pageSize=100,
                pageToken=page_token,
            )
            .execute()
        )

        submissions.extend(response.get("studentSubmissions", []))
        page_token = response.get("nextPageToken")

        if not page_token:
            break

    return submissions


# ==========================================================
# Filtros de entregas
# ==========================================================

def ya_fue_devuelta_antes(submission: dict[str, Any]) -> bool:
    """
    Detecta si la entrega ya había sido devuelta antes.
    """
    for event in submission.get("submissionHistory", []):
        state_history = event.get("stateHistory", {})
        if state_history.get("state") == "RETURNED":
            return True
    return False


def es_reentregada(submission: dict[str, Any]) -> bool:
    return (
        submission.get("state") == "TURNED_IN"
        and ya_fue_devuelta_antes(submission)
    )


def es_no_evaluada(submission: dict[str, Any]) -> bool:
    return (
        submission.get("state") == "TURNED_IN"
        and submission.get("assignedGrade") is None
    )


def es_tardia(submission: dict[str, Any]) -> bool:
    return (
        submission.get("state") == "TURNED_IN"
        and submission.get("late", False) is True
    )


def filtrar_entregas(submissions: list[dict[str, Any]], modo_descarga: str) -> list[dict[str, Any]]:
    """
    Aplica el filtro elegido a la lista de entregas.
    """
    if modo_descarga == "all":
        return [s for s in submissions if s.get("state") == "TURNED_IN"]

    if modo_descarga == "resubmitted":
        return [s for s in submissions if es_reentregada(s)]

    if modo_descarga == "ungraded":
        return [s for s in submissions if es_no_evaluada(s)]

    if modo_descarga == "late":
        return [s for s in submissions if es_tardia(s)]

    if modo_descarga == "resubmitted_ungraded":
        return [s for s in submissions if es_reentregada(s) and es_no_evaluada(s)]

    if modo_descarga == "late_ungraded":
        return [s for s in submissions if es_tardia(s) and es_no_evaluada(s)]

    return []


# ==========================================================
# Filtros de actividades
# ==========================================================

def actividad_publicada(coursework: dict[str, Any]) -> bool:
    """
    Considera publicada cuando state es PUBLISHED.
    Si el campo no viene, asumimos True para no esconder actividades válidas.
    """
    state = coursework.get("state")
    if state is None:
        return True
    return state == "PUBLISHED"


def actividad_reciente(coursework: dict[str, Any], dias: int = DIAS_RECIENTES) -> bool:
    """
    Considera reciente una actividad creada o actualizada dentro
    de los últimos N días.
    """
    limite = ahora_utc() - timedelta(days=dias)

    creation_time = parse_google_datetime(coursework.get("creationTime"))
    update_time = parse_google_datetime(coursework.get("updateTime"))
    due_date = None

    # Algunas actividades solo traen dueDate, así que también lo intentamos.
    due = coursework.get("dueDate")
    if isinstance(due, dict):
        try:
            year = due.get("year")
            month = due.get("month")
            day = due.get("day")
            if year and month and day:
                due_date = datetime(year, month, day, tzinfo=timezone.utc)
        except ValueError:
            due_date = None

    fechas_validas = [f for f in [creation_time, update_time, due_date] if f is not None]
    if not fechas_validas:
        return False

    return any(f >= limite for f in fechas_validas)


def filtrar_actividades(
    classroom_service,
    course_id: str,
    courseworks: list[dict[str, Any]],
    filtro: str,
) -> list[dict[str, Any]]:
    """
    Aplica filtro a las actividades.
    """
    if filtro == "all":
        return courseworks

    if filtro == "published":
        return [cw for cw in courseworks if actividad_publicada(cw)]

    if filtro == "recent":
        return [cw for cw in courseworks if actividad_reciente(cw)]

    if filtro == "with_submissions":
        filtradas = []
        for cw in courseworks:
            try:
                submissions = obtener_todas_las_entregas(
                    classroom_service=classroom_service,
                    course_id=course_id,
                    coursework_id=cw["id"],
                )
                if submissions:
                    filtradas.append(cw)
            except HttpError as err:
                print(
                    f"⚠️ No se pudieron revisar entregas para actividad "
                    f"'{cw.get('title', cw.get('id', 'sin_titulo'))}': {err}"
                )
        return filtradas

    return courseworks


# ==========================================================
# Perfil del alumno
# ==========================================================

def extraer_datos_usuario_desde_historial(submission: dict[str, Any]) -> dict[str, str]:
    """
    Intenta recuperar nombre/correo desde submissionHistory.
    Esto sirve como plan B cuando no hay scopes suficientes
    para consultar userProfiles.
    """
    history = submission.get("submissionHistory", [])

    for event in history:
        actor = event.get("actorUser", {})
        if not actor:
            continue

        # Algunas respuestas incluyen profile o campos similares
        # según el tipo de evento o permisos.
        name = actor.get("name", {}) or {}
        given_name = name.get("givenName", "") or ""
        family_name = name.get("familyName", "") or ""
        full_name = name.get("fullName", "") or ""

        email = actor.get("emailAddress", "") or ""

        if not given_name and full_name:
            partes = full_name.split()
            if partes:
                given_name = partes[0]
                if len(partes) > 1:
                    family_name = " ".join(partes[1:])

        if given_name or family_name or email:
            return {
                "correo": email,
                "nombre": given_name,
                "apellido": family_name,
            }

    return {
        "correo": "",
        "nombre": "",
        "apellido": "",
    }


def obtener_perfil_usuario(
    classroom_service,
    user_id: str,
    profile_scope_disponible: bool,
) -> dict[str, str]:
    """
    Recupera correo, nombre y apellido del alumno.

    Estrategia:
    1. Si hay scope disponible, intenta userProfiles
    2. Si no hay scope, el caller puede usar fallback desde historial
    """
    if not profile_scope_disponible:
        return {
            "correo": "",
            "nombre": "",
            "apellido": "",
        }

    try:
        profile = classroom_service.userProfiles().get(userId=user_id).execute()

        email = profile.get("emailAddress", "") or ""

        name = profile.get("name", {}) or {}
        given_name = name.get("givenName", "") or ""
        family_name = name.get("familyName", "") or ""
        full_name = name.get("fullName", "") or ""

        if not given_name and full_name:
            partes = full_name.split()
            if partes:
                given_name = partes[0]
                if len(partes) > 1:
                    family_name = " ".join(partes[1:])

        return {
            "correo": email,
            "nombre": given_name,
            "apellido": family_name,
        }

    except HttpError as err:
        # Si el scope no alcanza, el programa no debe estar avisando esto
        # 27 veces por actividad. Lo manejamos en el caller.
        raise err


def detectar_scope_perfil(classroom_service) -> bool:
    """
    Prueba una sola vez si el token tiene permisos para consultar perfiles.
    Así evitamos un error 403 repetido para cada alumno.
    """
    try:
        classroom_service.userProfiles().get(userId="me").execute()
        return True
    except HttpError as err:
        status = getattr(err, "status_code", None)
        contenido = str(err)

        if status == 403 or "ACCESS_TOKEN_SCOPE_INSUFFICIENT" in contenido:
            return False

        # Si fue otro error, mejor propagamos.
        raise err


# ==========================================================
# Adjuntos
# ==========================================================

def obtener_adjuntos(submission: dict[str, Any]) -> list[dict[str, Any]]:
    assignment_submission = submission.get("assignmentSubmission", {})
    return assignment_submission.get("attachments", [])


def tiene_adjuntos(submission: dict[str, Any]) -> bool:
    return len(obtener_adjuntos(submission)) > 0




def obtener_due_date_texto(coursework: dict[str, Any]) -> str:
    """
    Convierte dueDate de Classroom a texto YYYY-MM-DD.
    """
    due_date = coursework.get("dueDate")
    if not isinstance(due_date, dict):
        return ""

    year = due_date.get("year")
    month = due_date.get("month")
    day = due_date.get("day")

    if not (year and month and day):
        return ""

    try:
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
    except (TypeError, ValueError):
        return ""


def obtener_due_time_texto(coursework: dict[str, Any]) -> str:
    """
    Convierte dueTime de Classroom a texto HH:MM:SS.
    """
    due_time = coursework.get("dueTime")
    if not isinstance(due_time, dict):
        return ""

    hours = due_time.get("hours", 0)
    minutes = due_time.get("minutes", 0)
    seconds = due_time.get("seconds", 0)

    try:
        return f"{int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}"
    except (TypeError, ValueError):
        return ""
def imprimir_resumen_entrega(submission: dict[str, Any]) -> None:
    """
    Imprime información útil para la revisión en terminal.
    """
    print(f"Entrega ID: {submission.get('id', 'Sin ID')}")
    print(f"  userId: {submission.get('userId', 'Sin userId')}")
    print(f"  estado: {submission.get('state', 'Sin estado')}")
    print(f"  late: {submission.get('late', False)}")
    print(f"  assignedGrade: {submission.get('assignedGrade')}")
    print(f"  draftGrade: {submission.get('draftGrade')}")
    print(f"  reentregada: {es_reentregada(submission)}")
    print(f"  no evaluada: {es_no_evaluada(submission)}")
    print(f"  attached: {tiene_adjuntos(submission)}")


def descargar_adjuntos_entrega(
    submission: dict[str, Any],
    drive_service,
    carpeta_entrega: str,
) -> int:
    """
    Descarga adjuntos de una entrega.
    Regresa cuántos archivos reales de Drive se bajaron.
    """
    attachments = obtener_adjuntos(submission)
    descargados = 0

    if not attachments:
        print("  adjuntos: ninguno")
        return descargados

    print("  adjuntos:")

    for att in attachments:
        if "driveFile" in att:
            drive_file = att.get("driveFile", {})
            drive_meta = drive_file.get("driveFile") or drive_file

            file_id = drive_meta.get("id")
            title = drive_meta.get("title", "archivo")

            print(f"    - DriveFile: {title} | id={file_id}")

            if file_id:
                ruta = download_file(
                    drive_service=drive_service,
                    file_id=file_id,
                    file_name=title,
                    folder=carpeta_entrega,
                )
                if ruta:
                    descargados += 1

        elif "link" in att:
            link = att["link"]
            print(
                f"    - Link: {link.get('title', 'Sin título')} "
                f"| url={link.get('url', 'Sin URL')}"
            )

        elif "form" in att:
            form = att["form"]
            print(
                f"    - Form: {form.get('title', 'Sin título')} "
                f"| url={form.get('formUrl', 'Sin URL')}"
            )

        elif "youTubeVideo" in att:
            video = att["youTubeVideo"]
            print(
                f"    - YouTube: {video.get('title', 'Sin título')} "
                f"| url={video.get('alternateLink', 'Sin URL')}"
            )

        else:
            print("    - Tipo de adjunto no manejado directamente.")

    return descargados


# ==========================================================
# CSV y ZIP
# ==========================================================

def escribir_csv_resumen(csv_path: str, filas: list[dict[str, str]]) -> None:
    """
    Genera CSV general con información de curso, actividad y alumno.
    """
    asegurar_directorio(os.path.dirname(csv_path))

    with open(csv_path, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(
            csvfile,
            fieldnames=[
                "curso",
                "actividad",
                "due_date",
                "due_time",
                "correo",
                "nombre",
                "apellido",
                "attached",
            ],
        )
        writer.writeheader()
        writer.writerows(filas)

    print(f"\n✅ CSV generado: {csv_path}")


def comprimir_carpeta_a_zip(carpeta_origen: str, zip_sin_extension: str) -> str:
    """
    Comprime toda la carpeta en un zip.
    """
    zip_path = shutil.make_archive(zip_sin_extension, "zip", carpeta_origen)
    print(f"✅ ZIP generado: {zip_path}")
    return zip_path


# ==========================================================
# Procesamiento de una actividad
# ==========================================================

def procesar_actividad(
    classroom_service,
    drive_service,
    course: dict[str, Any],
    coursework: dict[str, Any],
    modo_descarga: str,
    carpeta_base: str,
    perfiles_cache: dict[str, dict[str, str]],
    filas_csv: list[dict[str, str]],
    estadisticas: dict[str, int],
    profile_scope_disponible: bool,
) -> None:
    """
    Procesa una actividad completa:
    - recupera entregas
    - aplica filtro
    - descarga adjuntos
    - agrega filas al CSV
    - actualiza estadísticas
    """
    course_id = course["id"]
    course_name = course.get("name", f"curso_{course_id}")

    coursework_id = coursework["id"]
    coursework_title = coursework.get("title", f"actividad_{coursework_id}")
    coursework_display = coursework.get("display_name", coursework_title)

    print("\n" + "=" * 90)
    print(f"Procesando actividad: {coursework_display}")
    print("=" * 90)

    submissions = obtener_todas_las_entregas(
        classroom_service=classroom_service,
        course_id=course_id,
        coursework_id=coursework_id,
    )

    estadisticas["actividades_procesadas"] += 1
    estadisticas["entregas_totales"] += len(submissions)

    if not submissions:
        print("No se encontraron entregas para esta actividad.")
        return

    entregas_filtradas = filtrar_entregas(submissions, modo_descarga)
    estadisticas["entregas_filtradas"] += len(entregas_filtradas)

    print(f"Entregas totales encontradas: {len(submissions)}")
    print(f"Entregas a procesar con este filtro: {len(entregas_filtradas)}")

    if not entregas_filtradas:
        print("No hay entregas que coincidan con el filtro para esta actividad.")
        return

    carpeta_actividad = os.path.join(
        carpeta_base,
        f"{limpiar_nombre_archivo(coursework_title)}_{coursework_id}",
    )
    asegurar_directorio(carpeta_actividad)

    for submission in entregas_filtradas:
        imprimir_resumen_entrega(submission)

        user_id = submission.get("userId", "sin_userId")

        if user_id not in perfiles_cache:
            perfil = {
                "correo": "",
                "nombre": "",
                "apellido": "",
            }

            # Primero intentamos el perfil formal si el scope existe
            if profile_scope_disponible:
                try:
                    perfil = obtener_perfil_usuario(
                        classroom_service=classroom_service,
                        user_id=user_id,
                        profile_scope_disponible=True,
                    )
                except HttpError as err:
                    # Si por alguna razón volvió a fallar, degradamos con fallback
                    print(
                        f"  ⚠️ No se pudo leer userProfiles para userId={user_id}. "
                        f"Se usará fallback. Detalle: {err}"
                    )
                    perfil = extraer_datos_usuario_desde_historial(submission)

            else:
                # Fallback silencioso si sabemos de antemano que no hay scope
                perfil = extraer_datos_usuario_desde_historial(submission)

            perfiles_cache[user_id] = perfil

        perfil = perfiles_cache[user_id]

        nombre_carpeta_entrega = construir_nombre_carpeta_entrega(
            submission=submission,
            perfil=perfil,
        )

        carpeta_entrega = os.path.join(carpeta_actividad, nombre_carpeta_entrega)
        asegurar_directorio(carpeta_entrega)

        archivos_descargados = descargar_adjuntos_entrega(
            submission=submission,
            drive_service=drive_service,
            carpeta_entrega=carpeta_entrega,
        )

        if archivos_descargados > 0:
            estadisticas["archivos_descargados"] += archivos_descargados

        filas_csv.append(
            {
                "curso": course_name,
                "actividad": coursework_title,
                "due_date": obtener_due_date_texto(coursework),
                "due_time": obtener_due_time_texto(coursework),
                "correo": perfil.get("correo", ""),
                "nombre": perfil.get("nombre", ""),
                "apellido": perfil.get("apellido", ""),
                "attached": str(tiene_adjuntos(submission)).lower(),
            }
        )

        print("-" * 70)


# ==========================================================
# Flujo principal
# ==========================================================

def main() -> None:
    """
    Flujo principal:
    1. autentica
    2. elige curso
    3. elige alcance
    4. filtra actividades
    5. elige filtro de entregas
    6. elige formato de salida
    7. descarga
    8. genera CSV y zip
    """
    settings = get_settings()
    ensure_directories(settings)

    creds = get_credentials(
        credentials_path=settings.credentials_path,
        token_path=settings.token_path,
    )

    print("Autenticación correcta.")
    print(f"Token válido: {creds.valid}")

    try:
        classroom_service = build("classroom", "v1", credentials=creds)
        drive_service = build("drive", "v3", credentials=creds)

        # Probamos una sola vez si el token puede leer perfiles
        profile_scope_disponible = detectar_scope_perfil(classroom_service)
        if profile_scope_disponible:
            print("✅ Scope de perfiles disponible.")
        else:
            print(
                "⚠️ El token no tiene scope para leer perfiles de alumnos. "
                "Se continuará con fallback y el CSV puede traer nombre/correo vacíos."
            )

        # ==========================================================
        # 1) Curso
        # ==========================================================
        courses = obtener_todos_los_cursos(classroom_service)

        if not courses:
            print("No se encontraron cursos.")
            return

        selected_course = seleccionar_opcion(courses, "un curso")
        course_id = selected_course["id"]
        course_name = limpiar_nombre_archivo(selected_course.get("name", f"curso_{course_id}"))
        course_display = selected_course.get("display_name", selected_course.get("name", "Curso"))

        print(f"\n✅ Curso seleccionado: {course_display}")

        # ==========================================================
        # 2) Alcance
        # ==========================================================
        alcance_descarga = seleccionar_alcance_descarga()
        print(f"\n📚 Alcance seleccionado: {alcance_descarga}")

        # ==========================================================
        # 3) Filtro de actividades
        # ==========================================================
        filtro_actividades = seleccionar_filtro_actividades()
        print(f"🧩 Filtro de actividades: {filtro_actividades}")

        # ==========================================================
        # 4) Filtro de entregas
        # ==========================================================
        modo_descarga = seleccionar_modo_descarga()
        print(f"\n📥 Modo seleccionado: {describir_modo_descarga(modo_descarga)}")

        # ==========================================================
        # 5) Formato de salida
        # ==========================================================
        formato_salida = seleccionar_formato_salida()
        print(f"📦 Formato de salida: {formato_salida}")

        # ==========================================================
        # 6) Actividades
        # ==========================================================
        courseworks = obtener_todas_las_actividades(classroom_service, course_id)

        if not courseworks:
            print("No se encontraron actividades en este curso.")
            return

        courseworks_filtradas = filtrar_actividades(
            classroom_service=classroom_service,
            course_id=course_id,
            courseworks=courseworks,
            filtro=filtro_actividades,
        )

        if not courseworks_filtradas:
            print("No quedaron actividades después de aplicar el filtro.")
            return

        print(
            f"✅ Actividades encontradas: {len(courseworks)} | "
            f"después del filtro: {len(courseworks_filtradas)}"
        )

        # ==========================================================
        # 7) Carpeta base
        # ==========================================================
        carpeta_base = os.path.join(
            "downloads",
            f"{course_name}_{course_id}",
        )
        asegurar_directorio(carpeta_base)

        perfiles_cache: dict[str, dict[str, str]] = {}
        filas_csv: list[dict[str, str]] = []
        estadisticas = {
            "actividades_procesadas": 0,
            "entregas_totales": 0,
            "entregas_filtradas": 0,
            "archivos_descargados": 0,
        }

        # ==========================================================
        # 8) Procesamiento
        # ==========================================================
        if alcance_descarga == "single_coursework":
            selected_coursework = seleccionar_opcion(courseworks_filtradas, "una actividad")
            print(f"\n✅ Actividad seleccionada: {selected_coursework['display_name']}")

            procesar_actividad(
                classroom_service=classroom_service,
                drive_service=drive_service,
                course=selected_course,
                coursework=selected_coursework,
                modo_descarga=modo_descarga,
                carpeta_base=carpeta_base,
                perfiles_cache=perfiles_cache,
                filas_csv=filas_csv,
                estadisticas=estadisticas,
                profile_scope_disponible=profile_scope_disponible,
            )

            nombre_csv = "resumen_entregas.csv"
            nombre_zip = f"{course_name}_{course_id}"

        elif alcance_descarga == "all_courseworks":
            print(
                f"\n✅ Se procesarán todas las actividades filtradas del curso: "
                f"{len(courseworks_filtradas)}"
            )

            for idx, coursework in enumerate(courseworks_filtradas, start=1):
                print(f"\n[{idx}/{len(courseworks_filtradas)}] {coursework['display_name']}")

                procesar_actividad(
                    classroom_service=classroom_service,
                    drive_service=drive_service,
                    course=selected_course,
                    coursework=coursework,
                    modo_descarga=modo_descarga,
                    carpeta_base=carpeta_base,
                    perfiles_cache=perfiles_cache,
                    filas_csv=filas_csv,
                    estadisticas=estadisticas,
                    profile_scope_disponible=profile_scope_disponible,
                )

            nombre_csv = "resumen_todas_las_actividades.csv"
            nombre_zip = f"{course_name}_{course_id}_todas_las_actividades"

        else:
            print("❌ Alcance de descarga no reconocido.")
            return

        # ==========================================================
        # 9) CSV
        # ==========================================================
        csv_path = os.path.join(carpeta_base, nombre_csv)
        escribir_csv_resumen(csv_path, filas_csv)

        # ==========================================================
        # 10) ZIP
        # ==========================================================
        if formato_salida == "zip_and_folder":
            zip_base_name = os.path.join("downloads", nombre_zip)
            comprimir_carpeta_a_zip(carpeta_base, zip_base_name)

        # ==========================================================
        # 11) Resumen final
        # ==========================================================
        print("\n" + "=" * 90)
        print("RESUMEN FINAL")
        print("=" * 90)
        print(f"Curso: {course_display}")
        print(f"Actividades procesadas: {estadisticas['actividades_procesadas']}")
        print(f"Entregas totales vistas: {estadisticas['entregas_totales']}")
        print(f"Entregas que cumplieron filtro: {estadisticas['entregas_filtradas']}")
        print(f"Archivos descargados: {estadisticas['archivos_descargados']}")
        print(f"Filas en CSV: {len(filas_csv)}")
        print(f"Carpeta base: {carpeta_base}")
        if formato_salida == "zip_and_folder":
            print(f"ZIP: downloads/{nombre_zip}.zip")

        print("\n✅ Proceso terminado.")

    except HttpError as err:
        print(f"Error al consultar Classroom: {err}")


if __name__ == "__main__":
    main()