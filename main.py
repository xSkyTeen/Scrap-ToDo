import os
import re
import msal
import requests
from datetime import datetime

# =========================================================
# CONFIG MICROSOFT GRAPH
# =========================================================
# Busca el ID en GitHub Actions; si no existe (local), usa tu ID por defecto
CLIENT_ID = os.getenv("CLIENT_ID", "f7b10369-96d3-4a79-a7b6-1ebac4232def")

AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Tasks.ReadWrite", "User.Read"]
# =========================================================
# LOGIN MICROSOFT
# =========================================================
app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

result = None

# 1. Intentar levantar mediante el Refresh Token de GitHub (Modo Producción 24/7)
MS_REFRESH_TOKEN = os.getenv("MICROSOFT_REFRESH_TOKEN")

if MS_REFRESH_TOKEN:
    print("🔄 Intentando login en la nube usando Refresh Token...")
    result = app.acquire_token_by_refresh_token(
        refresh_token=MS_REFRESH_TOKEN,
        scopes=SCOPES
    )

# 2. Si no estamos en la nube, usamos el flujo local normal con la caché local
if not result:
    cache = msal.SerializableTokenCache()
    if os.path.exists("token_cache.bin"):
        with open("token_cache.bin", "r") as f:
            cache.deserialize(f.read())

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        print("⚠️ Requiere login interactivo local.")
        result = app.acquire_token_interactive(scopes=SCOPES)

    if cache.has_state_changed:
        with open("token_cache.bin", "w") as f:
            f.write(cache.serialize())

access_token = result["access_token"]
headers_graph = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}
print("✅ Login Microsoft Graph Exitoso")
# =========================================================
# COOKIES AULA VIRTUAL (Adaptado para Local y Servidor)
# =========================================================
session = requests.Session()

# Si estás en GitHub Actions, lee la cookie desde los Secrets para evitar usar browser_cookie3
AV_COOKIE = os.getenv("AULA_VIRTUAL_COOKIE")

if AV_COOKIE:
    session.headers.update({"Cookie": AV_COOKIE})
    print("✅ Cookies Aula Virtual cargadas desde Variables de Entorno (GitHub)")
else:
    import browser_cookie3

    cookies = browser_cookie3.firefox(
        cookie_file='/home/sky/.config/zen/fqk3pjrk.Default (release)/cookies.sqlite'
    )
    session.cookies.update(cookies)
    print("✅ Cookies Aula Virtual cargadas desde Zen Browser (Local)")


# =========================================================
# UTILS / LIMPIEZA
# =========================================================
def limpiar_html(html_text):
    """Elimina etiquetas HTML como <p>, </p>, etc., para dejar texto plano."""
    if not html_text:
        return ""
    # Reemplaza <p> cerrados o <br> por saltos de línea para mantener orden
    texto = re.sub(r'</p>|<br\s*/?>', '\n', html_text)
    # Remueve cualquier otra etiqueta HTML restante
    texto = re.sub(r'<[^>]+>', '', texto)
    return texto.strip()


def formatear_fecha(fecha_str):
    """Convierte 'DD/MM/YYYY HH:MM' al formato ISO requerido por Microsoft To Do."""
    try:
        fecha = datetime.strptime(fecha_str, "%d/%m/%Y %H:%M")
        return fecha.strftime("%Y-%m-%dT%H:%M:00")
    except:
        return None


# =========================================================
# OBTENER CURSOS
# =========================================================
url_cursos = "https://aulavirtual2.unap.edu.pe/web/user/info/system/courseinrole"
response = session.get(url_cursos)
cursos = response.json()
print(f"✅ Cursos encontrados: {len(cursos)}")

# =========================================================
# OBTENER LISTAS EXISTENTES
# =========================================================
url_lists = "https://graph.microsoft.com/v1.0/me/todo/lists"
response_lists = requests.get(url_lists, headers=headers_graph)
listas_existentes = response_lists.json().get("value", [])

mapa_listas = {lista["displayName"]: lista["id"] for lista in listas_existentes}

# =========================================================
# RECORRER CURSOS
# =========================================================
for curso in cursos:
    codigo = curso["codeCourse"]
    nombre = curso["name"].strip()
    section_id = curso["sectionId"]

    # Crear o encontrar la lista del curso
    if nombre in mapa_listas:
        LIST_ID = mapa_listas[nombre]
        print(f"\n📚 Curso: {nombre} (Lista encontrada)")
    else:
        response_create = requests.post(url_lists, headers=headers_graph, json={"displayName": nombre})
        LIST_ID = response_create.json()["id"]
        mapa_listas[nombre] = LIST_ID
        print(f"\n📚 Curso: {nombre} (Lista creada de cero)")

    # Obtener tareas ya existentes en la lista para evitar duplicados
    existing_titles = set()
    url_tasks = f"https://graph.microsoft.com/v1.0/me/todo/lists/{LIST_ID}/tasks"
    response_tasks = requests.get(url_tasks, headers=headers_graph)

    if response_tasks.status_code == 200:
        for tarea in response_tasks.json().get("value", []):
            existing_titles.add(tarea["title"])

    # =====================================================
    # PROCESAR FOROS
    # =====================================================
    url_foros = f"https://aulavirtual2.unap.edu.pe/web/forum/list?s={section_id}"
    response_foros = session.get(url_foros)
    foros = response_foros.json() if response_foros.status_code == 200 else []

    for foro in foros:
        titulo = f"📢 {foro['name']}"
        if titulo in existing_titles:
            continue

        # Observación formateada para la descripción de la tarea
        descripcion_cuerpo = (
            f"📖 Curso: {nombre}\n"
            f"📌 Unidad: {foro['unidadName']}\n"
            f"💬 Respuestas actuales: {foro['answers']}\n"
            f"⏰ Fecha Límite: {foro['dateEndView']}"
        )

        due_date = formatear_fecha(foro["dateEndView"])

        body = {
            "title": titulo,
            "body": {
                "content": descripcion_cuerpo,
                "contentType": "text"
            }
        }

        if due_date:
            body["dueDateTime"] = {
                "dateTime": due_date,
                "timeZone": "America/Lima"
            }

        requests.post(url_tasks, headers=headers_graph, json=body)
        print(f"   🔹 Foro agregado: {titulo}")

    # =====================================================
    # PROCESAR TAREAS
    # =====================================================
    url_tareas = f"https://aulavirtual2.unap.edu.pe/web/homework/list?s={section_id}"
    response_tareas = session.get(url_tareas)
    tareas = response_tareas.json() if response_tareas.status_code == 200 else []

    for tarea in tareas:
        titulo = f"📝 {tarea['title']}"
        if titulo in existing_titles:
            continue

        # Extraemos y limpiamos la descripción que viene en el JSON de la UNAP
        detalle_tarea = limpiar_html(tarea.get("description", "Sin descripción detallada."))

        # Estructuramos el campo de observación/cuerpo en To-Do
        descripcion_cuerpo = (
            f"📖 Curso: {nombre}\n"
            f"📌 Unidad: {tarea['unidad']}\n"
            f"🔄 Estado Aula: {tarea['state']}\n"
            f"⏰ Fecha Límite: {tarea['dateEnd']}\n"
            f"----------------------------------------\n"
            f"📝 DETALLE:\n{detalle_tarea}"
        )

        due_date = formatear_fecha(tarea["dateEnd"])

        body = {
            "title": titulo,
            "body": {
                "content": descripcion_cuerpo,
                "contentType": "text"
            }
        }

        if due_date:
            body["dueDateTime"] = {
                "dateTime": due_date,
                "timeZone": "America/Lima"
            }

        requests.post(url_tasks, headers=headers_graph, json=body)
        print(f"   🔹 Tarea agregada: {titulo}")

print("\n🔥 SINCRONIZACION FINALIZADA")