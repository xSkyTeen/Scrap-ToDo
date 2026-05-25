import os
import re
import msal
import requests
from datetime import datetime

# =========================================================
# CONFIG MICROSOFT GRAPH
# =========================================================
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
        refresh_token=MS_REFRESH_TOKEN.strip(),
        scopes=SCOPES
    )

    if result and "error" in result:
        print("❌ Error de Microsoft Auth con el Refresh Token proporcionado:")
        print(f"   Error: {result.get('error')}")
        print(f"   Descripción: {result.get('error_description')}")
        result = None

# 2. Si no estamos en la nube (o falló el refresh token), usamos el flujo local
if not result:
    print("💻 Modo Local: Buscando caché o inicio interactivo...")
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

# Validar Access Token final
if result and "access_token" in result:
    access_token = result["access_token"]
    headers_graph = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    print("✅ Login Microsoft Graph Exitoso")
else:
    print("🚨 FATAL: No se pudo obtener el Access Token de ninguna forma.")
    exit(1)

# =========================================================
# COOKIES / LOGIN AULA VIRTUAL (Adaptado Local y Servidor)
# =========================================================
session = requests.Session()

UNAP_USER = os.getenv("UNAP_USUARIO")
UNAP_PASS = os.getenv("UNAP_PASSWORD")

if os.getenv("GITHUB_ACTIONS") and UNAP_USER and UNAP_PASS:
    print("🔄 Modo Nube: Intentando login automático en el Aula Virtual UNAP...")
    url_login = "https://aulavirtual2.unap.edu.pe/api/auth/login"
    payload = {
        "username": UNAP_USER.strip(),
        "password": UNAP_PASS.strip()
    }

    response_login = session.post(url_login, json=payload)

    if response_login.status_code == 200:
        print("✅ Login en el Aula Virtual UNAP Exitoso (Sesión iniciada en la nube)")
    else:
        print(f"🚨 ERROR FATAL: Credenciales rechazadas por la UNAP. Status: {response_login.status_code}")
        exit(1)
else:
    print("💻 Modo Local: Cargando cookies desde Zen Browser...")
    import browser_cookie3

    try:
        cookies = browser_cookie3.firefox(
            cookie_file='/home/sky/.config/zen/fqk3pjrk.Default (release)/cookies.sqlite'
        )
        session.cookies.update(cookies)
        print("✅ Cookies Aula Virtual cargadas desde Zen Browser (Local)")
    except Exception as e:
        print(f"⚠️ No se pudieron cargar las cookies locales de Zen: {e}")


# =========================================================
# UTILS / LIMPIEZA
# =========================================================
def limpiar_html(html_text):
    """Elimina etiquetas HTML como <p>, </p>, etc., para dejar texto plano."""
    if not html_text:
        return ""
    texto = re.sub(r'</p>|<br\s*/?>', '\n', html_text)
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

try:
    cursos = response.json()
    print(f"✅ Cursos encontrados: {len(cursos)}")
except requests.exceptions.JSONDecodeError:
    print("🚨 ERROR CRÍTICO: El Aula Virtual de la UNAP no devolvió un JSON válido.")
    print(f"   Código de estado HTTP del servidor: {response.status_code}")
    print(f"   Fragmento recibido (Primeros 300 caracteres):\n{response.text[:300]}")
    exit(1)

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

    if nombre in mapa_listas:
        LIST_ID = mapa_listas[nombre]
        print(f"\n📚 Curso: {nombre} (Lista encontrada)")
    else:
        response_create = requests.post(url_lists, headers=headers_graph, json={"displayName": nombre})
        LIST_ID = response_create.json()["id"]
        mapa_listas[nombre] = LIST_ID
        print(f"\n📚 Curso: {nombre} (Lista creada de cero)")

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

        detalle_tarea = limpiar_html(tarea.get("description", "Sin descripción detallada."))

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