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
app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
result = None
MS_REFRESH_TOKEN = os.getenv("MICROSOFT_REFRESH_TOKEN")

if MS_REFRESH_TOKEN:
    print("🔄 Intentando login en la nube usando Refresh Token...")
    result = app.acquire_token_by_refresh_token(
        refresh_token=MS_REFRESH_TOKEN.strip(),
        scopes=SCOPES
    )
    if result and "error" in result:
        print("❌ Error de Microsoft Auth con el Refresh Token:")
        print(f"   Error: {result.get('error')}")
        result = None

if not result:
    print("💻 Modo Local: Buscando caché o inicio interactivo...")
    cache = msal.SerializableTokenCache()
    if os.path.exists("token_cache.bin"):
        with open("token_cache.bin", "r") as f:
            cache.deserialize(f.read())

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result:
        print("⚠️ Requiere login interactivo local.")
        result = app.acquire_token_interactive(scopes=SCOPES)
    if cache.has_state_changed:
        with open("token_cache.bin", "w") as f:
            f.write(cache.serialize())

if result and "access_token" in result:
    access_token = result["access_token"]
    headers_graph = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    print("✅ Login Microsoft Graph Exitoso")
else:
    print("🚨 FATAL: No se pudo obtener el Access Token.")
    exit(1)

# =========================================================
# COOKIES AULA VIRTUAL
# =========================================================
session = requests.Session()
AV_COOKIE = os.getenv("AULA_VIRTUAL_COOKIE")

if AV_COOKIE and AV_COOKIE.strip():
    session.headers.update({"Cookie": AV_COOKIE.strip()})
    print("✅ Cookies Aula Virtual cargadas desde Variables de Entorno (GitHub)")
else:
    if os.getenv("GITHUB_ACTIONS"):
        print("🚨 ERROR FATAL: No se encontró la variable 'AULA_VIRTUAL_COOKIE' en los Secrets.")
        exit(1)

    print("💻 Modo Local: Cargando cookies desde Zen Browser...")
    import browser_cookie3

    try:
        cookies = browser_cookie3.firefox(
            cookie_file='/home/sky/.config/zen/fqk3pjrk.Default (release)/cookies.sqlite'
        )
        session.cookies.update(cookies)
        print("✅ Cookies Aula Virtual cargadas desde Zen Browser (Local)")
    except Exception as e:
        print(f"⚠️ Error cargando cookies locales: {e}")


# =========================================================
# UTILS / LIMPIEZA
# =========================================================
def limpiar_html(html_text):
    if not html_text: return ""
    texto = re.sub(r'</p>|<br\s*/?>', '\n', html_text)
    texto = re.sub(r'<[^>]+>', '', texto)
    return texto.strip()


def formatear_fecha(fecha_str):
    try:
        fecha = datetime.strptime(fecha_str, "%d/%m/%Y %H:%M")
        return fecha.strftime("%Y-%m-%dT%H:%M:00")
    except:
        return None


# =========================================================
# OBTENER CURSOS (Con captura de errores de Login/Redirección)
# =========================================================
url_cursos = "https://aulavirtual2.unap.edu.pe/web/user/info/system/courseinrole"
response = session.get(url_cursos)

try:
    cursos = response.json()
    print(f"✅ Cursos encontrados: {len(cursos)}")
except requests.exceptions.JSONDecodeError:
    print("🚨 ERROR CRÍTICO: La UNAP nos redirigió al login.")
    print("💡 Diagnóstico: La 'AULA_VIRTUAL_COOKIE' que pusiste en GitHub expiró o está incompleta.")
    exit(1)

# =========================================================
# RECORRER CURSOS Y SUBIR A TO-DO
# =========================================================
url_lists = "https://graph.microsoft.com/v1.0/me/todo/lists"
response_lists = requests.get(url_lists, headers=headers_graph)
listas_existentes = response_lists.json().get("value", [])
mapa_listas = {lista["displayName"]: lista["id"] for lista in listas_existentes}

for curso in cursos:
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

    # PROCESAR FOROS
    url_foros = f"https://aulavirtual2.unap.edu.pe/web/forum/list?s={section_id}"
    response_foros = session.get(url_foros)
    foros = response_foros.json() if response_foros.status_code == 200 else []

    for foro in foros:
        titulo = f"📢 {foro['name']}"
        if titulo in existing_titles: continue

        descripcion_cuerpo = f"📖 Curso: {nombre}\n📌 Unidad: {foro['unidadName']}\n⏰ Límite: {foro['dateEndView']}"
        due_date = formatear_fecha(foro["dateEndView"])
        body = {"title": titulo, "body": {"content": descripcion_cuerpo, "contentType": "text"}}
        if due_date:
            body["dueDateTime"] = {"dateTime": due_date, "timeZone": "America/Lima"}
        requests.post(url_tasks, headers=headers_graph, json=body)
        print(f"   🔹 Foro agregado: {titulo}")

    # PROCESAR TAREAS
    url_tareas = f"https://aulavirtual2.unap.edu.pe/web/homework/list?s={section_id}"
    response_tareas = session.get(url_tareas)
    tareas = response_tareas.json() if response_tareas.status_code == 200 else []

    for tarea in tareas:
        titulo = f"📝 {tarea['title']}"
        if titulo in existing_titles: continue

        detalle_tarea = limpiar_html(tarea.get("description", "Sin descripción."))
        descripcion_cuerpo = f"📖 Curso: {nombre}\n📌 Unidad: {tarea['unidad']}\n⏰ Límite: {tarea['dateEnd']}\n---\n{detalle_tarea}"
        due_date = formatear_fecha(tarea["dateEnd"])
        body = {"title": titulo, "body": {"content": descripcion_cuerpo, "contentType": "text"}}
        if due_date:
            body["dueDateTime"] = {"dateTime": due_date, "timeZone": "America/Lima"}
        requests.post(url_tasks, headers=headers_graph, json=body)
        print(f"   🔹 Tarea agregado: {titulo}")

print("\n🔥 SINCRONIZACION FINALIZADA")