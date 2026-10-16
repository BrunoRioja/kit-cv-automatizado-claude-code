"""Conectores a los portales de empleo propios de las empresas objetivo (ejemplo: grandes empresas de Perú).

Cada portal usa el servicio público de su sistema de reclutamiento (el mismo que llama
su web, sin login). Se piden solo las vacantes de Perú, se filtran por título antes de
abrir cada aviso y se espera PAUSA_SEG entre consultas.

Todos devuelven vacantes con las mismas claves: id, fuente, empresa, titulo, ubicacion,
modalidad, jornada, fecha_publicacion (date), link, detalle.

Empresas revisadas en oct 2026 y por qué no tienen conector:
- BCP, Grupo Gloria: su "trabaja con nosotros" lleva a su perfil de Bumeran (ya cubierto).
- Pacífico Seguros: solo LinkedIn (no se consulta por scripts).
- Backus: su portal HiringRoom no publica vacantes (solo "subir CV a base general").
- Rímac, Prima AFP, Entel, Intercorp Retail, Izipay: portal no identificado.
"""

import html
import re
import time
import unicodedata
from datetime import date, datetime

import requests

PAUSA_SEG = 1.0
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/128.0 Safari/537.36")

# Un título pasa a revisión si contiene alguna de estas palabras (sin tildes).
# Los portales mezclan español e inglés, por eso la lista es más amplia que en Bumeran.
TITULO_INTERES = [
    "analista", "analyst", "associate", "asociado", "data", "datos", "bi",
    "business intelligence", "inteligencia", "analytics", "analitica", "analitico",
    "insights", "reporting", "reportes", "comercial", "commercial", "planeamiento",
    "planning", "producto", "product", "proyectos", "projects", "automatizacion",
    "automation", "ia", "ai", "scientist", "cientifico", "key account", "cuentas clave",
    "especialista", "specialist", "consultant", "consultor", "strategy", "estrategia",
    "pricing", "engineer", "owner", "performance", "growth",
]

# Portales. `alias`: cómo aparece la empresa en Bumeran (texto contenido en clave_empresa),
# para reconocer la misma vacante publicada en ambos lados.
PORTALES = [
    {"clave": "BBVA", "empresa": "BBVA Perú", "tipo": "workday", "alias": ["bbva"],
     "host": "https://bbva.wd3.myworkdayjobs.com", "tenant": "bbva", "sitio": "BBVA",
     "facetas": {"locationCountry": ["0eb156ca580c4db786a7894bdaa77450"]}},  # Perú
    {"clave": "IBK", "empresa": "Interbank", "tipo": "oracle",
     "alias": ["interbank", "banco internacional"],
     "host": "https://epvw.fa.us6.oraclecloud.com", "sitio": "CX_1"},
    {"clave": "SCOTIA", "empresa": "Scotiabank Perú", "tipo": "successfactors",
     "alias": ["scotiabank"], "host": "https://jobs.scotiabank.com",
     "params": {"optionsFacetsDD_country": "PE"}},
    {"clave": "TEF", "empresa": "Telefónica del Perú (Movistar)", "tipo": "successfactors",
     "alias": ["telefonica", "movistar"], "host": "https://jobs.telefonica.com",
     "params": {"optionsFacetsDD_country": "PE"}},
    {"clave": "ARCA", "empresa": "Arca Continental Lindley", "tipo": "successfactors",
     "alias": ["arca continental", "lindley"], "host": "https://careers.arcacontal.com",
     "params": {"locationsearch": "Peru"}},
    {"clave": "NESTLE", "empresa": "Nestlé Perú", "tipo": "successfactors",
     "alias": ["nestle"], "host": "https://jobdetails.nestle.com",
     "params": {"optionsFacetsDD_country": "PE"}},
    # Alicorp es casi todo Perú; el país se valida al abrir cada aviso.
    {"clave": "ALICORP", "empresa": "Alicorp", "tipo": "successfactors",
     "alias": ["alicorp"], "host": "https://oportunidadesalicorp.com", "params": {}},
    {"clave": "MIBANCO", "empresa": "Mibanco", "tipo": "genoma", "alias": ["mibanco"],
     "slug": "mibanco-pe"},
    {"clave": "BELCORP", "empresa": "Belcorp", "tipo": "rankmi", "alias": ["belcorp"],
     "host": "https://ats.rankmi.com", "tenant": "c00744d26678d6df", "organizacion": "belcorp-connect"},
]


def _sesion():
    s = requests.Session()
    s.headers["User-Agent"] = UA
    return s


def _texto(h):
    """HTML a texto plano (sin imágenes embebidas)."""
    h = re.sub(r"<img[^>]*>", " ", h or "")
    h = re.sub(r"<(br|/p|/li|/div|/h\d)[^>]*>", "\n", h, flags=re.I)
    t = html.unescape(re.sub(r"<[^>]+>", " ", h))
    t = re.sub(r"[ \t\xa0]+", " ", t)
    return re.sub(r"\n\s*\n+", "\n", t).strip()


def _get(s, url, **kw):
    try:
        r = s.get(url, timeout=30, **kw)
        r.raise_for_status()
        return r
    finally:
        time.sleep(PAUSA_SEG)


# ---------------------------------------------------------------------------
# Workday (BBVA)
# ---------------------------------------------------------------------------
def _workday(p, interesa, desde=None):
    s = _sesion()
    base = f"{p['host']}/wday/cxs/{p['tenant']}/{p['sitio']}"
    candidatos, offset = [], 0
    while True:
        r = s.post(f"{base}/jobs", json={"appliedFacets": p["facetas"], "limit": 20,
                                          "offset": offset, "searchText": ""}, timeout=30)
        time.sleep(PAUSA_SEG)
        r.raise_for_status()
        avisos = r.json().get("jobPostings", [])
        candidatos += [a for a in avisos if interesa(a.get("title"))]
        offset += 20
        if len(avisos) < 20:
            break
    for a in candidatos:
        ji = _get(s, base + a["externalPath"]).json()["jobPostingInfo"]
        yield {
            "id": f"{p['clave']}-{ji['jobReqId']}",
            "titulo": ji["title"],
            "ubicacion": ji.get("location"),
            "modalidad": a.get("remoteType") or "",
            "jornada": ji.get("timeType"),
            "fecha_publicacion": date.fromisoformat(ji["startDate"]),
            "link": ji.get("externalUrl"),
            "detalle": _texto(ji.get("jobDescription")),
        }


# ---------------------------------------------------------------------------
# Oracle Recruiting Cloud (Interbank)
# ---------------------------------------------------------------------------
def _oracle(p, interesa, desde=None):
    s = _sesion()
    s.headers["Accept"] = "application/json"
    api = f"{p['host']}/hcmRestApi/resources/latest"
    candidatos, offset = [], 0
    while True:
        finder = f"findReqs;siteNumber={p['sitio']},limit=25,offset={offset},sortBy=POSTING_DATES_DESC"
        r = _get(s, f"{api}/recruitingCEJobRequisitions",
                 params={"onlyData": "true", "expand": "requisitionList", "finder": finder})
        item = r.json()["items"][0]
        lista = item.get("requisitionList", [])
        candidatos += [a for a in lista if interesa(a.get("Title"))]
        offset += 25
        if len(lista) < 25 or offset >= (item.get("TotalJobsCount") or 0):
            break
    for a in candidatos:
        r = _get(s, f"{api}/recruitingCEJobRequisitionDetails",
                 params={"expand": "all", "onlyData": "true",
                         "finder": f'ById;Id="{a["Id"]}",siteNumber={p["sitio"]}'})
        d = r.json()["items"][0]
        partes = [d.get(k) for k in ("ExternalDescriptionStr", "ExternalResponsibilitiesStr",
                                     "ExternalQualificationsStr")]
        yield {
            "id": f"{p['clave']}-{a['Id']}",
            "titulo": d["Title"],
            "ubicacion": d.get("PrimaryLocation"),
            "modalidad": d.get("WorkplaceType") or a.get("WorkplaceType") or "",
            "jornada": d.get("JobSchedule"),
            "fecha_publicacion": date.fromisoformat(d["ExternalPostedStartDate"][:10]),
            "link": f"{p['host']}/hcmUI/CandidateExperience/es/sites/{p['sitio']}/job/{a['Id']}",
            "detalle": "\n".join(_texto(x) for x in partes if x),
        }


# ---------------------------------------------------------------------------
# SAP SuccessFactors Career Site (Scotiabank, Telefónica, Arca Continental, Nestlé)
# ---------------------------------------------------------------------------
_LINK_SF = re.compile(r'<a (?=[^>]*jobTitle-link)[^>]*href="(/job/[^"]+/(\d+)/)"[^>]*>\s*([^<]+?)\s*</a>')


def _meta(t, prop):
    m = re.search(rf'itemprop="{prop}"[^>]*content="([^"]*)"', t)
    return html.unescape(m.group(1)) if m else ""


def _norm(x):
    return unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode().lower().strip()


# Cuando el portal solo da el distrito ("San Isidro, PE"), estos se reconocen como Lima.
DISTRITOS_LIMA = {
    "lima", "cercado de lima", "san isidro", "miraflores", "surco", "santiago de surco", "san borja",
    "la molina", "lince", "jesus maria", "magdalena", "magdalena del mar", "pueblo libre", "san miguel",
    "surquillo", "barranco", "chorrillos", "la victoria", "brena", "rimac", "ate", "santa anita",
    "san luis", "independencia", "los olivos", "san martin de porres", "comas", "lurin",
    "lurigancho", "chaclacayo", "san juan de lurigancho", "san juan de miraflores",
    "villa el salvador", "villa maria del triunfo", "chilca", "callao",
}


def _get_sf(s, url, **kw):
    """SuccessFactors (sobre todo Nestlé) a veces redirige a /errorpage/: reintenta con sesión nueva."""
    for intento in range(4):
        try:
            r = _get(s, url, **kw)
            if "/errorpage/" not in r.url:
                return r
        except requests.TooManyRedirects:
            pass
        s.cookies.clear()
        time.sleep(2 * (intento + 1))
    raise requests.RequestException(f"{url} respondió con página de error tras 4 intentos")


def _successfactors(p, interesa, desde=None):
    s = _sesion()
    vistos, candidatos, fila = set(), [], 0
    while True:
        r = _get_sf(s, f"{p['host']}/search/", params={
            "q": "", "sortColumn": "referencedate", "sortDirection": "desc",
            "startrow": fila, **p["params"]})
        encontrados = _LINK_SF.findall(r.text)
        nuevos = [(h, i, html.unescape(t)) for h, i, t in encontrados if i not in vistos]
        vistos.update(i for _, i, _ in nuevos)
        candidatos += [c for c in nuevos if interesa(c[2])]
        fila += 25
        if len(nuevos) < 25 or fila > 1000:
            break
    for href, jid, titulo in candidatos:
        try:
            t = _get_sf(s, p["host"] + href).text
        except requests.RequestException:
            continue  # un aviso que no abre no debe tumbar el portal entero
        # Nestlé no publica addressCountry: solo streetAddress = "San Isidro, PE".
        calle = _meta(t, "streetAddress")
        pais = _meta(t, "addressCountry") or (calle.rsplit(",", 1)[-1].strip() if "," in calle else "")
        if pais.lower() not in ("pe", "peru", "perú"):
            continue
        # La descripción va desde su <span> hasta el botón de postular.
        m = re.search(r'class="jobdescription">(.*?)(?:class="applylink|<script|Postular »|$)', t, re.S)
        publicada = _meta(t, "datePosted")  # "Sat Oct 10 02:00:00 UTC 2026"
        lugar = ", ".join(x for x in (_meta(t, "addressLocality"), _meta(t, "addressRegion")) if x)
        if not lugar and calle:
            distrito = calle.rsplit(",", 1)[0].strip()
            lugar = f"{distrito}, Lima" if _norm(distrito) in DISTRITOS_LIMA else distrito
        yield {
            "id": f"{p['clave']}-{jid}",
            "titulo": titulo,
            "ubicacion": lugar,
            "modalidad": "Remoto" if re.search(r"remot", titulo + lugar, re.I) else "",
            "jornada": "",
            "fecha_publicacion": datetime.strptime(publicada, "%a %b %d %H:%M:%S UTC %Y").date()
            if publicada else date.today(),
            "link": p["host"] + href,
            "detalle": _texto(m.group(1)) if m else "",
        }


# ---------------------------------------------------------------------------
# Genoma (Mibanco)
# ---------------------------------------------------------------------------
_MODALIDAD_GENOMA = {"HI": "Híbrido", "RE": "Remoto", "PR": "Presencial", "ON": "Presencial"}


def _genoma(p, interesa, desde=None):
    s = _sesion()
    s.headers["Accept"] = "application/json"
    api = "https://api.genoma.work/api/v1/gamesandtests"
    lista = _get(s, f"{api}/{p['slug']}/jobslisting/").json()["jobapplications"]
    for a in lista:
        if a.get("location_country") != "PE" or not interesa(a.get("job_application")):
            continue
        if desde and date.fromisoformat(a["creation_date"]) < desde:
            continue
        d = _get(s, f"{api}/jobapplications/{a['id']}/public/").json()
        if d.get("status") != "ACTIV":
            continue
        yield {
            "id": f"{p['clave']}-{a['id'][:8]}",
            "titulo": d["job_application"],
            "ubicacion": d.get("location_city") or "",
            "modalidad": _MODALIDAD_GENOMA.get(d.get("job_modality"), d.get("job_modality") or ""),
            "jornada": {"FT": "Tiempo completo", "PT": "Medio tiempo"}.get(d.get("working_hours"), ""),
            "fecha_publicacion": date.fromisoformat(d["creation_date"]),
            "link": f"https://app.genoma.work/jobapplication/{a['id']}/",
            "detalle": _texto(d.get("description")),
        }


# ---------------------------------------------------------------------------
# Rankmi ATS (Belcorp)
# ---------------------------------------------------------------------------
def _rankmi(p, interesa, desde=None):
    s = _sesion()
    s.headers["Accept"] = "application/json"
    url = f"{p['host']}/api/v1/tenants/{p['tenant']}/organizations/{p['organizacion']}/selection_processes"
    pagina = 1
    while True:
        r = s.post(url, params={"page": pagina, "items": 100}, json={}, timeout=30)
        time.sleep(PAUSA_SEG)
        r.raise_for_status()
        lista = r.json()
        for a in lista:
            if "peru" not in (a.get("country") or "").lower().replace("ú", "u") or not interesa(a.get("title")):
                continue
            lugar = a.get("long_location") or ""
            if lugar.strip().lower() in ("perú", "peru"):
                lugar = "Lima (aviso sin ciudad; sede central de Belcorp)"
            yield {
                "id": f"{p['clave']}-{a['id']}",
                "titulo": a["title"].strip(),
                "ubicacion": lugar,
                "modalidad": "Remoto" if a.get("is_remote") else "",
                "jornada": a.get("job_disponibility") or "",
                "fecha_publicacion": date.fromisoformat((a.get("postulation_date") or a["created_at"])[:10]),
                "link": p["host"] + a["link_view_postulation"],
                "detalle": (a.get("description") or "").replace("\r", "").strip(),
            }
        if pagina >= int(r.headers.get("total-pages") or 1):
            break
        pagina += 1


CONECTORES = {"workday": _workday, "oracle": _oracle, "successfactors": _successfactors,
              "genoma": _genoma, "rankmi": _rankmi}


def buscar_portal(p, interesa, desde=None):
    """Vacantes de Perú del portal `p` cuyo título pasa `interesa`, ya con fuente y empresa.
    `desde` (date) es opcional: los portales que traen la fecha en el listado la usan para no abrir avisos viejos."""
    for v in CONECTORES[p["tipo"]](p, interesa, desde):
        v["fuente"] = f"Portal {p['empresa']}"
        v["empresa"] = p["empresa"]
        yield v


def empresa_canonica(nombre, normalizar):
    """Nombre del portal si `nombre` (p. ej. el de Bumeran) es una de estas empresas."""
    n = normalizar(nombre)
    for p in PORTALES:
        if any(a in n for a in p["alias"]):
            return p["empresa"]
    return nombre
