# Kit de CVs con Claude Code

Un sistema para buscar trabajo con [Claude Code](https://claude.com/claude-code) como copiloto. Busca vacantes en los portales de las empresas que te interesan, las califica contra tu experiencia real, investiga a la empresa y genera un CV de una página distinto para cada postulación, sin inventar nada.

El código lo escribí junto con Claude Code, y está pensado para usarse de la misma forma: **tú conversas, Claude ejecuta los scripts y tú decides a qué postular**.

👉 Ejemplo de CV generado (perfil ficticio): [`docs/cv_ejemplo.pdf`](docs/cv_ejemplo.pdf)

---

## Cómo trabajamos (lo más importante)

### 1. Cuéntale todo a tu Claude
El kit funciona tan bien como lo que Claude sabe de ti. La primera vez que abras Claude Code en esta carpeta, te va a entrevistar (está en `CLAUDE.md`). Dale todo:

- Cada trabajo, práctica, freelance o proyecto: qué hacías, para quién, con qué herramientas y **qué cambió gracias a ti, con cifras**.
- Qué puedes defender de verdad en una entrevista técnica por cada herramienta, y qué solo viste en un curso.
- Qué puestos buscas, en qué ciudades, si aceptas remoto, cuánto quieres ganar.
- Lo que no quieres que aparezca (clientes confidenciales, un trabajo que no terminaste, etc.).

Claude guarda todo eso en `experiencia_maestra.md`, la **fuente de verdad**, y con eso arma un banco de bullets aprobados en `cv_data.xlsx`. Cada CV es una selección y un reordenamiento de ese banco según la vacante. Si un dato no está en la maestra, no entra al CV.

### 2. Claude propone, tú decides
- Claude busca, califica y analiza vacantes, pero **solo tú** marcas a cuáles postulas.
- Claude te muestra cada CV antes de enviarlo y te dice qué palabras clave de la vacante **no** pudo cubrir porque no son verdad.
- Cuando dices que postulaste, el CV se **congela** en `enviados/` y no se vuelve a tocar, para que siempre sepas exactamente qué enviaste.

### 3. Corrige a Claude y él aprende
Si algo no te gusta ("no menciones la IA si la vacante no la pide", "ese bullet exagera", "el SQL no me pesa tanto"), díselo. Las reglas que valen para siempre van a la sección **"Mis criterios"** de `CLAUDE.md`, y los datos nuevos a la maestra.

---

## Qué hace

```
 buscar_vacantes.py ──► vacantes.xlsx ──► Claude califica (encaje 0–100) y arma el podio
   (portales + Bumeran)       │
                              ▼
                 skill analizar-vacante-cv: empresa · puesto · encaje · sueldo · traslado
                              │
                              ▼
          radar/ (página con las vacantes analizadas; marcas "Postular" / "Analizar")
                              │  tú decides
                              ▼
   cv_data.xlsx + experiencia_maestra.md ──► generar_cv.py ──► salida/CV.docx + .pdf
                                                              └► enviados/ (congelado)
```

- **Búsqueda** (`buscar_vacantes.py`, `portales_empresas.py`): consulta los mismos servicios públicos que usan las webs de empleo, sin login. Trae conectores para Workday, Oracle Recruiting, SAP SuccessFactors, Genoma y Rankmi, más Bumeran Perú. Deduplica entre fuentes y prioriza el portal de la empresa. Lo que ya revisaste pasa al Historial y no vuelve a aparecer.
- **Análisis** (`.claude/skills/analizar-vacante-cv/`): Claude investiga en la web la reputación de la empresa (sanciones laborales, denuncias ambientales, clima laboral, finanzas), evalúa si el puesto está bien planteado, mide tu encaje con evidencia (`bullet_id`), estima un rango de sueldo y el tiempo de traslado desde tu casa.
- **CV** (`generar_cv.py`): una página, una columna, apto para ATS. Genera Word y PDF, avisa si pasa de una página y valida que todos los IDs existan.
- **Formularios:** pídele a Claude respuestas para las preguntas de postulación ("máximo 200 caracteres"). Responde solo con hechos de tu maestra.

## Empezar

```bash
git clone https://github.com/BrunoRioja/kit-cv-automatizado-claude-code.git
cd kit-cv-automatizado-claude-code
pip install -r requirements.txt

# Prueba con el perfil ficticio:
cp ejemplo/cv_data.xlsx ejemplo/experiencia_maestra.md .
python generar_cv.py --cv EJEMPLO-BI-RETAIL     # → salida/CV_Camila_Rojas_EJEMPLO-BI-RETAIL.pdf
```

El PDF se crea con Microsoft Word (vía `docx2pdf`) o, si no tienes Word, con LibreOffice.

Después abre Claude Code en la carpeta y dile algo como *"quiero armar mi perfil, entrevístame"*. Él reemplaza el ejemplo con tu información. A partir de ahí, las frases que usarás son:

| Le dices a Claude | Qué pasa |
|---|---|
| "busca vacantes" | Corre la búsqueda, califica, arma el podio, analiza el top y arma el radar |
| "analiza esta vacante: <link o texto>" | La agrega y aplica la skill completa |
| "ya marqué a cuáles postulo" | Genera, revisa y congela un CV por cada una |
| "genera el CV para <ficha>" | Un CV puntual, en borrador hasta que digas "listo" |
| "responde estas preguntas, máximo 200 caracteres" | Respuestas de formulario con tus hechos reales |
| "lo demás pásalo a historial" | Archiva lo revisado y conserva tus postulaciones |

## Hazlo tuyo: criterios personalizables

Todo lo que refleja *mis* preferencias está separado para que lo cambies:

| Qué | Dónde | Ejemplo del autor |
|---|---|---|
| Cómo evaluar empresas | `CLAUDE.md` → "Mis criterios" | Pesa mucho cómo trata a sus empleados (SUNAFIL, reseñas, Merco/GPTW) y si tiene denuncias ambientales (OEFA) |
| Brechas que pesan poco | `CLAUDE.md` → "Mis criterios" | SQL resta máximo 5 puntos porque se aprende rápido |
| Qué nunca va en el CV | `CLAUDE.md` → "Mis criterios" | Sin referencias; IA solo si la vacante la pide |
| Puestos que busco | `PALABRAS_CLAVE` en `buscar_vacantes.py` | Datos/BI, comercial, producto |
| Ciudades | `UBICACIONES` en `buscar_vacantes.py` | Lima, Callao, Chiclayo o remoto |
| Títulos que descarto | `EXCLUIR_TITULO` en `buscar_vacantes.py` | Gerencias, jefaturas, prácticas |
| Empresas objetivo | `PORTALES` en `portales_empresas.py` | Bancos y consumo masivo de Perú |
| Cómo investigar empresas | `.claude/skills/analizar-vacante-cv/references/` | Consultas para SUNAFIL, OEFA, Poder Judicial, SMV, Glassdoor |
| Diseño del CV | constantes al inicio de `generar_cv.py` | Calibri 9.5, azul marino, A4 |

¿Otras ideas? Penalizar sectores que no quieres (apuestas, tabaco…), exigir sueldo publicado, priorizar empresas con certificación B o de equidad de género. Escríbelo en "Mis criterios" con tus palabras y Claude lo aplica.

### Agregar una empresa
Busca su página de "Trabaja con nosotros". Si la URL es de `myworkdayjobs.com`, `oraclecloud.com`, `successfactors` (o `/search/?q=` con `jobTitle-link`), `genoma.work` o `rankmi.com`, basta con agregar una entrada en `PORTALES` con sus parámetros (mira los ejemplos). Si usa otro sistema, pídele a Claude que escriba el conector: le toma una conversación.

## Privacidad

Tus datos **no** deben ir a un repo público. El `.gitignore` ya excluye `cv_data.xlsx`, `experiencia_maestra.md`, `vacantes.xlsx`, `postulaciones/`, `salida/`, `enviados/` y el radar generado. Antes de hacer `git add`, revisa con `git status` que no aparezcan. Tu dirección (para estimar traslados) guárdala en la memoria de Claude, no en archivos.

## Uso responsable

- Los scripts usan los mismos servicios públicos que las páginas de empleo, sin iniciar sesión, una consulta a la vez y con pausas. Úsalos para tu búsqueda personal y de forma moderada; revisa los términos de cada sitio.
- **LinkedIn no se consulta con scripts** (riesgo para tu cuenta). Pega esas vacantes en el chat.
- Los portales cambian: si un conector deja de traer resultados, el script lo avisa y sigue con los demás. Pídele a Claude que lo revise.
- El análisis de empresas sale de búsquedas web: verifica lo importante antes de tomar decisiones.

## Estructura

```
├── CLAUDE.md                  # Instrucciones para Claude: flujos, reglas y "Mis criterios"
├── generar_cv.py              # Excel → CV .docx + .pdf
├── buscar_vacantes.py         # Búsqueda, deduplicación, Historial, --manual, --archivar
├── portales_empresas.py       # Conectores por tipo de portal
├── plantilla_postulacion.md   # Formato de la ficha por vacante
├── radar/                     # Página con las vacantes analizadas
├── .claude/skills/analizar-vacante-cv/   # Skill de análisis profundo
├── ejemplo/                   # Perfil ficticio (Camila Rojas Vega) para probar
└── docs/cv_ejemplo.pdf
```

## Licencia

MIT. Úsalo, cámbialo y compártelo.
