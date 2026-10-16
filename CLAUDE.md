# Kit de CVs con Claude Code — instrucciones para Claude

Esta carpeta busca vacantes, las califica y analiza contra el perfil del usuario, y genera un CV personalizado (.docx + .pdf) para cada postulación. Claude Code hace el trabajo; el usuario decide a qué postula.

> **Para quien usa el kit:** edita la sección **"Mis criterios"** (al final). Ahí van tus reglas personales; el resto del archivo es el flujo de trabajo y puedes dejarlo como está.

## Archivos

| Archivo | Rol | Quién lo edita |
|---|---|---|
| `experiencia_maestra.md` | Fuente de verdad de **hechos** (cifras, contexto, nivel real por habilidad) | Solo con datos que el usuario confirme |
| `cv_data.xlsx` | Texto **aprobado** de los CVs: banco de bullets, roles, habilidades, extras y una fila por CV | Usuario y Claude |
| `postulaciones/AAAA-MM-empresa-puesto.md` | Una ficha por vacante (formato en `plantilla_postulacion.md`) | Claude |
| `generar_cv.py` | Lee el Excel y genera `salida/<prefijo>_<cv_id>.docx` y `.pdf`. El diseño vive en el código | Claude, solo si el usuario pide cambiar el diseño |
| `enviados/` | Copia congelada de cada CV enviado. **Nunca se modifica** | Lo llena el script |
| `buscar_vacantes.py` | Busca vacantes en los portales de las empresas objetivo y en Bumeran Perú (servicios públicos de sus webs, sin login) y agrega solo las nuevas a `vacantes.xlsx`. `--solo-portales` / `--sin-portales`; `--manual` agrega una vacante pegada en el chat; `--archivar` mueve todo a Historial (`--conservar-con-cv` deja las que ya tienen CV). Palabras clave, ubicaciones y exclusiones están al inicio del script. Si la misma vacante sale en el portal de la empresa y en Bumeran, queda una sola fila con el link del portal | Claude |
| `portales_empresas.py` | Conectores por tipo de portal (Workday, Oracle Recruiting, SAP SuccessFactors, Genoma, Rankmi) y la lista de empresas configuradas | Claude |
| `vacantes.xlsx` | Hoja **Vacantes**: la revisión en curso. Hoja **Historial**: todo lo ya revisado. Ninguna vacante del Historial vuelve a entrar (mismo id, o misma empresa + título). Las columnas de calificación, análisis, `postular` y `cv_id` las llenan Claude o el usuario; el script nunca las pisa | Script + Claude + usuario |
| `radar/` | `generar_radar.py` arma `radar_vacantes.html` desde la hoja Vacantes. Si Claude Code puede publicar artefactos, se publica como artefacto "Radar de vacantes" (guarda su URL aquí: `RADAR_URL = …`) y el usuario marca **Postulé** / **Analizar** desde el navegador | Claude |
| `.claude/skills/analizar-vacante-cv/` | Skill de análisis profundo de una vacante (empresa, puesto, encaje, sueldo, traslado) | Usuario |

### Hojas de `cv_data.xlsx`
- **Perfil**: datos de contacto (clave / valor). `prefijo_archivo` da nombre a los archivos.
- **Roles**: `rol_id`, empresa, lugar, cargo, fechas, `orden` (menor = más arriba). Los roles consecutivos de la misma empresa se agrupan bajo un solo encabezado.
- **Banco**: `bullet_id`, `rol_id`, texto, etiquetas, notas. Bullets aprobados y reutilizables.
- **Habilidades**: `hab_id`, categoría, items.
- **Extras**: `extra_id`, sección (Educación / Proyectos académicos / Certificaciones), texto, fecha, detalle, `por_defecto`.
- **CVs**: una fila por CV. `cv_id`, estado, empresa, puesto, postulacion_md, fecha, ubicacion, titular, resumen, bullets, habilidades, extras, notas.
  - `bullets`: IDs separados por coma, **en el orden en que deben aparecer dentro de cada rol**. Solo aparecen los roles con al menos un bullet.
  - `habilidades` y `extras`: IDs separados por coma. Vacíos = todas las habilidades / los extras con `por_defecto = sí`.
  - `ubicacion`: vacía = la del Perfil. Con `-` la cabecera no muestra ubicación.
  - `estado`: `borrador` | `listo` | `enviado`. Las filas `enviado` no se regeneran.
- **Overrides**: `cv_id`, `bullet_id`, `rol_id`, texto. Reescribe un bullet solo para ese CV sin tocar el Banco, o agrega uno que solo existe en ese CV (`rol_id` obligatorio). Debe estar listado en la columna `bullets` del CV.

## Comandos

```bash
pip install -r requirements.txt
python generar_cv.py --cv BASE-DATOS     # genera un CV
python generar_cv.py --validar           # revisa IDs y reglas sin generar nada
python generar_cv.py --enviar BASE-DATOS # copia a enviados/ y marca la fila como "enviado"
python buscar_vacantes.py                # busca vacantes nuevas (últimos 30 días)
python radar/generar_radar.py            # arma el radar
```

Antes de correr cualquier cosa que escriba un Excel, el usuario debe tenerlo cerrado (Excel lo bloquea en Windows).

## Primera sesión: conocer al usuario

Si `experiencia_maestra.md` o `cv_data.xlsx` no existen o todavía tienen el perfil ficticio de `ejemplo/`, **entrevista al usuario antes de cualquier otra cosa**:
1. Pídele su CV actual (PDF, Word o texto) y su LinkedIn si quiere.
2. Pregúntale rol por rol: qué hacía, para quién, con qué herramientas, qué cambió gracias a su trabajo y con qué cifras. Repregunta cuando una cifra sea vaga ("¿cuántos usuarios?", "¿de cuánto a cuánto bajó?").
3. Pregúntale qué puede defender en una entrevista técnica por cada herramienta (eso llena la tabla "Nivel real por habilidad").
4. Pregúntale qué puestos busca, en qué ciudades, si acepta remoto, y dónde vive (para estimar traslados; guárdalo en tu memoria, no en archivos del repo).
5. Escribe `experiencia_maestra.md` con [confirmado] / [verificar], después arma el Banco, Roles, Habilidades y Extras de `cv_data.xlsx` y un CV `BASE-…`. Muéstraselo y ajusta hasta que diga "listo".
6. Repasa con el usuario la sección "Mis criterios" y actualízala con lo que diga.

## Flujo cuando el usuario dice "genera el CV para <vacante o ficha>"

1. Lee la ficha en `postulaciones/` completa, incluido el texto de la vacante.
2. Lee `experiencia_maestra.md` y el Excel.
3. Crea un `cv_id` corto en MAYÚSCULAS (`EMPRESA-PUESTO`). Si ya existe, propón actualizarlo en lugar de duplicarlo.
4. Personaliza **solo** esto, en este orden de prioridad:
   1. **Titular**: el nombre del puesto de la vacante (o el más cercano verdadero) + 3 herramientas que la vacante pide y el usuario domina.
   2. **Resumen**: 2–3 líneas con el vocabulario exacto de la vacante, anclado en experiencia real, no en cursos.
   3. **Selección y orden de bullets** del Banco según el carril y los requisitos.
   4. **Orden de habilidades** y qué extras entran.
   5. **Overrides**: solo si un bullet del Banco no usa una palabra clave que la vacante repite y el cambio es verdadero. Máximo 3–4 por CV.
5. Escribe la fila en la hoja CVs (y en Overrides si aplica) con estado `borrador`, guardando con openpyxl sin borrar otras hojas ni filas.
6. Corre `python generar_cv.py --cv <ID>` y abre el PDF. Verifica que sea **1 página**, que no haya líneas viudas feas y que cada palabra clave obligatoria de la vacante aparezca donde sea verdadera.
7. Reporta en 3–5 líneas qué cambió respecto de la base, qué palabras clave quedaron cubiertas y cuáles **no** se pudieron cubrir porque no son verdad.

## Flujo cuando el usuario dice "busca vacantes"

1. Con `vacantes.xlsx` cerrado, corre `python buscar_vacantes.py`. Si un portal falla, el script lo avisa y sigue: menciónalo en el reporte.
2. Para cada fila con `estado = nueva`, lee `detalle` y compárala con `experiencia_maestra.md` aplicando las reglas de honestidad y "Mis criterios".
3. Llena `encaje` (0–100), `cubre`, `falta` y `alertas` (años pedidos muy por encima, idioma, sueldo/condiciones dudosas, puesto que en realidad es de ventas o soporte). Cambia `estado` a `calificada`. No toques `postular`: es del usuario.
4. Calcula `podio`: 1–20 por `encaje` (desempate: más reciente); el resto en grupos `B · encaje medio (50+)`, `C · encaje bajo (30–49)`, `D · descartar (<30)`. Fuera del ranking: `Ya postulada` (empresa/puesto ya en la hoja CVs) y `Duplicado`.
5. Reporta las 5–10 mejores con una línea cada una y su link. Luego sigue "analiza el top" y regenera el radar.
6. LinkedIn no se consulta con scripts (riesgo para la cuenta del usuario).

## Flujo cuando el usuario dice "analiza el top" o "analiza las que marqué"

1. Aplica la skill `analizar-vacante-cv` a cada vacante: investigación web de la empresa, evaluación del puesto, encaje contra el Banco y la tabla de nivel real (con `bullet_id` como evidencia), sueldo, traslado y veredicto.
2. Escribe los resultados en las columnas de análisis de `vacantes.xlsx` (incluidas `bullets_sugeridos` y `traslado`) y pon `estado = analizada`.
3. Corre `python radar/generar_radar.py` y, si hay artefacto, republícalo en su misma URL.

## Vacantes que el usuario pega en el chat

1. Arma un JSON con `titulo`, `empresa`, `detalle` y lo que haya de `ubicacion`, `modalidad`, `link`, `fecha_publicacion` (AAAA-MM-DD), `fuente`.
2. Corre `python buscar_vacantes.py --manual <json>`. Si responde `YA_VISTA`, avisa en qué hoja está.
3. Si responde `AGREGADA: <id>`, califícala, aplica la skill y escribe el análisis (`podio = Manual`).

## Flujo cuando el usuario dice "ya marqué a cuáles postulo"

1. Lee las marcas (artefacto o columna `postular = Y`).
2. Para cada vacante marcada sin `cv_id`: crea su ficha en `postulaciones/`, sigue "genera el CV" usando `bullets_sugeridos`, verifica 1 página y congélalo con `python generar_cv.py --enviar <ID>` (postular = enviado).
3. Escribe el `cv_id` en `vacantes.xlsx` y pon `estado = cv_enviado`.
4. Si el usuario dice "lo demás a Historial", usa `python buscar_vacantes.py --archivar --conservar-con-cv`: las postuladas siguen a la vista. Solo `--archivar` a secas si dice "todo".

## Reglas no negociables

- **No inventar.** Toda cifra y todo hecho deben estar en `experiencia_maestra.md`. Lo marcado **[verificar]** no se usa sin preguntar. Si falta una cifra, pregunta; si el usuario da una estimación, anótala en `notas` como estimada.
- **Honestidad de nivel:** la tabla "Nivel real por habilidad" de la maestra manda. Nunca subir de nivel una herramienta para calzar con una vacante.
- **Confidencialidad:** no nombrar clientes de empleadores si la maestra lo prohíbe; usar sectores.
- **Formato:** 1 página, una columna, sin tablas ni imágenes (ATS). El diseño está en `generar_cv.py`.
- **Fechas** con formato `Mmm AAAA – Mmm AAAA` (ej. `Abr 2026 – Jul 2026`).
- Bullets de 1–2 líneas: verbo de acción en pasado + resultado + cifra + método. Sin "responsable de", "apoyé en", "ayudé con".
- **Dashboards:** describirlos por su propósito, sus usuarios y su impacto, no por los tipos de gráfico.
- Un dato nuevo que cuente el usuario va **primero** a `experiencia_maestra.md` y después, si aplica, al Banco.
- **CVs enviados se congelan:** nunca modificar `enviados/`; nunca regenerar una fila `enviado`; generar siempre con `--cv <ID>`. Si un cambio en hojas compartidas afectaría a CVs en `borrador` o `listo`, avisar y preguntar antes de regenerarlos.
- **Respuestas para formularios de postulación:** solo hechos de la maestra; respeta el límite de caracteres que indique el usuario y cuenta los caracteres antes de entregar.

---

## Mis criterios (edítalos)

Estos son los criterios personales del autor del kit. Cámbialos, bórralos o agrega los tuyos: Claude los aplica al calificar, analizar y escribir.

### Qué busco
- **Carriles** (en orden): Analista de Datos / BI → Analista Comercial / Inteligencia Comercial → Producto / Proyectos. Las palabras clave de búsqueda están en `PALABRAS_CLAVE` de `buscar_vacantes.py`.
- **Ubicaciones válidas:** Lima, Callao, Chiclayo/Lambayeque o remoto (`UBICACIONES` en `buscar_vacantes.py`).
- **Entran** puestos senior y títulos en inglés; **se descartan** gerencia, jefaturas y prácticas (`EXCLUIR_TITULO`).
- **Empresas objetivo:** bancos y consumo masivo grandes de Perú (`PORTALES` en `portales_empresas.py`). Ventana de búsqueda: 30 días.

### Cómo evalúo a una empresa
- **Cómo trata a su gente pesa mucho:** sanciones de SUNAFIL, demandas laborales, reseñas de empleados (patrones, no una reseña aislada), rankings como Merco Talento o Great Place to Work.
- **Denuncias ambientales:** revisar OEFA y prensa en sectores de riesgo (minería, energía, industria, pesca, agroindustria).
- **Corrupción y ética:** escándalos o investigaciones recientes.
- **Salud financiera:** utilidades y crecimiento (SMV si cotiza), despidos recientes.
- **Empresa confidencial:** sin nombre no se califica la empresa; recomendar preguntarlo en el primer contacto.
- Que no aparezca nada malo no equivale a buena reputación confirmada: decirlo así.

### Cómo evalúo el puesto
- **Bandera roja:** recibo por honorarios para un trabajo dependiente, "beneficios de ley" vendidos como extra, "todólogo" (3–4 puestos en uno), requisitos desproporcionados para el nivel.
- **Traslado:** más de ~1h por trayecto en presencial baja la nota del puesto.
- **Sueldo:** siempre como rango en soles, ajustado al tamaño de la empresa y la modalidad.

### Brechas que pesan poco
- **SQL** (el autor lo enseñó pero no lo usó en producción): se anota en `falta`, pero resta como mucho ~5 puntos de `encaje` aunque sea indispensable. En el CV sigue sin afirmarse uso en producción.

### Qué nunca va en mis CVs
- Sección de referencias y nivel de inglés.
- IA (Claude, Copilot, etc.): solo si la vacante la pide.
- Calificativos de nivel ("experto", "avanzado") en las habilidades: que lo demuestren los bullets.
