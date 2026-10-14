# Planificaciones (mensual y anual) · EduDocente-Studio

Genera en formato del colegio (oficio horizontal, logo, tablas de OA, indicadores y clases) dos documentos:

- **Planificación mensual** (o de varios meses, p. ej. Noviembre–Diciembre).
- **Plan anual** (objetivos, indicadores y grilla de semanas con marcas **X** = "este objetivo se aborda en esa semana").

La docente solo escribe **lo de cada clase**. Todo lo demás (OA, indicadores, numeración, franja DICIEMBRE, estilos) lo arma el módulo.
No requiere internet ni IA; la IA es opcional y solo *sugiere*.

```
python -m planificacion <comando> ...        (desde la raíz del proyecto)
```

## Flujo para la docente (sin tocar código)

1. Pedir (o ejecutar) `generar-mensual` → sale el Word con las filas vacías.
2. Escribir en Word: objetivo de la clase, estrategia de actividad, estrategia de evaluación y recursos.
3. Si quiere ayuda: `leer-mensual` + `sugerir` (ver abajo). Las sugerencias salen en **gris cursiva** y no pisan lo que ella escribió.

## Flujo para una IA (pensado para gastar pocos tokens)

Nunca hace falta leer ni editar el .docx a mano. Se trabaja con un JSON chico y comandos cortos:

| Paso | Comando | Qué devuelve |
|---|---|---|
| Ver qué falta | `estado plan.json` | 3–4 líneas: clases completas `✓`, parciales `~`, vacías `·` |
| Leer lo que escribió la docente en Word | `leer-mensual archivo.docx --curso 5°AB` | `archivo.json` actualizado |
| Escribir una clase | `semana plan.json Noviembre 3 --objetivo "..." --actividad "..." --evaluacion "..." --recursos "..."` | confirmación |
| Proponer las clases vacías | `sugerir plan.json` (offline) o `sugerir plan.json --ia claude` | cuántas clases con sugerencia |
| Dar por buenas las sugerencias | `aceptar plan.json --todas` o `--clases Noviembre:3 Diciembre:1` | cuántas aceptó |
| Generar el Word | `generar-mensual plan.json [--sugerencias]` | ruta del .docx |

Reglas para la IA:
- **No inventar páginas del texto escolar** (el módulo no tiene acceso a los libros). Si no las conoce: `págs. ____`.
- Objetivo de la clase en infinitivo ("Analizar…"), actividad en 3.ª persona plural ("Observan…").
- Los OA e indicadores **no se escriben**: salen del banco curricular por número de OA.
- La evaluación por defecto es formativa; la docente puede cambiarla (p. ej. "Evaluación sumativa de la unidad, fecha por confirmar").
- Proveedores de `--ia`: `gemini`, `claude`, `openai`, `deepseek`, `kimi`, `grok`, `local` (usan las claves `*_API_KEY` de `ai_connector.py`). Sin clave → usar `offline`.

### Formato del JSON mensual (resumen)

```json
{
 "anio": 2026, "profesor": "…", "curso": "5°A – 5°B", "curso_codigo": "5°AB", "asignatura": "Ciencias Naturales",
 "unidad": {"numero": "4", "nombre": "Energía eléctrica."}, "horas": "…",
 "oa": ["8", "9", "10", "11"],
 "periodos": [
   {"nombre": "Noviembre", "clases": [{"objetivo": "", "actividad": "", "evaluacion": "", "recursos": ""}]},
   {"nombre": "Diciembre", "clases": [{"objetivo": "", "actividad": "", "evaluacion": "", "recursos": ""}]}
 ]
}
```
El primer período va sin franja; los siguientes se separan con una franja con su nombre (DICIEMBRE). Se pueden usar uno o varios meses.
Un OA también puede escribirse a mano: `{"codigo": "3", "texto": "…", "puntos": [], "indicadores": ["…"]}` (útil si aún no está en el banco).
Si el plan anual del curso tiene otro nombre (p. ej. Orientación 5.º A usa el plan anual 5.º AB) usar `"banco_curso": "5°AB"`.

## Crear una planificación nueva

```
python -m planificacion nuevo-mensual --curso 5°AB --asignatura "Ciencias Naturales" --anio 2026 \
   --unidad 4 --nombre-unidad "Energía eléctrica." --oa 8 9 10 11 --mes Noviembre:15 Diciembre:2 -o plan.json
python -m planificacion generar-mensual plan.json
```
`--mes Mes:N` define cuántas clases tiene cada mes (las filas salen numeradas 1…N y vuelven a 1 en cada mes nuevo).

## Banco curricular (OA e indicadores) y Plan Anual

El banco son archivos JSON en `planificacion/datos/banco/<año>/` sacados de los Planes Anuales:

```
python -m planificacion banco-importar "5°AB - Ciencias Naturales - Plan Anual 2026.docx" --curso 5°AB --asignatura "Ciencias Naturales" --anio 2026
python -m planificacion banco-importar plan.docx --curso 5°AB --asignatura Orientación --anio 2026 --codigo "Practicar en forma=4"   # OA sin número en el Word
```
Para **2027**: escribir los objetivos del Mineduc en el plan anual (en Word o en el JSON), importarlo con `banco-importar --anio 2027` y las planificaciones mensuales lo usan solas.

```
python -m planificacion nuevo-anual --curso 5°AB --asignatura "Ciencias Naturales" --anio 2027 --horas "4 hrs Semanales" --clases "10 Mensuales"
python -m planificacion generar-anual "5°AB - Ciencias Naturales - Plan Anual 2027.json"      # plantilla en blanco o con OA
```
Formato de un OA del plan anual: `{"semestre": 2, "unidad": 4, "codigo": "9", "texto": "…", "puntos": [], "indicadores": ["…"], "marcas": {"NOVIEMBRE": [1, 2], "DICIEMBRE": [1]}}`.
Cada semestre termina con el cuadro **OA Enseñados / OA No Enseñados / % Cobertura Curricular**.

## Archivos

| Archivo | Rol |
|---|---|
| `cli.py`, `__main__.py` | comandos |
| `mensual.py`, `anual.py` | JSON → Word |
| `banco.py` | Plan Anual (.docx) → banco JSON; búsqueda de OA por número y unidad |
| `leer.py` | Word escrito por la docente → JSON |
| `ia.py` | sugerencias (offline o proveedor de IA) |
| `docx_base.py` | estilos y tablas |
| `test_planificacion.py` | `python -m unittest planificacion.test_planificacion` |
| `datos/` | **local, no se sube a GitHub**: `colegio.json` (docente y cursos), `banco/`, `mensual/` |
| `ejemplos/colegio.ejemplo.json` | plantilla de `datos/colegio.json` |

## Pendiente / límites

- El módulo no lee los libros del Mineduc; los números de página los pone la docente.
- Los indicadores se toman tal cual del plan anual; si cambia el currículum hay que reimportar el banco.
