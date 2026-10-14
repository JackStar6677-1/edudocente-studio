# -*- coding: utf-8 -*-
"""Sugerencias para completar las semanas en blanco de una planificacion mensual.

Dos motores:
  - "offline": reparte los indicadores de los OA entre las clases libres (sin internet, sin costo).
  - un proveedor de IA ("gemini", "claude", "openai", "deepseek", "kimi", "grok", "local"):
    usa la misma configuracion y claves de ai_connector.py (variables de entorno *_API_KEY).

Las sugerencias se guardan en cada clase como "sugerencia" y NO reemplazan lo que escribio el docente.
Regla: la IA no inventa paginas del texto escolar; si no las conoce deja "págs. ____".
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from ai_connector import AI_PROVIDERS, AIConnector   # noqa: E402

from .mensual import CAMPOS, resolver_oa   # noqa: E402

EVAL_FORMATIVA = 'Evaluación formativa a través de las actividades realizadas en clases.'
RECURSOS_BASE = 'Texto del estudiante, págs. ____; cuaderno de actividades.'

# verbos de indicador (3.ª persona plural) -> infinitivo, para escribir el objetivo de la clase
VERBOS = {
    'explican': 'Explicar', 'describen': 'Describir', 'identifican': 'Identificar', 'comparan': 'Comparar',
    'señalan': 'Señalar', 'representan': 'Representar', 'registran': 'Registrar', 'clasifican': 'Clasificar',
    'planifican': 'Planificar', 'construyen': 'Construir', 'relacionan': 'Relacionar', 'reconocen': 'Reconocer',
    'distinguen': 'Distinguir', 'investigan': 'Investigar', 'analizan': 'Analizar', 'formulan': 'Formular',
    'proponen': 'Proponer', 'cantan': 'Cantar', 'tocan': 'Tocar', 'participan': 'Participar', 'escuchan': 'Escuchar',
    'comentan': 'Comentar', 'utilizan': 'Utilizar', 'ejecutan': 'Ejecutar', 'realizan': 'Realizar',
    'mantienen': 'Mantener', 'conocen': 'Conocer', 'muestran': 'Mostrar', 'dan': 'Dar', 'infieren': 'Inferir',
    'determinan': 'Determinar', 'argumentan': 'Argumentar', 'miden': 'Medir', 'definen': 'Definir',
    'observan': 'Observar', 'dibujan': 'Dibujar', 'conectan': 'Conectar', 'resguardan': 'Resguardar',
    'discriminan': 'Discriminar', 'establecen': 'Establecer', 'evalúan': 'Evaluar', 'expresan': 'Expresar',
    'se expresan': 'Expresarse', 'reaccionan': 'Reaccionar', 'acompañan': 'Acompañar', 'reproducen': 'Reproducir',
}
ACTIVIDAD = {
    'Observar': 'Observan y comentan', 'Identificar': 'Identifican y registran en su cuaderno',
    'Explicar': 'Explican con sus palabras y con ejemplos', 'Describir': 'Describen y representan',
    'Comparar': 'Comparan y completan un cuadro', 'Construir': 'Construyen en parejas',
    'Investigar': 'Investigan y comunican sus conclusiones', 'Cantar': 'Cantan y repasan en grupo',
}


def _clases_libres(plan):
    for p in plan['periodos']:
        for n, c in enumerate(p['clases']):
            if not any(c.get(k, '').strip() for k in CAMPOS):
                yield p['nombre'], n, c


def _objetivo_desde_indicador(ind):
    ind = ind.strip()
    if ind.lower().startswith('se expresan'):
        base, resto = 'se expresan', ind[len('se expresan'):].strip()
    else:
        palabras = ind.split(' ', 1)
        base, resto = palabras[0].lower(), (palabras[1] if len(palabras) > 1 else '')
    inf = VERBOS.get(base)
    if not inf and base.endswith('an') and len(base) > 4:      # verbos en -ar: "asocian" -> "Asociar"
        inf = base[:-2].capitalize() + 'ar'
    if inf:
        return f'{inf} {resto}'.strip().rstrip('.') + '.', inf
    return f'Trabajar: {ind.rstrip(".")}.', 'Trabajar'


def sugerir_offline(plan):
    """Reparte, en orden, los indicadores de los OA entre las clases libres."""
    oas = resolver_oa(plan)
    inds = [(o['codigo'], i) for o in oas for i in o['indicadores']]
    libres = list(_clases_libres(plan))
    if not libres or not inds:
        return 0
    por_clase = max(1, round(len(inds) / len(libres)))
    for k, (_, _, clase) in enumerate(libres):
        trozo = inds[k * por_clase:(k + 1) * por_clase] or inds[-1:]
        obj, verbo = _objetivo_desde_indicador(trozo[0][1])
        accion = ACTIVIDAD.get(verbo)
        clase['sugerencia'] = {
            'objetivo': obj,
            'actividad': (f'{accion}: ' if accion else 'Actividad guiada para lograr: ') + trozo[0][1].rstrip('.') + f' (OA {trozo[0][0]}).',
            'evaluacion': plan.get('evaluacion_por_defecto', EVAL_FORMATIVA),
            'recursos': plan.get('recursos_por_defecto', RECURSOS_BASE),
        }
    return len(libres)


# ---------------------------------------------------------------- proveedores de IA
def _pedir(proveedor, system, user, modelo=None, api_key=None):
    ai = AIConnector(proveedor, api_key=api_key, model=modelo)
    info = AI_PROVIDERS[ai.provider]
    if not ai.api_key and ai.provider != 'local':
        raise RuntimeError(f'Falta la clave {info["env_key"]} (variable de entorno) para usar {info["name"]}.')
    if ai.provider == 'claude':
        r = ai._call_http(ai.endpoint, {'x-api-key': ai.api_key, 'anthropic-version': '2023-06-01', 'content-type': 'application/json'},
                          {'model': ai.model, 'max_tokens': 4000, 'system': system, 'messages': [{'role': 'user', 'content': user}]})
        return ''.join(b.get('text', '') for b in r.get('content', []))
    if ai.provider == 'gemini':
        url = ai.endpoint.format(model=ai.model, api_key=ai.api_key)
        r = ai._call_http(url, {'content-type': 'application/json'},
                          {'systemInstruction': {'parts': [{'text': system}]}, 'contents': [{'parts': [{'text': user}]}]})
        return r['candidates'][0]['content']['parts'][0]['text']
    r = ai._call_http(ai.endpoint, {'Authorization': f'Bearer {ai.api_key}', 'Content-Type': 'application/json'},
                      {'model': ai.model, 'messages': [{'role': 'system', 'content': system}, {'role': 'user', 'content': user}]})
    return r['choices'][0]['message']['content']


SISTEMA = ('Eres asistente de planificación docente en Chile (Mineduc). Redactas filas de una planificación mensual: '
           'objetivo de la clase (infinitivo), estrategia de actividad (qué hacen los estudiantes, 3.ª persona plural), '
           'estrategia de evaluación y recursos. Sé breve y concreto. NO inventes números de página del texto escolar: '
           'si no los conoces escribe "págs. ____". Responde SOLO con un arreglo JSON, un objeto por clase, con las claves '
           '"objetivo","actividad","evaluacion","recursos".')


def sugerir_con_ia(plan, proveedor, modelo=None):
    oas = resolver_oa(plan)
    libres = list(_clases_libres(plan))
    if not libres:
        return 0
    escritas = [f'{p["nombre"]} {n + 1}: {c["objetivo"]}' for p in plan['periodos'] for n, c in enumerate(p['clases']) if c.get('objetivo')]
    user = json.dumps({
        'curso': plan['curso'], 'asignatura': plan['asignatura'], 'unidad': plan['unidad'],
        'oa': [{'codigo': o['codigo'], 'texto': o['texto'], 'indicadores': o['indicadores']} for o in oas],
        'clases_ya_escritas': escritas, 'clases_a_sugerir': [f'{nom} {n + 1}' for nom, n, _ in libres],
        'evaluacion_por_defecto': plan.get('evaluacion_por_defecto', EVAL_FORMATIVA),
    }, ensure_ascii=False)
    bruto = _pedir(proveedor, SISTEMA, user, modelo)
    m = re.search(r'\[.*\]', bruto, re.S)
    datos = json.loads(m.group(0))
    for (_, _, clase), sug in zip(libres, datos):
        clase['sugerencia'] = {k: str(sug.get(k, '')) for k in CAMPOS}
    return min(len(libres), len(datos))


def sugerir(plan, motor='offline', modelo=None):
    return sugerir_offline(plan) if motor == 'offline' else sugerir_con_ia(plan, motor, modelo)
