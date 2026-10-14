# -*- coding: utf-8 -*-
"""Formatos editables (perfiles) para planificacion mensual, plan anual y temario.

Los formatos incluidos son *marcadores de posicion* (placeholders): reproducen un formato escolar tipico, pero
cualquier establecimiento puede cambiarlos: colores, fuente, titulos, columnas (renombrar, reordenar, agregar,
quitar), campos de la cabecera, bloques visibles, textos de pie, etc.

Orden de prioridad (de menor a mayor):  formato por defecto  <  perfil guardado en datos/formatos  <  "formato" dentro del plan.
"""
import copy
import json
import os
import re

RAIZ = os.path.dirname(os.path.abspath(__file__))
CARPETA = os.path.join(RAIZ, 'datos', 'formatos')

TIPOS = ('mensual', 'anual', 'temario', 'evaluacion')

COLORES = {'primario': '1F3864', 'secundario': '2E5597', 'claro': 'DCE6F2', 'fondo': 'F4F7FB', 'borde': '8EA9DB'}
PAGINA = {'ancho_cm': 33.0, 'alto_cm': 21.6, 'margen_lateral_cm': 1.5, 'margen_superior_cm': 1.0, 'margen_inferior_cm': 1.1}

DEFECTO = {
    'mensual': {
        'nombre': 'Predeterminado (placeholder)',
        'fuente': 'Arial',
        'colores': dict(COLORES),
        'pagina': dict(PAGINA),
        'titulo': 'PLANIFICACIÓN MENSUAL',
        'subtitulo': '{meses} {anio}',
        'mostrar_logo': True,
        'institucion': ['Complejo Educacional Particular', '', 'UTP {anio}'],
        # Cabecera de datos: lista de filas; cada fila es una lista de pares [etiqueta, valor].
        # Valores posibles: {profesor} {curso} {asignatura} {meses} {unidad_numero} {unidad_nombre} {horas} {anio}
        'datos': [
            [['PROFESOR:', '{profesor}'], ['CURSO:', '{curso}'], ['ASIGNATURA:', '{asignatura}'], ['MES:', '{meses}']],
            [['UNIDAD PROGRAMÁTICA N°:', '{unidad_numero}'], ['NOMBRE DE LA UNIDAD:', '{unidad_nombre}'], ['N° HORAS MENSUAL:', '{horas}']],
        ],
        'bloques': ['datos', 'oa', 'clases', 'indicadores'],          # orden y visibilidad
        'oa': {'titulo': 'OBJETIVOS DE APRENDIZAJE QUE SE ABORDARÁN EN EL MES:', 'prefijo': 'OA'},
        'clases': {
            'columnas': [
                {'clave': 'n', 'titulo': 'N° de clases', 'ancho': 1150},
                {'clave': 'objetivo', 'titulo': 'OBJETIVO DE LA CLASE', 'ancho': 2700},
                {'clave': 'actividad', 'titulo': 'ESTRATEGIA DE ACTIVIDAD', 'ancho': 5400},
                {'clave': 'evaluacion', 'titulo': 'ESTRATEGIA DE EVALUACIÓN', 'ancho': 3950},
                {'clave': 'recursos', 'titulo': 'RECURSOS O MATERIALES', 'ancho': 3200},
            ],
            'franja_entre_periodos': True,
            'alto_fila': 760,
            'alternar_color': True,
        },
        'indicadores': {'titulo': 'Indicadores de Evaluación:'},
        'pie': 'Planificación {meses} {anio}  ·  {asignatura}  ·  {curso}',
    },
    'anual': {
        'nombre': 'Predeterminado (placeholder)',
        'fuente': 'Arial',
        'colores': dict(COLORES),
        'pagina': dict(PAGINA),
        'titulo': 'PLAN ANUAL {anio} · PRIORIZACIÓN CURRICULAR',
        'mostrar_logo': True,
        'institucion': ['Complejo Educacional Particular', '', 'UTP {anio}'],
        'datos': [['Sector o Subsector', '{asignatura}'], ['Profesor/a', '{profesor}'], ['Curso', '{curso}'],
                  ['Año', '{anio}'], ['Nº de Horas', '{horas}'], ['Nº de Clases', '{clases}']],
        'semestres': [
            {'nombre': '1er SEMESTRE', 'etiqueta': 'PRIMER SEMESTRE', 'meses': ['MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO'],
             'semanas': [5, 5, 5, 5, 4], 'unidades': [1, 2]},
            {'nombre': '2do SEMESTRE', 'etiqueta': 'SEGUNDO SEMESTRE', 'meses': ['AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'],
             'semanas': [5, 5, 5, 5, 4], 'unidades': [3, 4]},
        ],
        'tabla': {'titulo_oa': 'OBJETIVOS DE APRENDIZAJE', 'titulo_indicadores': 'INDICADORES DE EVALUACIÓN', 'marca': 'X',
                  'ancho_oa': 3510, 'ancho_indicadores': 3855, 'ancho_semana': 425, 'filas_en_blanco': 4, 'prefijo_unidad': 'Unidad N°'},
        'leyenda': 'X = semana en que se aborda el objetivo. Ajustar el número de semanas de cada mes según el calendario escolar.',
        'cobertura': ['OA Enseñados', 'OA No Enseñados', '% Cobertura Curricular'],      # [] = no mostrar el cuadro
        'pie': 'Plan Anual {anio}  ·  {asignatura}  ·  {curso}',
    },
    'temario': {
        'nombre': 'Predeterminado (placeholder)',
        'fuente': 'Arial',
        'colores': dict(COLORES),
        'pagina': {'ancho_cm': 21.6, 'alto_cm': 27.9, 'margen_lateral_cm': 2.0, 'margen_superior_cm': 1.5, 'margen_inferior_cm': 1.5},
        'titulo': 'TEMARIO DE EVALUACIÓN',
        'subtitulo': '{asignatura}  |  {curso}',
        'mostrar_logo': True,
        'institucion': ['Complejo Educacional Particular', '', ''],
        'datos': [['Curso', '{curso}'], ['Fecha de evaluación', '{fecha}'], ['Puntaje total', '{puntaje}']],
        # Secciones libres: tipo "lista" (un elemento por linea) o "texto"
        'secciones': [
            {'titulo': 'Contenidos', 'tipo': 'lista', 'contenido': '{contenidos}'},
            {'titulo': 'Cómo será la evaluación', 'tipo': 'texto', 'contenido': '{como_sera}'},
        ],
        'cierre': 'Saludos cordiales,\n{profesor}',
        'pie': '{asignatura} · {curso}',
    },
}

DEFECTO['evaluacion'] = {
    'nombre': 'Predeterminado (placeholder)',
    'mostrar_oa': True,                 # mostrar "Objetivo de Aprendizaje" en la prueba del estudiante
    'mostrar_contenidos': True,
    'exigencia_pct': 0,                 # p. ej. 60 -> muestra "Exigencia: 60% (puntaje minimo de aprobacion: ...)"; 0 = no mostrar
    'ponderacion': {},                  # p. ej. {"item1": 40, "item2": 40, "item3": 20} = % del puntaje total de cada item
    'etiquetas': {},                    # p. ej. {"nombre": "Alumno/a:", "curso": "Curso:", "fecha": "Fecha:", ...}
    'instrucciones': '',                # vacio = instrucciones por defecto del motor
}

# claves cuyo contenido es una lista/estructura propia: se reemplazan completas, no se mezclan elemento a elemento
_REEMPLAZAR = {'columnas', 'datos', 'bloques', 'institucion', 'semestres', 'secciones', 'cobertura', 'ponderacion', 'etiquetas'}


def _mezclar(base, extra):
    out = copy.deepcopy(base)
    for k, v in (extra or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict) and k not in _REEMPLAZAR:
            out[k] = _mezclar(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


def slug(nombre):
    s = re.sub(r'[^a-z0-9]+', '_', (nombre or '').lower().encode('ascii', 'ignore').decode()).strip('_')
    return s or 'perfil'


def _ruta(tipo, nombre):
    return os.path.join(CARPETA, f'{tipo}.{slug(nombre)}.json')


def defecto(tipo):
    return copy.deepcopy(DEFECTO[tipo])


def listar(tipo):
    """Nombres de perfiles guardados (siempre incluye 'defecto')."""
    nombres = ['defecto']
    if os.path.isdir(CARPETA):
        for f in sorted(os.listdir(CARPETA)):
            if f.startswith(tipo + '.') and f.endswith('.json'):
                nombres.append(f[len(tipo) + 1:-5])
    return nombres


def cargar(tipo, perfil=None, extra=None):
    """Formato final para un documento: defecto + perfil guardado + ajustes del propio plan."""
    if tipo not in TIPOS:
        raise ValueError(f'tipo de formato desconocido: {tipo}')
    fmt = defecto(tipo)
    if perfil and perfil != 'defecto':
        ruta = _ruta(tipo, perfil)
        if not os.path.exists(ruta):
            raise FileNotFoundError(f'No existe el formato "{perfil}" ({tipo})')
        with open(ruta, encoding='utf-8') as f:
            fmt = _mezclar(fmt, json.load(f))
    if extra:
        fmt = _mezclar(fmt, extra)
    return limpiar(tipo, fmt)


def guardar(tipo, nombre, fmt):
    """Guarda un perfil (validado). Devuelve el nombre normalizado."""
    fmt = limpiar(tipo, _mezclar(defecto(tipo), fmt))
    fmt['nombre'] = str(nombre)[:60]
    os.makedirs(CARPETA, exist_ok=True)
    with open(_ruta(tipo, nombre), 'w', encoding='utf-8') as f:
        json.dump(fmt, f, ensure_ascii=False, indent=1)
    return slug(nombre)


def restablecer(tipo, nombre):
    """Borra un perfil guardado (vuelve al formato por defecto). No toca el formato incluido."""
    if nombre in (None, '', 'defecto'):
        return False
    ruta = _ruta(tipo, nombre)
    if os.path.exists(ruta):
        os.remove(ruta)
        return True
    return False


# ---------------------------------------------------------------- validacion (los formatos pueden venir de la web)
def _hex(v, por_defecto):
    v = str(v or '').lstrip('#').upper()
    return v if re.fullmatch(r'[0-9A-F]{6}', v) else por_defecto


def _num(v, minimo, maximo, por_defecto):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return por_defecto
    return min(max(v, minimo), maximo)


def _txt(v, largo=200):
    return str(v if v is not None else '')[:largo]


def limpiar(tipo, fmt):
    f = fmt
    f['nombre'] = _txt(f.get('nombre'), 60)
    if tipo == 'evaluacion':
        f['mostrar_oa'] = bool(f.get('mostrar_oa', True))
        f['mostrar_contenidos'] = bool(f.get('mostrar_contenidos', True))
        f['exigencia_pct'] = _num(f.get('exigencia_pct'), 0, 100, 0)
        pond = f.get('ponderacion') or {}
        f['ponderacion'] = {k: _num(pond.get(k), 0, 100, 0) for k in ('item1', 'item2', 'item3') if k in pond and pond.get(k) not in (None, '')}
        etiquetas = f.get('etiquetas') or {}
        f['etiquetas'] = {k: _txt(etiquetas[k], 60) for k in ('nombre', 'curso', 'fecha', 'puntaje_total', 'puntaje_obtenido') if etiquetas.get(k)}
        f['instrucciones'] = _txt(f.get('instrucciones'), 600)
        return f
    f['fuente'] = f.get('fuente') if f.get('fuente') in ('Arial', 'Calibri', 'Times New Roman', 'Verdana', 'Tahoma', 'Georgia', 'Cambria', 'Segoe UI') else 'Arial'
    base = DEFECTO[tipo]
    f['colores'] = {k: _hex(f.get('colores', {}).get(k), base['colores'][k]) for k in base['colores']}
    pg = f.get('pagina', {})
    f['pagina'] = {k: _num(pg.get(k), 0.3 if 'margen' in k else 10, 12 if 'margen' in k else 100, base['pagina'][k]) for k in base['pagina']}
    for k in ('titulo', 'subtitulo', 'pie', 'leyenda', 'cierre'):
        if k in f:
            f[k] = _txt(f[k], 300)
    f['mostrar_logo'] = bool(f.get('mostrar_logo', True))
    f['institucion'] = [_txt(x, 120) for x in (f.get('institucion') or [])][:4]
    if tipo == 'mensual':
        cols, vistos = [], set()
        for c in (f['clases'].get('columnas') or [])[:8]:
            clave = re.sub(r'[^a-z0-9_]', '', str(c.get('clave', '')).lower())[:20]
            if not clave or clave in vistos:
                continue
            vistos.add(clave)
            cols.append({'clave': clave, 'titulo': _txt(c.get('titulo'), 60), 'ancho': int(_num(c.get('ancho'), 500, 9000, 2500))})
        if 'n' not in vistos:
            cols.insert(0, {'clave': 'n', 'titulo': 'N° de clases', 'ancho': 1150})
        f['clases']['columnas'] = cols
        f['clases']['alto_fila'] = int(_num(f['clases'].get('alto_fila'), 300, 4000, 760))
        f['bloques'] = [b for b in f.get('bloques', []) if b in ('datos', 'oa', 'clases', 'indicadores')] or ['clases']
        f['datos'] = [[[_txt(p[0], 80), _txt(p[1], 120)] for p in fila if len(p) == 2] for fila in (f.get('datos') or [])[:4]]
        f['oa']['titulo'] = _txt(f['oa'].get('titulo'), 200); f['oa']['prefijo'] = _txt(f['oa'].get('prefijo'), 10)
        f['indicadores']['titulo'] = _txt(f['indicadores'].get('titulo'), 120)
    if tipo == 'anual':
        sems = []
        for s in (f.get('semestres') or [])[:4]:
            meses = [_txt(m, 20).upper() for m in (s.get('meses') or [])][:8]
            sem = [int(_num(x, 1, 8, 5)) for x in (s.get('semanas') or [])][:len(meses)]
            sem += [5] * (len(meses) - len(sem))
            sems.append({'nombre': _txt(s.get('nombre'), 40), 'etiqueta': _txt(s.get('etiqueta'), 40), 'meses': meses,
                         'semanas': sem, 'unidades': [int(_num(u, 1, 20, 1)) for u in (s.get('unidades') or [])][:6]})
        f['semestres'] = sems or defecto('anual')['semestres']
        t = f['tabla']
        for k in ('ancho_oa', 'ancho_indicadores', 'ancho_semana'):
            t[k] = int(_num(t.get(k), 200, 8000, base['tabla'][k]))
        t['filas_en_blanco'] = int(_num(t.get('filas_en_blanco'), 1, 12, 4))
        t['marca'] = _txt(t.get('marca'), 3) or 'X'
        for k in ('titulo_oa', 'titulo_indicadores', 'prefijo_unidad'):
            t[k] = _txt(t.get(k), 80)
        f['datos'] = [[_txt(p[0], 60), _txt(p[1], 120)] for p in (f.get('datos') or [])[:10] if len(p) == 2]
        f['cobertura'] = [_txt(x, 60) for x in (f.get('cobertura') or [])][:4]
    if tipo == 'temario':
        f['datos'] = [[_txt(p[0], 60), _txt(p[1], 120)] for p in (f.get('datos') or [])[:10] if len(p) == 2]
        f['secciones'] = [{'titulo': _txt(s.get('titulo'), 80), 'tipo': 'lista' if s.get('tipo') == 'lista' else 'texto',
                           'contenido': _txt(s.get('contenido'), 4000)} for s in (f.get('secciones') or [])[:8]]
    return f


def rellenar(texto, valores):
    """Sustituye {campo} por su valor; los campos desconocidos quedan vacios."""
    return re.sub(r'\{(\w+)\}', lambda m: str(valores.get(m.group(1), '')), texto or '')
