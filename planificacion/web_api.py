# -*- coding: utf-8 -*-
"""Funciones de la API web del modulo de planificaciones (las usa app.py). Todo lo que llega del navegador se sanea aqui."""
import copy
import glob
import json
import os
import re
import threading

from . import anual, banco, formato, ia, mensual, temario

RAIZ = os.path.dirname(os.path.abspath(__file__))
BLOQUEO = threading.Lock()          # docx_base usa un tema global (colores/fuente): una construccion a la vez
PROVEEDORES_IA = ('offline', 'gemini', 'claude', 'openai', 'deepseek', 'kimi', 'grok', 'local')


def _s(v, n=300):
    return str(v if v is not None else '')[:n]


def _n(v, minimo, maximo, defecto):
    try:
        return int(min(max(float(v), minimo), maximo))
    except (TypeError, ValueError):
        return defecto


def catalogo():
    for ruta in (os.path.join(RAIZ, 'datos', 'colegio.json'), os.path.join(RAIZ, 'ejemplos', 'colegio.ejemplo.json')):
        if os.path.exists(ruta):
            with open(ruta, encoding='utf-8') as f:
                return json.load(f)
    return {'profesor': '', 'colegio': {}, 'cursos': {}}


# ------------------------------------------------------------------ lectura para el panel
def opciones():
    cat = catalogo()
    bancos = []
    for ruta in sorted(glob.glob(os.path.join(banco.CARPETA_BANCO, '*', '*.json'))):
        try:
            with open(ruta, encoding='utf-8') as f:
                b = json.load(f)
        except (OSError, ValueError):
            continue
        bancos.append({'anio': b.get('anio'), 'curso_codigo': b.get('curso_codigo', b.get('curso')), 'asignatura': b.get('asignatura'),
                       'oa': [{'codigo': o.get('codigo'), 'unidad': o.get('unidad'), 'semestre': o.get('semestre'), 'texto': o['texto'][:160]}
                              for o in b.get('oa', [])]})
    return {
        'profesor': cat.get('profesor', ''), 'cursos': cat.get('cursos', {}), 'bancos': bancos,
        'formatos': {t: formato.listar(t) for t in formato.TIPOS},
        'defecto': {t: formato.defecto(t) for t in formato.TIPOS},
        'proveedores_ia': list(PROVEEDORES_IA),
    }


def ver_formato(d):
    tipo = d.get('tipo')
    if tipo not in formato.TIPOS:
        raise ValueError('tipo de formato inválido')
    inline = d.get('formato') if isinstance(d.get('formato'), dict) and len(json.dumps(d['formato'])) <= 40000 else None
    return formato.cargar(tipo, _s(d.get('perfil'), 60) or None, inline)


def guardar_formato(d):
    tipo = d.get('tipo')
    if tipo not in formato.TIPOS or not _s(d.get('nombre'), 60).strip():
        raise ValueError('tipo o nombre de formato inválido')
    if len(json.dumps(d.get('formato', {}))) > 40000:
        raise ValueError('formato demasiado grande')
    return formato.guardar(tipo, _s(d['nombre'], 60), d.get('formato') or {})


def restablecer_formato(d):
    tipo = d.get('tipo')
    if tipo not in formato.TIPOS:
        raise ValueError('tipo de formato inválido')
    return formato.restablecer(tipo, _s(d.get('nombre'), 60))


def oa_banco(d):
    """Busca en el banco curricular los OA indicados (por numero) y devuelve texto + indicadores editables."""
    codigos = [_s(c, 10).strip() for c in (d.get('codigos') or [])][:12]
    curso, asig, anio = _s(d.get('curso_codigo'), 40), _s(d.get('asignatura'), 80), _n(d.get('anio'), 2000, 2100, 2026)
    unidad = _s(d.get('unidad'), 10) or None
    try:
        b = banco.cargar(curso, asig, anio)
    except FileNotFoundError:
        b = None
    out = []
    for c in codigos:
        try:
            if b is None:
                raise KeyError(c)
            o = banco.buscar_oa(b, c, unidad)
            out.append({'codigo': c, 'texto': o['texto'], 'puntos': o.get('puntos', []), 'indicadores': o.get('indicadores', []), 'encontrado': True})
        except KeyError:
            out.append({'codigo': c, 'texto': '', 'puntos': [], 'indicadores': [], 'encontrado': False})
    return {'oa': out, 'hay_banco': b is not None}


# ------------------------------------------------------------------ saneamiento de lo que envia el navegador
def _clase(c):
    out = {}
    for k, v in (c or {}).items():
        k = re.sub(r'[^a-z0-9_]', '', str(k).lower())[:20]
        if k == 'sugerencia' and isinstance(v, dict):
            out['sugerencia'] = {re.sub(r'[^a-z0-9_]', '', str(a).lower())[:20]: _s(b, 3000) for a, b in v.items()}
        elif k:
            out[k] = _s(v, 3000)
    return out


def limpiar_plan(p, profesor_sesion=''):
    cat = catalogo()
    plan = {
        'anio': _n(p.get('anio'), 2000, 2100, 2026), 'profesor': _s(p.get('profesor'), 120) or profesor_sesion or cat.get('profesor', ''),
        'curso': _s(p.get('curso'), 60), 'curso_codigo': _s(p.get('curso_codigo') or p.get('curso'), 40),
        'asignatura': _s(p.get('asignatura'), 80), 'horas': _s(p.get('horas'), 80), 'clases': _s(p.get('clases'), 80),
        'unidad': {'numero': _s((p.get('unidad') or {}).get('numero'), 10), 'nombre': _s((p.get('unidad') or {}).get('nombre'), 150)},
        'colegio': cat.get('colegio', {}),                    # lo define el servidor, no el navegador
        'evaluacion_por_defecto': _s(p.get('evaluacion_por_defecto'), 300),
    }
    if p.get('formato_perfil'):
        plan['formato_perfil'] = formato.slug(_s(p['formato_perfil'], 60))
    if isinstance(p.get('formato'), dict) and len(json.dumps(p['formato'])) <= 40000:
        plan['formato'] = p['formato']
    oas = []
    for o in (p.get('oa') or [])[:12]:
        if isinstance(o, dict):
            oas.append({'codigo': _s(o.get('codigo'), 10), 'texto': _s(o.get('texto'), 1500),
                        'puntos': [_s(x, 400) for x in (o.get('puntos') or [])][:8],
                        'indicadores': [_s(x, 500) for x in (o.get('indicadores') or [])][:30]})
        else:
            oas.append(_s(o, 10))
    plan['oa'] = oas
    plan['periodos'] = [{'nombre': _s(q.get('nombre'), 30), 'clases': [_clase(c) for c in (q.get('clases') or [])][:40]}
                        for q in (p.get('periodos') or [])][:6]
    if not plan['periodos'] or not plan['asignatura']:
        raise ValueError('Faltan datos: asignatura y al menos un período')
    return plan


def limpiar_anual(p, profesor_sesion=''):
    cat = catalogo()
    plan = {'anio': _n(p.get('anio'), 2000, 2100, 2027), 'profesor': _s(p.get('profesor'), 120) or profesor_sesion or cat.get('profesor', ''),
            'curso': _s(p.get('curso'), 60), 'curso_codigo': _s(p.get('curso_codigo') or p.get('curso'), 40),
            'asignatura': _s(p.get('asignatura'), 80), 'horas': _s(p.get('horas'), 80), 'clases': _s(p.get('clases'), 80),
            'colegio': cat.get('colegio', {}), 'oa': []}
    if p.get('formato_perfil'):
        plan['formato_perfil'] = formato.slug(_s(p['formato_perfil'], 60))
    if isinstance(p.get('formato'), dict) and len(json.dumps(p['formato'])) <= 40000:
        plan['formato'] = p['formato']
    if not plan['asignatura']:
        raise ValueError('Falta la asignatura')
    if p.get('desde_banco'):                                  # plan anual completo desde el banco curricular
        b = banco.cargar(plan['curso_codigo'], plan['asignatura'], _n(p.get('anio_banco'), 2000, 2100, plan['anio']))
        plan['oa'] = b['oa']; plan['meses'] = b.get('meses', {})
        for k in ('horas', 'clases'):
            plan[k] = plan[k] or b.get(k, '')
    return plan


def limpiar_temario(p):
    cat = catalogo()
    cont = p.get('contenidos', '')
    if isinstance(cont, list):
        cont = '\n'.join(_s(x, 300) for x in cont[:40])
    d = {k: _s(p.get(k), 200) for k in ('curso', 'asignatura', 'profesor', 'fecha', 'puntaje')}
    d.update({'contenidos': _s(cont, 4000), 'como_sera': _s(p.get('como_sera'), 2000), 'colegio': cat.get('colegio', {}),
              'anio': _n(p.get('anio'), 2000, 2100, 2026)})
    if p.get('formato_perfil'):
        d['formato_perfil'] = formato.slug(_s(p['formato_perfil'], 60))
    if isinstance(p.get('formato'), dict) and len(json.dumps(p['formato'])) <= 40000:
        d['formato'] = p['formato']
    return d


def limpiar_formato_evaluacion(d, perfil=None):
    """Formato de la evaluacion que llega del panel -> dict validado para engine.DocenteEngine (cfg['formato'])."""
    inline = d if isinstance(d, dict) else {}
    return formato.cargar('evaluacion', formato.slug(perfil) if perfil else None, inline)


# ------------------------------------------------------------------ acciones
def sugerir(d):
    motor = d.get('motor', 'offline')
    if motor not in PROVEEDORES_IA:
        raise ValueError('motor de sugerencias inválido')
    plan = limpiar_plan(d.get('plan') or {})
    n = ia.sugerir(plan, motor, _s(d.get('modelo'), 60) or None, _s(d.get('api_key'), 200) or None)
    return {'plan': plan, 'sugeridas': n}


def generar(tipo, d, ruta, profesor_sesion=''):
    """Genera el .docx pedido en `ruta` (bajo bloqueo). tipo: mensual | anual | temario."""
    with BLOQUEO:
        if tipo == 'mensual':
            return mensual.construir(limpiar_plan(d.get('plan') or {}, profesor_sesion), ruta, incluir_sugerencias=bool(d.get('sugerencias')))
        if tipo == 'anual':
            return anual.construir(limpiar_anual(d.get('plan') or d, profesor_sesion), ruta)
        if tipo == 'temario':
            return temario.construir(limpiar_temario(d.get('temario') or d), ruta)
    raise ValueError('tipo de documento inválido')
