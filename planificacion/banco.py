# -*- coding: utf-8 -*-
"""Banco curricular: OA e indicadores por curso/asignatura, extraidos de los Planes Anuales (.docx) a JSON.

El banco permite que una planificacion mensual solo indique los numeros de OA ("8, 9, 10, 11"):
el texto del objetivo y sus indicadores salen de aqui.
"""
import json
import os
import re
import unicodedata

from docx import Document
from docx.oxml.ns import qn

RAIZ = os.path.dirname(os.path.abspath(__file__))
CARPETA_BANCO = os.path.join(RAIZ, 'datos', 'banco')

MESES = ['MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
ARREGLOS = (('Identi car', 'Identificar'), ('seg- n', 'según'), ('Eval- an', 'Evalúan'))


def slug(texto):
    t = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode()
    return re.sub(r'[^a-z0-9]+', '_', t.lower()).strip('_')


def clave(curso, asignatura):
    """'5°AB' + 'Ciencias Naturales' -> '5ab__ciencias_naturales'"""
    return f'{slug(curso)}__{slug(asignatura)}'


def _tx(el):
    return ''.join(t.text or '' for t in el.iter(qn('w:t'))).strip()


def _parrafos(tc):
    return [''.join(t.text or '' for t in p.iter(qn('w:t'))).strip() for p in tc.iter(qn('w:p'))]


def limpiar(t):
    for a, b in ARREGLOS:
        t = t.replace(a, b)
    return re.sub(r'\s+', ' ', t).strip()


def separar_indicadores(parrafos):
    """Une lineas cortadas a mitad de frase y devuelve una lista: un indicador por elemento."""
    paras = [limpiar(t) for t in parrafos if t.strip()]
    marcados = any(t.startswith(('-', '›', '•')) for t in paras)
    if not marcados:
        out = []
        for t in paras:
            for x in re.split(r'\s*›\s*', t):
                x = x.strip(' -•\t')
                if not x:
                    continue
                if out and not re.match(r'^[A-ZÁÉÍÓÚÑ¿¡0-9«"(]', x):
                    out[-1] += ' ' + x
                else:
                    out.append(x)
        return out
    junto = ' '.join(paras)
    partes = re.split(r'(?:^|\s)(?:-|›|•)\s+(?=\S)', junto)
    return [x.strip(' -•›\t') for x in partes if x.strip(' -•›\t')]


def separar_objetivo(parrafos):
    """-> (codigo | None, texto principal, [sub-puntos])"""
    oa = limpiar(' '.join(p for p in parrafos if p.strip()))
    partes = [x.strip() for x in re.split(r'\s(?:-|>|›|–)\s+(?=\S)', oa) if x.strip()]
    principal, puntos = partes[0], partes[1:]
    cola = re.compile(r'\s*\(OA\s*(\d+)\)\s*\.?\s*$')
    codigo_cola = None
    if puntos and cola.search(puntos[-1]):
        codigo_cola = cola.search(puntos[-1]).group(1)
        puntos[-1] = cola.sub('', puntos[-1]).strip()
    codigo = None
    m = re.match(r'^\s*\(?\s*OA\s*(\d+)\s*\)?\s*\*?\s*[:.]?\s*', principal)
    if m:
        codigo = m.group(1); principal = principal[m.end():]
    else:
        m = re.search(r'\s*\(OA\s*(\d+)\)\s*\.?\s*$', principal)
        if m:
            codigo = m.group(1); principal = principal[:m.start()]
    return codigo or codigo_cola, principal.strip(), puntos


def importar_plan_anual(ruta_docx, curso, asignatura, anio, codigos_manuales=None):
    """Lee un Plan Anual (.docx, formato del colegio) y devuelve su contenido como dict (ver README)."""
    codigos_manuales = codigos_manuales or {}
    d = Document(ruta_docx)
    oas = []
    meses_semestre = {}
    for tbl in d.element.body.iter(qn('w:tbl')):
        filas = list(tbl.iter(qn('w:tr')))
        if len(filas) < 2:
            continue
        r0 = [_tx(tc) for tc in filas[0].iter(qn('w:tc'))]
        if not r0 or not r0[0].upper().startswith('OBJETIVOS'):
            continue
        semestre = unidad = None
        semanas = []      # por columna de semana: (mes, n)
        for fila in filas:
            tcs = list(fila.iter(qn('w:tc')))
            t0 = _tx(tcs[0]).upper()
            if t0.startswith('OBJETIVOS'):
                nombres = [_tx(tc).upper() for tc in tcs[2:]]
                nombres = [n for n in nombres if n in MESES]
                semestre = None
                continue
            if t0.startswith(('PRIMER', 'SEGUNDO')) and len(tcs) >= 10:
                semestre = 1 if t0.startswith('PRIMER') else 2
                m = re.search(r'Unidad\s*N?[°º]?\s*(\d+)', _tx(tcs[0]), re.I)
                unidad = int(m.group(1)) if m else None
                semanas, mes_i = [], -1
                for tc in tcs[2:]:
                    s = _tx(tc)
                    if s == '1':
                        mes_i += 1
                    if s.isdigit() and 0 <= mes_i < len(nombres):
                        semanas.append((nombres[mes_i], int(s)))
                    else:
                        semanas.append(None)
                meses_semestre[semestre] = list(dict.fromkeys(n for n, _ in (x for x in semanas if x)))
                continue
            if len(tcs) < 10 or semestre is None or not _tx(tcs[0]) or not _tx(tcs[1]):
                continue
            codigo, texto, puntos = separar_objetivo(_parrafos(tcs[0]))
            for inicio, cod in codigos_manuales.items():
                if codigo is None and texto.lower().startswith(inicio.lower()):
                    codigo = cod
            marcas = {}
            for k, tc in enumerate(tcs[2:]):
                if _tx(tc).upper() == 'X' and k < len(semanas) and semanas[k]:
                    mes, n = semanas[k]
                    marcas.setdefault(mes, []).append(n)
            oas.append({'semestre': semestre, 'unidad': unidad, 'codigo': codigo, 'texto': texto, 'puntos': puntos,
                        'indicadores': separar_indicadores(_parrafos(tcs[1])), 'marcas': marcas})
    info = {}
    for tbl in d.element.body.iter(qn('w:tbl')):
        if _tx(tbl).startswith('Sector o Subsector'):
            for tr in tbl.iter(qn('w:tr')):
                c = [_tx(tc) for tc in tr.iter(qn('w:tc'))]
                if len(c) >= 2:
                    info[c[0]] = c[1]
            break
    return {'anio': anio, 'curso': curso, 'asignatura': asignatura,
            'horas': info.get('Nº de Horas', ''), 'clases': info.get('Nº de Clases', ''),
            'meses': meses_semestre, 'oa': oas}


def guardar(banco, anio=None):
    anio = anio or banco['anio']
    carpeta = os.path.join(CARPETA_BANCO, str(anio))
    os.makedirs(carpeta, exist_ok=True)
    ruta = os.path.join(carpeta, clave(banco.get('curso_codigo', banco['curso']), banco['asignatura']) + '.json')
    with open(ruta, 'w', encoding='utf-8') as f:
        json.dump(banco, f, ensure_ascii=False, indent=1)
    return ruta


def cargar(curso, asignatura, anio):
    ruta = os.path.join(CARPETA_BANCO, str(anio), clave(curso, asignatura) + '.json')
    if not os.path.exists(ruta):
        raise FileNotFoundError(f'No hay banco curricular {anio} para {curso} {asignatura}: {ruta}\n'
                                f'Importarlo con: python -m planificacion banco-importar <plan_anual.docx> ...')
    with open(ruta, encoding='utf-8') as f:
        return json.load(f)


def buscar_oa(banco, codigo, unidad=None):
    """OA por numero. Prefiere la unidad indicada; si no esta ahi, usa la ultima unidad donde aparece."""
    codigo = str(codigo)
    cand = [o for o in banco['oa'] if o.get('codigo') == codigo]
    if not cand:
        raise KeyError(f'OA {codigo} no esta en el banco de {banco["curso"]} {banco["asignatura"]} {banco["anio"]}')
    if unidad is not None:
        exactos = [o for o in cand if o.get('unidad') == int(unidad)]
        if exactos:
            return exactos[-1]
    return cand[-1]
