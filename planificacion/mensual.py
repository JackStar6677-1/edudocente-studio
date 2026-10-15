# -*- coding: utf-8 -*-
"""Planificacion mensual (o de varios meses): JSON -> .docx segun un FORMATO editable (ver formato.py).

El docente solo escribe, por cada clase, los campos de las columnas del formato (por defecto: objetivo, actividad,
evaluacion y recursos). Los OA y sus indicadores se toman del banco curricular segun los numeros indicados.
"""
import json
import os

from . import banco as B
from . import docx_base as X
from . import formato as F

RAIZ = os.path.dirname(os.path.abspath(__file__))
CAMPOS = ('objetivo', 'actividad', 'evaluacion', 'recursos')     # columnas del formato por defecto


def cargar_json(ruta):
    with open(ruta, encoding='utf-8') as f:
        return json.load(f)


def campos_de(fmt):
    """Claves de dato por clase segun las columnas del formato (sin la numeracion)."""
    return [c['clave'] for c in fmt['clases']['columnas'] if c['clave'] != 'n']


def clases_vacias(n, fmt=None):
    claves = campos_de(fmt) if fmt else CAMPOS
    return [{c: '' for c in claves} for _ in range(n)]


def resolver_oa(plan):
    """Lista de OA ya con texto e indicadores (del banco, o escritos a mano en el JSON)."""
    anio = plan.get('anio_banco', plan['anio'])
    banco = None
    out = []
    unidad = plan['unidad'].get('numero')
    for item in plan['oa']:
        if isinstance(item, dict):                      # OA escrito a mano en el JSON
            out.append({'codigo': str(item['codigo']), 'texto': item['texto'],
                        'puntos': item.get('puntos', []), 'indicadores': item.get('indicadores', [])})
            continue
        if banco is None:
            banco = B.cargar(plan.get('banco_curso', plan['curso_codigo']), plan['asignatura'], anio)
        out.append(B.buscar_oa(banco, item, unidad))
    return out


def valores_plan(plan):
    meses = ' - '.join(p['nombre'] for p in plan['periodos'])
    return {'profesor': plan.get('profesor', ''), 'curso': plan.get('curso', ''), 'asignatura': plan.get('asignatura', ''),
            'meses': meses, 'anio': plan.get('anio', ''), 'unidad_numero': plan['unidad'].get('numero', ''),
            'unidad_nombre': plan['unidad'].get('nombre', ''), 'horas': plan.get('horas', '')}


def formato_de(plan):
    return F.cargar('mensual', plan.get('formato_perfil'), plan.get('formato'))


def construir(plan, destino, incluir_sugerencias=False):
    fmt = formato_de(plan)
    X.usar_tema(fmt['colores'], fmt['fuente'])
    pg = fmt['pagina']
    doc = X.nuevo_documento(pg['ancho_cm'], pg['alto_cm'], pg['margen_lateral_cm'], pg['margen_superior_cm'], pg['margen_inferior_cm'])
    ancho = int((pg['ancho_cm'] - 2 * pg['margen_lateral_cm']) * 567)
    v = valores_plan(plan)
    oas = resolver_oa(plan) if ('oa' in fmt['bloques'] or 'indicadores' in fmt['bloques']) else []
    colegio = plan.get('colegio', {})
    logo = (colegio.get('logo') or X.LOGO_DEFECTO) if fmt['mostrar_logo'] else None

    # ---- franja de titulo
    lineas = [F.rellenar(t, v) for t in fmt['institucion']] or ['']
    for i, k in enumerate(('linea1', 'linea2')):               # la configuracion del colegio completa lo que el formato deja vacio
        if i < len(lineas) and not lineas[i] and colegio.get(k):
            lineas[i] = colegio[k]
    estilos = [(10, X.NAVY, True), (10, X.NAVY, True), (8.5, '595959', False), (8.5, '595959', False)]
    X.franja_titulo(doc, logo, [(t, estilos[i][0], estilos[i][1], estilos[i][2]) for i, t in enumerate(lineas) if t],
                    F.rellenar(fmt['titulo'], v), F.rellenar(fmt.get('subtitulo', ''), v).upper(), ancho)

    for bloque in fmt['bloques']:
        if bloque == 'datos':
            _bloque_datos(doc, fmt, v, ancho)
        elif bloque == 'oa':
            _bloque_oa(doc, fmt, oas, ancho)
        elif bloque == 'clases':
            _bloque_clases(doc, fmt, plan, incluir_sugerencias, ancho)
        elif bloque == 'indicadores':
            _bloque_indicadores(doc, fmt, oas, ancho)
    fin = X.espaciador(doc, 1); fin.runs[0].font.size = X.Pt(1)
    X.pie_de_pagina(doc, F.rellenar(fmt['pie'], v))
    os.makedirs(os.path.dirname(os.path.abspath(destino)), exist_ok=True)
    X.finalizar(doc, 'Planificación mensual')
    doc.save(destino)
    return destino


def _bloque_datos(doc, fmt, v, ancho):
    for fila in fmt['datos']:
        if not fila:
            continue
        n = len(fila)
        anchos = []                                       # cada par: ~1/3 etiqueta y ~2/3 valor
        for _ in fila:
            anchos += [int(ancho / n * 0.34), int(ancho / n * 0.66)]
        anchos[-1] += ancho - sum(anchos)
        t = X.tabla(doc, 1, anchos)
        for i, (etiqueta, valor) in enumerate(fila):
            c0, c1 = t.rows[0].cells[2 * i], t.rows[0].cells[2 * i + 1]
            X.sombrear(c0, X.ICE); X.parrafo(c0, F.rellenar(etiqueta, v), True, 9, X.NAVY); X.valign(c0)
            X.parrafo(c1, F.rellenar(valor, v), False, 10); X.valign(c1)
    X.espaciador(doc, 5)


def _bloque_oa(doc, fmt, oas, ancho):
    t = X.tabla(doc, 2, [ancho])
    X.sombrear(t.rows[0].cells[0], X.MID)
    X.parrafo(t.rows[0].cells[0], fmt['oa']['titulo'], True, 10, 'FFFFFF')
    cuerpo = t.rows[1].cells[0]; X.sombrear(cuerpo, X.PALE)
    pref = fmt['oa'].get('prefijo', 'OA')
    for i, oa in enumerate(oas):
        p = X.parrafo(cuerpo, '', first=(i == 0), after=3 if not oa['puntos'] else 1)
        X.run(p, f'{pref} {oa["codigo"]}: '.lstrip(), True, 10, X.NAVY)
        X.run(p, oa['texto'], False, 10)
        for k, pt in enumerate(oa['puntos']):
            X.parrafo(cuerpo, '–  ' + pt, False, 10, first=False, left=510, hanging=200, after=3 if k == len(oa['puntos']) - 1 else 0)
    X.fila_fija(t.rows[1])
    X.espaciador(doc, 6)


def _bloque_clases(doc, fmt, plan, incluir_sugerencias, ancho):
    cols = fmt['clases']['columnas']
    total = sum(c['ancho'] for c in cols)
    anchos = [int(c['ancho'] * ancho / total) for c in cols]
    anchos[-1] += ancho - sum(anchos)
    franja = fmt['clases']['franja_entre_periodos']
    n_filas = 1 + sum(len(p['clases']) for p in plan['periodos']) + (max(0, len(plan['periodos']) - 1) if franja else 0)
    t = X.tabla(doc, n_filas, anchos)
    for i, col in enumerate(cols):
        c = t.rows[0].cells[i]; X.sombrear(c, X.NAVY)
        X.parrafo(c, col['titulo'], True, 9.5, 'FFFFFF', X.CENTER if col['clave'] == 'n' else None); X.valign(c)
    X.fila_fija(t.rows[0], encabezado=True)
    ri = 1
    for k, periodo in enumerate(plan['periodos']):
        if k > 0 and franja:                              # franja del mes (DICIEMBRE, etc.)
            fila = t.rows[ri]; ri += 1
            m = fila.cells[0].merge(fila.cells[len(cols) - 1]); X.sombrear(m, X.MID)
            X.parrafo(m, periodo['nombre'].upper(), True, 10, 'FFFFFF', X.CENTER).paragraph_format.keep_with_next = True
            X.fila_fija(fila)
        for n, cl in enumerate(periodo['clases']):
            fila = t.rows[ri]; ri += 1
            relleno = X.PALE if (n % 2 and fmt['clases']['alternar_color']) else 'FFFFFF'
            sug = cl.get('sugerencia', {}) if incluir_sugerencias else {}
            for cel, col in zip(fila.cells, cols):
                X.sombrear(cel, relleno)
                if col['clave'] == 'n':
                    X.parrafo(cel, str(n + 1), True, 9.5, X.NAVY, X.CENTER); X.valign(cel)
                    continue
                txt = cl.get(col['clave'], '')
                if not txt and sug.get(col['clave']):
                    X.parrafo(cel, sug[col['clave']], False, 9, '7F7F7F', italic=True)   # sugerencia en gris cursiva
                else:
                    X.parrafo(cel, txt, False, 9.5)
                X.valign(cel, 'top')
            X.fila_fija(fila, alto_min=fmt['clases']['alto_fila'])
    X.espaciador(doc, 6)


def _bloque_indicadores(doc, fmt, oas, ancho):
    t = X.tabla(doc, 1 + len(oas), [1500, ancho - 1500])
    m = t.rows[0].cells[0].merge(t.rows[0].cells[1]); X.sombrear(m, X.MID)
    X.parrafo(m, fmt['indicadores']['titulo'], True, 10, 'FFFFFF').paragraph_format.keep_with_next = True
    X.fila_fija(t.rows[0], encabezado=True)
    pref = fmt['oa'].get('prefijo', 'OA')
    for i, oa in enumerate(oas):
        a, b = t.rows[1 + i].cells
        X.sombrear(a, X.ICE); X.parrafo(a, f'{pref} {oa["codigo"]}'.strip(), True, 10, X.NAVY, X.CENTER); X.valign(a)
        for j, ind in enumerate(oa['indicadores']):
            X.parrafo(b, '•  ' + ind, False, 9, first=(j == 0), after=2, left=255, hanging=255)
        if len(oa['indicadores']) <= 6:
            X.fila_fija(t.rows[1 + i])


def nombre_archivo(plan):
    meses = '-'.join(p['nombre'] for p in plan['periodos'])
    return f'{plan["curso_codigo"]} - {plan["asignatura"]} - Planificación {meses}.docx'
