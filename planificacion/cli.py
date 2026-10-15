# -*- coding: utf-8 -*-
"""Linea de comandos:  python -m planificacion <comando> ...   (ver planificacion/README.md)"""
import argparse
import json
import os
import sys

from . import anual, banco, formato, ia, leer, mensual, temario

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import documento_meta   # noqa: E402

RAIZ = os.path.dirname(os.path.abspath(__file__))
CATALOGO = os.path.join(RAIZ, 'datos', 'colegio.json')
CATALOGO_EJEMPLO = os.path.join(RAIZ, 'ejemplos', 'colegio.ejemplo.json')


def _catalogo():
    ruta = CATALOGO if os.path.exists(CATALOGO) else CATALOGO_EJEMPLO       # sin datos locales: configuracion generica
    with open(ruta, encoding='utf-8') as f:
        return json.load(f)


def _leer(ruta):
    with open(ruta, encoding='utf-8') as f:
        return json.load(f)


def _guardar(obj, ruta):
    os.makedirs(os.path.dirname(os.path.abspath(ruta)), exist_ok=True)
    with open(ruta, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    return ruta


def _datos_base(cat, curso_codigo):
    curso = cat['cursos'].get(curso_codigo, curso_codigo)
    return {'colegio': cat['colegio'], 'profesor': cat['profesor'], 'curso': curso, 'curso_codigo': curso_codigo}


# ------------------------------------------------------------------ mensual
def cmd_nuevo_mensual(a):
    cat = _catalogo()
    plan = _datos_base(cat, a.curso)
    if a.banco:
        plan['banco_curso'] = a.banco
    if a.formato:
        plan['formato_perfil'] = a.formato
    plan.update({'anio': a.anio, 'asignatura': a.asignatura,
                 'unidad': {'numero': a.unidad, 'nombre': a.nombre_unidad}, 'horas': a.horas or '',
                 'oa': [x.replace('OA', '').strip() for x in a.oa]})
    fmt = mensual.formato_de(plan)
    plan['periodos'] = []
    for item in a.mes:                                   # Noviembre:15  Diciembre:2
        nombre, n = item.split(':')
        plan['periodos'].append({'nombre': nombre, 'clases': mensual.clases_vacias(int(n), fmt)})
    plan['evaluacion_por_defecto'] = 'Evaluación formativa a través de las actividades realizadas en clases.'
    ruta = a.salida or mensual.nombre_archivo(plan).replace('.docx', '.json')
    print(_guardar(plan, ruta))


def cmd_generar_mensual(a):
    plan = _leer(a.plan)
    destino = a.salida or os.path.join(os.path.dirname(os.path.abspath(a.plan)), mensual.nombre_archivo(plan))
    print(mensual.construir(plan, destino, incluir_sugerencias=a.sugerencias))


def cmd_leer_mensual(a):
    fmt = formato.cargar('mensual', a.formato)
    plan = leer.leer_mensual(a.docx, a.curso, a.anio, fmt)
    cat = _catalogo(); plan['colegio'] = cat['colegio']
    if a.formato:
        plan['formato_perfil'] = a.formato
    print(_guardar(plan, a.salida or os.path.splitext(a.docx)[0] + '.json'))


def cmd_estado(a):
    """Resumen compacto (pocas lineas) de que clases estan escritas y cuales faltan."""
    plan = _leer(a.plan)
    claves = mensual.campos_de(mensual.formato_de(plan))
    print(f'{plan["curso"]} · {plan["asignatura"]} · unidad {plan["unidad"]["numero"]} «{plan["unidad"]["nombre"]}» · OA {",".join(map(str, [o if not isinstance(o, dict) else o["codigo"] for o in plan["oa"]]))}')
    for p in plan['periodos']:
        marcas = []
        for n, c in enumerate(p['clases']):
            llenos = sum(bool(str(c.get(k, '')).strip()) for k in claves)
            marcas.append(f'{n + 1}{"✓" if llenos == len(claves) else ("~" if llenos else "·")}')
        print(f'  {p["nombre"]}: ' + ' '.join(marcas))
    print('  (✓ completa · ~ parcial · · vacía)')


def cmd_semana(a):
    plan = _leer(a.plan)
    periodo = next(p for p in plan['periodos'] if p['nombre'].lower() == a.periodo.lower())
    clase = periodo['clases'][a.n - 1]
    for campo in mensual.CAMPOS:
        valor = getattr(a, campo)
        if valor is not None:
            clase[campo] = valor
    for par in (a.campo or []):                          # columnas personalizadas: --campo clave=valor
        k, _, v = par.partition('=')
        clase[k.strip()] = v
    _guardar(plan, a.plan)
    print(f'{a.periodo} {a.n} actualizada')


def cmd_sugerir(a):
    plan = _leer(a.plan)
    n = ia.sugerir(plan, a.ia, a.modelo)
    _guardar(plan, a.plan)
    print(f'{n} clases con sugerencia (motor: {a.ia}). Generar el Word con --sugerencias para verlas en gris.')


def cmd_aceptar(a):
    """Convierte las sugerencias en texto definitivo (para las clases libres que se indiquen o todas)."""
    plan = _leer(a.plan)
    claves = mensual.campos_de(mensual.formato_de(plan))
    n = 0
    for p in plan['periodos']:
        for i, c in enumerate(p['clases']):
            if 'sugerencia' in c and (a.todas or f'{p["nombre"]}:{i + 1}' in (a.clases or [])):
                for k in claves:
                    if not str(c.get(k, '')).strip():
                        c[k] = c['sugerencia'].get(k, '')
                del c['sugerencia']; n += 1
    _guardar(plan, a.plan)
    print(f'{n} clases aceptadas')


# ------------------------------------------------------------------ anual / banco
def cmd_banco_importar(a):
    cat = _catalogo()
    manuales = dict(x.split('=', 1) for x in (a.codigo or []))
    datos = banco.importar_plan_anual(a.docx, a.curso, a.asignatura, a.anio, manuales)
    base = _datos_base(cat, a.curso)
    datos.update({k: base[k] for k in ('colegio', 'profesor', 'curso', 'curso_codigo')})
    ruta = banco.guardar(datos)
    sin_codigo = [o['texto'][:50] for o in datos['oa'] if not o['codigo']]
    print(ruta, f'| {len(datos["oa"])} OA')
    if sin_codigo:
        print('OA sin numero (usar --codigo "inicio del texto=N"):', sin_codigo)


def cmd_nuevo_anual(a):
    cat = _catalogo()
    base = _datos_base(cat, a.curso)
    plan = anual.plan_en_blanco(base['curso'], a.curso, a.asignatura, a.anio, cat['profesor'], a.horas or '', a.clases or '')
    plan['colegio'] = cat['colegio']
    if a.formato:
        plan['formato_perfil'] = a.formato
    ruta = a.salida or f'{a.curso} - {a.asignatura} - Plan Anual {a.anio}.json'
    print(_guardar(plan, ruta))


def cmd_generar_anual(a):
    plan = _leer(a.plan)
    plan.setdefault('curso_codigo', plan.get('curso'))
    destino = a.salida or os.path.splitext(a.plan)[0] + '.docx'
    print(anual.construir(plan, destino))


# ------------------------------------------------------------------ temario
def cmd_generar_temario(a):
    datos = _leer(a.datos)
    destino = a.salida or os.path.splitext(a.datos)[0] + '.docx'
    print(temario.construir(datos, destino))


# ------------------------------------------------------------------ idioma y autor de documentos ya creados
def cmd_normalizar(a):
    """Pone idioma es-CL y autor en uno o varios .docx existentes, sin tocar su contenido."""
    rutas = []
    for objetivo in a.rutas:
        if os.path.isdir(objetivo):
            for raiz, _, archivos in os.walk(objetivo):
                rutas += [os.path.join(raiz, f) for f in archivos if f.lower().endswith('.docx')]
        else:
            rutas.append(objetivo)
    ok = 0
    for ruta in rutas:
        if os.path.basename(ruta).startswith('~$'):
            continue
        try:
            documento_meta.normalizar_docx(ruta, autor_=a.autor)
            ok += 1
        except PermissionError:
            print('OMITIDO (abierto en Word):', ruta)
    print(f'{ok} documentos normalizados · autor: {a.autor or documento_meta.autor()} · idioma: {documento_meta.idioma()}')


# ------------------------------------------------------------------ formatos editables
def cmd_formato_listar(a):
    for tipo in ([a.tipo] if a.tipo else formato.TIPOS):
        print(f'{tipo}: ' + ', '.join(formato.listar(tipo)))


def cmd_formato_ver(a):
    print(json.dumps(formato.cargar(a.tipo, a.perfil), ensure_ascii=False, indent=1))


def cmd_formato_guardar(a):
    fmt = _leer(a.archivo)
    print('guardado como:', formato.guardar(a.tipo, a.nombre, fmt))


def cmd_formato_restablecer(a):
    print('perfil eliminado' if formato.restablecer(a.tipo, a.nombre) else 'no existia ese perfil')


def construir_parser():
    ap = argparse.ArgumentParser(prog='python -m planificacion', description='Planificaciones mensuales, planes anuales y temarios')
    sp = ap.add_subparsers(dest='cmd', required=True)

    p = sp.add_parser('nuevo-mensual', help='crea el JSON de una planificacion mensual con filas vacias')
    p.add_argument('--curso', required=True, help='codigo de curso, ej: 5°AB'); p.add_argument('--asignatura', required=True)
    p.add_argument('--anio', type=int, default=2026); p.add_argument('--unidad', required=True); p.add_argument('--nombre-unidad', required=True)
    p.add_argument('--oa', nargs='+', required=True, help='numeros de OA, ej: 8 9 10 11')
    p.add_argument('--mes', nargs='+', required=True, help='Mes:cantidad_de_clases, ej: Noviembre:15 Diciembre:2')
    p.add_argument('--banco', help='curso del banco curricular si difiere (ej: Orientacion 5°A usa el plan anual 5°AB)')
    p.add_argument('--formato', help='perfil de formato (ver formato-listar)')
    p.add_argument('--horas'); p.add_argument('-o', '--salida'); p.set_defaults(f=cmd_nuevo_mensual)

    p = sp.add_parser('generar-mensual', help='JSON -> Word'); p.add_argument('plan'); p.add_argument('-o', '--salida')
    p.add_argument('--sugerencias', action='store_true', help='mostrar las sugerencias (gris cursiva) en las celdas vacias')
    p.set_defaults(f=cmd_generar_mensual)

    p = sp.add_parser('leer-mensual', help='Word escrito por el docente -> JSON'); p.add_argument('docx')
    p.add_argument('--curso', required=True); p.add_argument('--anio', type=int, default=2026); p.add_argument('--formato')
    p.add_argument('-o', '--salida'); p.set_defaults(f=cmd_leer_mensual)

    p = sp.add_parser('estado', help='resumen corto de clases escritas/faltantes'); p.add_argument('plan'); p.set_defaults(f=cmd_estado)

    p = sp.add_parser('semana', help='escribe una clase concreta'); p.add_argument('plan'); p.add_argument('periodo'); p.add_argument('n', type=int)
    for c in mensual.CAMPOS:
        p.add_argument('--' + c)
    p.add_argument('--campo', nargs='*', help='columnas personalizadas: clave=valor')
    p.set_defaults(f=cmd_semana)

    p = sp.add_parser('sugerir', help='propone contenido para las clases vacias')
    p.add_argument('plan'); p.add_argument('--ia', default='offline', help='offline | gemini | claude | openai | deepseek | kimi | grok | local')
    p.add_argument('--modelo'); p.set_defaults(f=cmd_sugerir)

    p = sp.add_parser('aceptar', help='convierte sugerencias en texto definitivo'); p.add_argument('plan')
    p.add_argument('--todas', action='store_true'); p.add_argument('--clases', nargs='*', help='ej: Noviembre:3 Diciembre:1')
    p.set_defaults(f=cmd_aceptar)

    p = sp.add_parser('banco-importar', help='Plan Anual (.docx) -> banco de OA e indicadores (JSON)')
    p.add_argument('docx'); p.add_argument('--curso', required=True); p.add_argument('--asignatura', required=True)
    p.add_argument('--anio', type=int, required=True); p.add_argument('--codigo', nargs='*', help='OA sin numero: "inicio del texto=4"')
    p.set_defaults(f=cmd_banco_importar)

    p = sp.add_parser('nuevo-anual', help='JSON de plan anual en blanco (p.ej. para el proximo año)')
    p.add_argument('--curso', required=True); p.add_argument('--asignatura', required=True); p.add_argument('--anio', type=int, required=True)
    p.add_argument('--horas'); p.add_argument('--clases'); p.add_argument('--formato'); p.add_argument('-o', '--salida')
    p.set_defaults(f=cmd_nuevo_anual)

    p = sp.add_parser('generar-anual', help='JSON del plan anual (o del banco) -> Word'); p.add_argument('plan'); p.add_argument('-o', '--salida')
    p.set_defaults(f=cmd_generar_anual)

    p = sp.add_parser('generar-temario', help='JSON de temario -> Word'); p.add_argument('datos'); p.add_argument('-o', '--salida')
    p.set_defaults(f=cmd_generar_temario)

    p = sp.add_parser('normalizar', help='pone idioma es-CL y autor en .docx ya creados (archivos o carpetas)')
    p.add_argument('rutas', nargs='+'); p.add_argument('--autor', help='por defecto: config/documento.json o EDUDOCENTE_AUTOR')
    p.set_defaults(f=cmd_normalizar)

    p = sp.add_parser('formato-listar', help='perfiles de formato disponibles'); p.add_argument('tipo', nargs='?', choices=formato.TIPOS)
    p.set_defaults(f=cmd_formato_listar)
    p = sp.add_parser('formato-ver', help='muestra un formato (JSON editable)'); p.add_argument('tipo', choices=formato.TIPOS)
    p.add_argument('perfil', nargs='?'); p.set_defaults(f=cmd_formato_ver)
    p = sp.add_parser('formato-guardar', help='guarda un perfil de formato desde un JSON'); p.add_argument('tipo', choices=formato.TIPOS)
    p.add_argument('nombre'); p.add_argument('archivo'); p.set_defaults(f=cmd_formato_guardar)
    p = sp.add_parser('formato-restablecer', help='elimina un perfil guardado'); p.add_argument('tipo', choices=formato.TIPOS)
    p.add_argument('nombre'); p.set_defaults(f=cmd_formato_restablecer)
    return ap


def main(argv=None):
    if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    a = construir_parser().parse_args(argv)
    a.f(a)


if __name__ == '__main__':
    main()
