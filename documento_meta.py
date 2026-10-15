# -*- coding: utf-8 -*-
"""Metadatos estandar de los documentos Word generados por EduDocente-Studio.

Problema que resuelve: los .docx creados con python-docx salen con idioma ingles (Word subraya en rojo lo que se
escribe en español) y con autor "python-docx". Aqui se fija, para TODO documento del proyecto:

  * idioma de revision ortografica: español de Chile (es-CL)
  * autor / "ultima modificacion por": configurable

Autor (de mayor a menor prioridad):
  1. variable de entorno  EDUDOCENTE_AUTOR
  2. config/documento.local.json   (archivo local, ignorado por git)      {"autor": "Nombre"}
  3. config/documento.json         (valor por defecto del proyecto)       {"autor": "Jack", "idioma": "es-CL"}
  El valor "auto" usa el nombre de usuario del equipo (getpass.getuser()).

Uso:
    from documento_meta import aplicar, normalizar_docx
    aplicar(doc, titulo="Planificación mensual")        # antes de doc.save(...)
    normalizar_docx("archivo.docx")                      # corrige un .docx ya creado (conserva todo el contenido)
"""
import datetime
import getpass
import json
import os
import re
import tempfile
import zipfile

from lxml import etree

RAIZ = os.path.dirname(os.path.abspath(__file__))
IDIOMA_DEFECTO = 'es-CL'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}


def _leer_json(ruta):
    try:
        with open(ruta, encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def configuracion():
    """Configuracion efectiva {autor, idioma}."""
    cfg = {'autor': 'Jack', 'idioma': IDIOMA_DEFECTO}
    cfg.update({k: v for k, v in _leer_json(os.path.join(RAIZ, 'config', 'documento.json')).items() if v})
    cfg.update({k: v for k, v in _leer_json(os.path.join(RAIZ, 'config', 'documento.local.json')).items() if v})
    if os.environ.get('EDUDOCENTE_AUTOR'):
        cfg['autor'] = os.environ['EDUDOCENTE_AUTOR']
    if str(cfg['autor']).strip().lower() == 'auto':
        cfg['autor'] = getpass.getuser()
    return cfg


def autor():
    return configuracion()['autor']


def idioma():
    return configuracion()['idioma']


# ---------------------------------------------------------------- documentos nuevos (python-docx)
def _lang_en_rpr(rpr, lang):
    """Asegura <w:lang> dentro de un <w:rPr> respetando el orden del esquema."""
    for e in rpr.findall('w:lang', NS):
        rpr.remove(e)
    e = etree.SubElement(rpr, f'{{{W}}}lang')
    e.set(f'{{{W}}}val', lang); e.set(f'{{{W}}}eastAsia', lang); e.set(f'{{{W}}}bidi', 'ar-SA')
    for tag in ('eastAsianLayout', 'specVanish', 'oMath'):               # estos van despues de lang
        sig = rpr.find(f'w:{tag}', NS)
        if sig is not None:
            sig.addprevious(e)
            break


def aplicar(doc, titulo=None, autor_=None, idioma_=None):
    """Idioma es-CL y autor en un documento python-docx (llamar justo antes de guardar)."""
    cfg = configuracion()
    lang = idioma_ or cfg['idioma']
    quien = autor_ or cfg['autor']
    st = doc.styles.element
    dd = st.find('w:docDefaults', NS)
    if dd is not None:
        rpr = dd.find('w:rPrDefault/w:rPr', NS)
        if rpr is not None:
            _lang_en_rpr(rpr, lang)
    for e in st.iter(f'{{{W}}}lang'):                                      # estilos con idioma propio
        e.set(f'{{{W}}}val', lang)
        if e.get(f'{{{W}}}eastAsia'):
            e.set(f'{{{W}}}eastAsia', lang)
    ajustes = doc.settings.element
    for e in ajustes.iter(f'{{{W}}}themeFontLang'):
        e.set(f'{{{W}}}val', lang)
    cp = doc.core_properties
    ahora = datetime.datetime.now()
    cp.author = quien
    cp.last_modified_by = quien
    cp.comments = ''
    cp.language = lang
    cp.created = cp.modified = ahora
    cp.revision = 1
    if titulo:
        cp.title = titulo
    return doc


# ---------------------------------------------------------------- documentos ya creados (corrige el .docx en el disco)
def _fijar_idioma_parte(raiz, lang):
    """Cambia el idioma de revision de toda etiqueta <w:lang> por `lang`. Devuelve cuantas modifico."""
    n = 0
    for e in raiz.iter(f'{{{W}}}lang'):
        val = e.get(f'{{{W}}}val')
        if val != lang:
            e.set(f'{{{W}}}val', lang); n += 1
        ea = e.get(f'{{{W}}}eastAsia')
        if ea and ea.lower().startswith(('en', 'es')) and ea != lang:
            e.set(f'{{{W}}}eastAsia', lang)
    return n


def normalizar_docx(ruta, autor_=None, idioma_=None, titulo=None):
    """Pone idioma es-CL y autor en un .docx existente sin tocar su contenido (texto, tablas, formato).

    Edita unicamente: estilos/ajustes/encabezados (etiquetas de idioma) y docProps/core.xml (autor).
    Si el archivo esta abierto en Word lanza PermissionError. Devuelve {'idioma': n, 'autor': str}.
    """
    cfg = configuracion()
    lang = idioma_ or cfg['idioma']
    quien = autor_ or cfg['autor']
    cambios = 0
    fd, tmp = tempfile.mkstemp(suffix='.docx', dir=os.path.dirname(os.path.abspath(ruta)))
    os.close(fd)
    try:
        with zipfile.ZipFile(ruta) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                datos = zin.read(item.filename)
                nombre = item.filename
                if nombre.startswith('word/') and nombre.endswith('.xml') and nombre.count('/') == 1:
                    raiz = etree.fromstring(datos)
                    n = _fijar_idioma_parte(raiz, lang)
                    if nombre == 'word/styles.xml':
                        rpr = raiz.find('w:docDefaults/w:rPrDefault/w:rPr', NS)
                        if rpr is not None and rpr.find('w:lang', NS) is None:
                            _lang_en_rpr(rpr, lang); n += 1
                    if nombre == 'word/settings.xml':
                        for e in raiz.iter(f'{{{W}}}themeFontLang'):
                            e.set(f'{{{W}}}val', lang); n += 1
                    if n:
                        cambios += n
                        datos = etree.tostring(raiz, xml_declaration=True, encoding='UTF-8', standalone=True)
                elif nombre == 'docProps/core.xml':
                    datos = _core(datos, quien, lang, titulo)
                zout.writestr(item, datos)
        os.replace(tmp, ruta)                              # si Word tiene el archivo abierto, falla aqui (no se pierde nada)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise
    return {'idioma': cambios, 'autor': quien}


def _core(datos, quien, lang, titulo):
    ns = {'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
          'dc': 'http://purl.org/dc/elements/1.1/', 'dcterms': 'http://purl.org/dc/terms/'}
    raiz = etree.fromstring(datos)

    def poner(ruta, valor, ns_uri, tag):
        e = raiz.find(ruta, ns)
        if e is None:
            e = etree.SubElement(raiz, f'{{{ns_uri}}}{tag}')
        e.text = valor
    poner('dc:creator', quien, ns['dc'], 'creator')
    poner('cp:lastModifiedBy', quien, ns['cp'], 'lastModifiedBy')
    poner('dc:language', lang, ns['dc'], 'language')
    for e in raiz.findall('dc:description', ns):           # "generated by python-docx"
        if 'python-docx' in (e.text or ''):
            e.text = ''
    if titulo:
        poner('dc:title', titulo, ns['dc'], 'title')
    return etree.tostring(raiz, xml_declaration=True, encoding='UTF-8', standalone=True)
