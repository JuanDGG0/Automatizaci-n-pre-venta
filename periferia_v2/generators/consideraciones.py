"""
generators/consideraciones.py
Responsable: Juan

Edita slide 10 (Consideraciones) del PPTX.
El slide tiene 4 shapes tipo 'Redondear rectángulo de esquina diagonal 14'.

Lógica:
  - Toma consideraciones del Generales_para_todos.xlsx hoja 'Consideraciones'
  - Filtra por las torres activas del proyecto
  - Toma max 4 consideraciones (una por shape)
  - Reemplaza 'XXXXXXXXXX' por el nombre del cliente
  - Reemplaza 'Filial' por el nombre de la filial
"""

import io, zipfile, unicodedata, re
from pathlib import Path
from lxml import etree
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / 'data'
GENERALES = DATA_DIR / 'Generales_para_todos.xlsx'

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

SHAPE_NAME = 'Redondear rectángulo de esquina diagonal 14'

FILIAL_NOMBRES = {
    'corp':  'Periferia IT Corp',
    'group': 'Periferia IT Group',
    'cbit':  'Contact & Business IT'
}

def _norm(s):
    s = (s or '').strip().upper()
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', s).strip()

def _esc(t):
    return (t or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

def _load_consideraciones(torres_activas, cliente, filial):
    """Carga consideraciones del Excel filtrando por torres activas."""
    if not GENERALES.exists():
        return []

    wb = load_workbook(GENERALES)
    ws = wb['Consideraciones']

    torres_norm = [_norm(t) for t in torres_activas]
    consideraciones = []
    torre_actual = None

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and str(row[0]).strip():
            torre_actual = _norm(str(row[0]).strip())

        if row[1] and torre_actual:
            # Verificar si la torre aplica
            aplica = any(torre_actual in t or t in torre_actual for t in torres_norm)
            if aplica or torre_actual == 'SOPORTE':
                texto = str(row[1]).strip()
                texto = texto.replace('XXXXXXXXXX', cliente or 'El cliente')
                texto = texto.replace('Filial', filial or 'Periferia IT')
                consideraciones.append(texto)

    # Tomar máximo 4 distribuidas
    if len(consideraciones) > 4:
        step = len(consideraciones) // 4
        consideraciones = [consideraciones[i*step] for i in range(4)]

    return consideraciones[:4]

def _get_slide_order(pptx_bytes):
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        rid_map = {r.attrib['Id']: r.attrib['Target'] for r in rels}
        prs = etree.fromstring(z.read('ppt/presentation.xml'))
        ns = {'p': P, 'r': R}
        return ['ppt/' + rid_map[s.attrib[f'{{{R}}}id']]
                for s in prs.find('.//p:sldIdLst', ns)]


def _find_cons_slide(slides_order, files_dict):
    """
    Localiza el slide de Consideraciones por contenido (nombre de shape),
    no por índice fijo. Necesario porque fda_perfiles puede insertar slides
    extra antes de este slide, desplazando todos los índices posteriores.

    El slide de Consideraciones tiene EXACTAMENTE 4 shapes con SHAPE_NAME.
    Otros slides del template pueden tener 1 de este mismo shape — hay que
    buscar el que tenga >= 4 para no confundirlos.

    Fallback al índice 9 si no se encuentra ningún slide con 4+ shapes.
    """
    for path in slides_order:
        root = etree.fromstring(files_dict[path])
        count = sum(
            1 for sp in root.iter(f'{{{P}}}sp')
            for nvpr in [sp.find(f'.//{{{P}}}cNvPr')]
            if nvpr is not None and SHAPE_NAME in nvpr.attrib.get('name', '')
        )
        if count >= 4:
            return path
    # Fallback defensivo: índice original del template
    fallback_idx = min(9, len(slides_order) - 1)
    print(f'[CONSIDERACIONES] Advertencia: marcador "{SHAPE_NAME}" (×4) no encontrado. '
          f'Usando índice {fallback_idx} como fallback.')
    return slides_order[fallback_idx]

def _edit_consideraciones_slide(xml_bytes, consideraciones):
    """Edita el slide 10 con las consideraciones."""
    root = etree.fromstring(xml_bytes)

    shapes_cons = []
    for sp in root.iter(f'{{{P}}}sp'):
        nvpr = sp.find(f'.//{{{P}}}cNvPr')
        if nvpr is None:
            continue
        name = nvpr.attrib.get('name', '')
        if SHAPE_NAME in name:
            shapes_cons.append(sp)

    for i, sp in enumerate(shapes_cons):
        txb = sp.find(f'{{{P}}}txBody')
        if txb is None:
            continue
        if i < len(consideraciones):
            # Reemplazar texto preservando formato
            for t_el in txb.findall(f'.//{{{A}}}t'):
                t_el.text = consideraciones[i] if t_el == txb.findall(f'.//{{{A}}}t')[0] else ''
        # Si no hay consideración para este shape, dejar el placeholder

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def edit(pptx_bytes, config):
    """
    Punto de entrada.
    config = {
        filial: str,
        excel_data: { torres: [...], cliente: str },
        torres_seleccionadas: [...],
        consideraciones: [...],  # si el usuario las ingresó manualmente
        opciones: { consideraciones: 'genericos'|'manual' }
    }
    """
    excel_data = config.get('excel_data') or {}
    excel_torres = excel_data.get('torres', [])
    torres_sel = config.get('torres_seleccionadas', [])
    filial = config.get('filial', 'corp')
    cliente = excel_data.get('cliente', '')
    opciones = config.get('opciones', {})

    torres_activas = [t['nombre'] for t in excel_torres] if excel_torres else torres_sel
    filial_nombre = FILIAL_NOMBRES.get(filial, 'Periferia IT')

    # Determinar consideraciones
    # El JS envía opciones.consideraciones como boolean (true = "con genéricos").
    # Si la pill está activada → cargar genéricos.
    # Si la pill está desactivada → también cargar genéricos (no hay fuente Excel
    # para consideraciones; "solo del Excel" equivale a dejar el template sin cambios,
    # pero para esta sección siempre aplicamos los genéricos filtrados por torre).
    if opciones.get('consideraciones') == 'manual':
        # Compatibilidad con llamadas antiguas que envíen el string 'manual'
        consideraciones = config.get('consideraciones', [])[:4]
    else:
        consideraciones = _load_consideraciones(torres_activas, cliente, filial_nombre)

    # Editar slide
    slides_order = _get_slide_order(pptx_bytes)

    files_dict = {}
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zin:
        files_dict = {name: zin.read(name) for name in zin.namelist()}

    # Localizar slide de Consideraciones por contenido, no por índice fijo.
    # Índice fijo (9) se rompe si fda_perfiles añadió slides extra antes.
    cons_slide_key = _find_cons_slide(slides_order, files_dict)
    files_dict[cons_slide_key] = _edit_consideraciones_slide(
        files_dict[cons_slide_key],
        consideraciones
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files_dict.items():
            zout.writestr(name, data)

    return buf.getvalue()
