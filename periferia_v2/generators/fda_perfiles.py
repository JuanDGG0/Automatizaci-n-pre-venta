"""
generators/fda_perfiles.py
Responsable: Heidy

Edita slides 8 (Perfiles) y 11 (FDA) del PPTX de la filial.

Lógica Perfiles:
  - Si hay perfiles en Anexos del Excel → usar esos
  - Si hay torres pero sin perfiles → usar genéricos de esas torres
  - Si Excel vacío + usuario eligió torres → usar genéricos de esas torres

Lógica FDA:
  - Si hay 1 torre → mostrar ítems de esa torre (max 6)
  - Si hay más de 1 torre → mostrar cláusula general aplicable a todas
  - Si la torre es QA → ocultar card verde de QA (ya está como torre)
"""

import io, re, zipfile, unicodedata, random
from pathlib import Path
from lxml import etree
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / 'data'
GENERALES = DATA_DIR / 'Generales_para_todos.xlsx'

A  = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P  = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# Shapes slide FDA (11)
BULLET_RECTS = ['Rectángulo 10','Rectángulo 13','Rectángulo 19',
                'Rectángulo 22','Rectángulo 25','Rectángulo 28']
QA_CARD      = 'CuadroTexto 32'
TITULO_TECNICO = 'Fuera del Alcance Técnico'

# Cláusula general para múltiples torres
CLAUSULA_GENERAL = [
    'El servicio será ejecutado conforme al alcance técnico aprobado y a la información disponible al momento de la estimación.',
    'Cualquier modificación posterior en requerimientos funcionales, técnicos o de negocio será gestionada mediante la metodología formal de control de cambios.',
    'Cualquier actividad no descrita explícitamente en el alcance aprobado se considerará fuera de alcance.',
    'No incluye actividades de soporte continuo posterior al periodo de garantía definido.',
    'No incluye licenciamiento de herramientas, plataformas o componentes de terceros.',
    'No incluye infraestructura productiva ni ambientes no definidos en el alcance.',
]

def _norm(s):
    s = (s or '').strip().upper()
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', s).strip()

def _esc(t):
    return (t or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

def _load_generales():
    """Carga el Excel de generales y retorna dicts por hoja."""
    if not GENERALES.exists():
        return {}, {}
    wb = load_workbook(GENERALES)

    # FDA por torre
    fda_db = {}
    ws = wb['Fuera del Alcance']
    torre_actual = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and str(row[0]).strip() and str(row[0]).strip() != 'Torre':
            torre_actual = _norm(str(row[0]).strip())
            if row[1]:
                fda_db.setdefault(torre_actual, []).append(str(row[1]).strip())
        elif row[0] is None and row[1] and torre_actual:
            fda_db.setdefault(torre_actual, []).append(str(row[1]).strip())

    # Perfiles por torre
    perf_db = {}
    ws2 = wb['Perfiles']
    torre_actual = None
    for row in ws2.iter_rows(min_row=2, values_only=True):
        if row[0] and str(row[0]).strip():
            torre_actual = _norm(str(row[0]).strip().replace('TORRE ', ''))
        if torre_actual and row[2] and row[3]:  # rol, descripción
            perf_db.setdefault(torre_actual, []).append({
                'rol': str(row[2]).strip(),
                'desc': str(row[3]).strip()
            })

    return fda_db, perf_db

def _get_slide_order(pptx_bytes):
    """Retorna lista ordenada de paths de slides según presentation.xml."""
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        rid_map = {r.attrib['Id']: r.attrib['Target'] for r in rels}
        prs = etree.fromstring(z.read('ppt/presentation.xml'))
        ns = {'p': P, 'r': R}
        return ['ppt/' + rid_map[s.attrib[f'{{{R}}}id']]
                for s in prs.find('.//p:sldIdLst', ns)]

def _hide_shape(sp):
    """Oculta un shape poniéndolo invisible."""
    spPr = sp.find(f'{{{P}}}spPr')
    if spPr is None:
        spPr = etree.SubElement(sp, f'{{{P}}}spPr')
    ln = spPr.find(f'{{{A}}}ln')
    if ln is None:
        ln = etree.SubElement(spPr, f'{{{A}}}ln')
    noFill = etree.SubElement(ln, f'{{{A}}}noFill')
    solidFill = spPr.find(f'{{{A}}}solidFill')
    if solidFill is not None:
        spPr.remove(solidFill)
    noFill2 = etree.SubElement(spPr, f'{{{A}}}noFill')
    txb = sp.find(f'{{{P}}}txBody')
    if txb is not None:
        for t in txb.findall(f'.//{{{A}}}t'):
            t.text = ''

def _set_text_in_shape(sp, text, font_size=None):
    """Reemplaza el texto de un shape preservando formato."""
    txb = sp.find(f'{{{P}}}txBody')
    if txb is None:
        return
    # Obtener primer párrafo como template
    paras = txb.findall(f'{{{A}}}p')
    if not paras:
        return

    # Limpiar todos los párrafos
    for p in paras:
        txb.remove(p)

    # Crear nuevo párrafo con el texto
    p_xml = f'<a:p xmlns:a="{A}"><a:r>'
    if font_size:
        p_xml += f'<a:rPr sz="{font_size}"/>'
    p_xml += f'<a:t>{_esc(text)}</a:t></a:r></a:p>'
    txb.append(etree.fromstring(p_xml))

def _edit_fda_slide(xml_bytes, torres, fda_db, include_qa_card):
    """
    Edita el slide 11 (FDA).
    - Si 1 torre: ítems específicos de esa torre (max 6)
    - Si múltiples torres: cláusula general
    - Si QA es una de las torres: ocultar card QA
    """
    root = etree.fromstring(xml_bytes)

    torres_norm = [_norm(t) for t in torres]
    es_solo_qa = len(torres) == 1 and 'QA' in torres_norm[0]
    hay_qa = any('QA' in t for t in torres_norm)

    # Determinar ítems FDA
    if len(torres) == 1:
        torre_key = torres_norm[0]
        items = fda_db.get(torre_key, [])
        if not items:
            # Buscar match parcial
            for k in fda_db:
                if torre_key in k or k in torre_key:
                    items = fda_db[k]
                    break
        items = items[:6] if items else CLAUSULA_GENERAL[:6]
    else:
        items = CLAUSULA_GENERAL[:6]

    # Llenar bullets
    for sp in root.iter(f'{{{P}}}sp'):
        nvpr = sp.find(f'.//{{{P}}}cNvPr')
        if nvpr is None:
            continue
        name = nvpr.attrib.get('name', '')

        # Bullets izquierda
        if name in BULLET_RECTS:
            idx = BULLET_RECTS.index(name)
            txb = sp.find(f'{{{P}}}txBody')
            if txb is None:
                continue
            if idx < len(items):
                for t in txb.findall(f'.//{{{A}}}t'):
                    t.text = items[idx]
            else:
                _hide_shape(sp)

        # Card QA derecha
        elif name == QA_CARD:
            if hay_qa or not include_qa_card:
                # Ocultar card QA si QA es una torre o si no se quiere
                _hide_shape(sp)

        # Títulos
        elif 'Título 1' in name:
            txb = sp.find(f'{{{P}}}txBody')
            if txb is None:
                continue
            txt = ''.join(t.text or '' for t in txb.findall(f'.//{{{A}}}t')).strip()
            if TITULO_TECNICO in txt:
                if len(torres) == 1:
                    for t in txb.findall(f'.//{{{A}}}t'):
                        if t.text and TITULO_TECNICO in t.text:
                            t.text = f'Fuera del Alcance {torres[0]}'
                else:
                    for t in txb.findall(f'.//{{{A}}}t'):
                        if t.text and TITULO_TECNICO in t.text:
                            t.text = 'Fuera del Alcance General'
            elif 'Pruebas de Calidad' in txt and (hay_qa or not include_qa_card):
                _hide_shape(sp)

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def _edit_perfiles_slide(xml_bytes, perfiles):
    """
    Edita slide 8 (Perfiles).
    perfiles = [{'rol': str, 'desc': str}, ...] max 4
    """
    root = etree.fromstring(xml_bytes)

    # Shapes del slide: 4 perfiles, cada uno con CuadroTexto nombre + CuadroTexto desc
    PERFIL_SHAPES = [
        ('CuadroTexto 10', 'CuadroTexto 22'),
        ('CuadroTexto 30', 'CuadroTexto 28'),
        ('CuadroTexto 47', 'CuadroTexto 34'),
        ('CuadroTexto 53', 'CuadroTexto 51'),
    ]

    perfiles = perfiles[:4]

    shape_map = {}
    for sp in root.iter(f'{{{P}}}sp'):
        nvpr = sp.find(f'.//{{{P}}}cNvPr')
        if nvpr is not None:
            shape_map[nvpr.attrib.get('name','')] = sp

    for i, (nombre_id, desc_id) in enumerate(PERFIL_SHAPES):
        sp_nombre = shape_map.get(nombre_id)
        sp_desc   = shape_map.get(desc_id)

        if i < len(perfiles):
            p = perfiles[i]
            if sp_nombre:
                txb = sp_nombre.find(f'{{{P}}}txBody')
                if txb is not None:
                    for t in txb.findall(f'.//{{{A}}}t'):
                        t.text = p['rol']
                        break
            if sp_desc:
                txb = sp_desc.find(f'{{{P}}}txBody')
                if txb is not None:
                    # Limpiar y escribir descripción (puede ser larga)
                    lines = p['desc'].split('\n')
                    paras = txb.findall(f'{{{A}}}p')
                    # Usar primer párrafo como template
                    for para in paras[1:]:
                        txb.remove(para)
                    if paras:
                        ts = paras[0].findall(f'.//{{{A}}}t')
                        if ts:
                            ts[0].text = lines[0] if lines else ''
                    for line in lines[1:6]:  # max 6 líneas
                        p_xml = f'<a:p xmlns:a="{A}"><a:r><a:t>{_esc(line)}</a:t></a:r></a:p>'
                        txb.append(etree.fromstring(p_xml))
        else:
            # Ocultar perfiles vacíos
            if sp_nombre:
                _hide_shape(sp_nombre)
            if sp_desc:
                _hide_shape(sp_desc)

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def edit(pptx_bytes, config):
    """
    Punto de entrada principal.
    config = {
        filial: str,
        excel_data: { torres: [...], perfiles: [...] },
        torres_seleccionadas: [...],  # si excel vacío
        opciones: { perfiles: 'excel'|'genericos', fda: 'excel'|'genericos' },
        torres_qa: { 'TORRE': True/False }
    }
    """
    fda_db, perf_db = _load_generales()

    excel_data = config.get('excel_data') or {}
    excel_torres = excel_data.get('torres', [])
    excel_perfiles = excel_data.get('perfiles', [])
    torres_sel = config.get('torres_seleccionadas', [])
    opciones = config.get('opciones', {})
    torres_qa = config.get('torres_qa', {})

    # Determinar torres activas
    torres_activas = [t['nombre'] for t in excel_torres] if excel_torres else torres_sel

    # ── Determinar perfiles ──────────────────────────────────────────────
    if opciones.get('perfiles') == 'excel' and excel_perfiles:
        # Usar perfiles del Excel (Anexos)
        perfiles = [{'rol': p['perfil'], 'desc': p.get('seniority', '')} for p in excel_perfiles]
    else:
        # Usar genéricos de las torres activas
        perfiles = []
        for torre in torres_activas:
            key = _norm(torre)
            genericos = perf_db.get(key, [])
            # Buscar match parcial
            if not genericos:
                for k in perf_db:
                    if key in k or k in key:
                        genericos = perf_db[k]
                        break
            if genericos:
                perfiles.append(genericos[0])  # un perfil representativo por torre

    # ── Determinar si incluir card QA ────────────────────────────────────
    torres_norm = [_norm(t) for t in torres_activas]
    hay_torre_qa = any('QA' in t for t in torres_norm)
    include_qa_card = not hay_torre_qa  # si QA es torre, no mostrar card

    # ── Editar PPTX ──────────────────────────────────────────────────────
    slides_order = _get_slide_order(pptx_bytes)

    files_dict = {}
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zin:
        files_dict = {name: zin.read(name) for name in zin.namelist()}

    # Slide 11 → FDA (índice 10)
    fda_slide_key = slides_order[10]
    files_dict[fda_slide_key] = _edit_fda_slide(
        files_dict[fda_slide_key],
        torres_activas,
        fda_db,
        include_qa_card
    )

    # Slide 8 → Perfiles (índice 7)
    perf_slide_key = slides_order[7]
    files_dict[perf_slide_key] = _edit_perfiles_slide(
        files_dict[perf_slide_key],
        perfiles
    )

    # Reconstruir ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files_dict.items():
            zout.writestr(name, data)

    return buf.getvalue()
