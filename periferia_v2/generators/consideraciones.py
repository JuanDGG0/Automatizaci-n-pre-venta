"""
generators/consideraciones.py
Responsable: Juan

Edita el slide de Consideraciones del PPTX de la filial.

Estructura del slide:
  - Cada consideración es un grpSp (grupo) que contiene:
      · El shape redondeado (Redondear rectángulo de esquina diagonal 14)
      · El ícono ⓘ (pic)
  - El slide template tiene 4 grupos (Grupo 2, Grupo 9, Grupo 12, Grupo 15)
  - El color del template es un gradiente bg1 → E4EFD9: NO se toca

Lógica:
  - Pill OFF: usa excel_data.consideraciones (columna J del Excel del usuario)
  - Pill ON:  usa Generales_para_todos.xlsx hoja 'Consideraciones' por torre
  - Máximo 6 consideraciones por slide; si hay más se duplica el slide
  - Grupos sobrantes se eliminan del XML por completo (grpSp + pic interno)
  - Slide localizado por contenido (>= 4 shapes con SHAPE_NAME), no por índice
"""

import copy, io, re, unicodedata, zipfile
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

PLACEHOLDER_CLIENTE = 'XXXXXXXXXX'
PLACEHOLDER_FILIAL  = 'Filial'
MAX_POR_SLIDE       = 4

FILIAL_NOMBRES = {
    'corp':  'Periferia IT Corp',
    'group': 'Periferia IT Group',
    'cbit':  'Contact & Business IT',
}


# ═══════════════════════════════ utilidades ══════════════════════════════════

def _norm(s):
    s = (s or '').strip().upper()
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', s).strip()


def _apply_replacements(texto, cliente, filial_nombre):
    texto = texto.replace(PLACEHOLDER_CLIENTE, cliente or 'el cliente')
    texto = texto.replace(PLACEHOLDER_FILIAL,  filial_nombre)
    return texto


# ═══════════════════════════════ fuentes de datos ════════════════════════════

def _load_desde_excel(excel_consideraciones, cliente, filial_nombre):
    """Pill OFF: lista de strings de columna J, parseada por el frontend."""
    if not excel_consideraciones:
        print('[CONSIDERACIONES] Pill OFF: excel_data.consideraciones vacío.')
        return []

    resultado = []
    for item in excel_consideraciones:
        if not item or not str(item).strip():
            continue
        texto = _apply_replacements(str(item).strip(), cliente, filial_nombre)
        if texto not in resultado:
            resultado.append(texto)
    return resultado


def _load_desde_generales(torres_activas, cliente, filial_nombre):
    """Pill ON: genéricos de Generales_para_todos.xlsx filtrados por torre."""
    if not GENERALES.exists():
        print('[CONSIDERACIONES] Advertencia: Generales_para_todos.xlsx no encontrado.')
        return []

    wb           = load_workbook(GENERALES)
    ws           = wb['Consideraciones']
    torres_norm  = {_norm(t) for t in torres_activas}
    torre_actual = None
    resultado    = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        col_a = row[0]
        col_b = row[1] if len(row) > 1 else None

        if col_a and str(col_a).strip():
            torre_actual = _norm(str(col_a).strip().replace('TORRE ', ''))

        if not (col_b and str(col_b).strip() and torre_actual):
            continue

        aplica = any(torre_actual in t or t in torre_actual for t in torres_norm)
        if not aplica:
            continue

        texto = _apply_replacements(str(col_b).strip(), cliente, filial_nombre)
        if texto not in resultado:
            resultado.append(texto)

    return resultado


# ═══════════════════════════════ manipulación XML ════════════════════════════

def _find_grupos(root):
    """
    Retorna lista de grpSp que contienen un shape con SHAPE_NAME,
    ordenados por posición Y (orden visual de arriba a abajo).
    """
    spTree = root.find(f'.//{{{P}}}spTree')
    grupos = []
    for child in list(spTree):
        if child.tag != f'{{{P}}}grpSp':
            continue
        tiene_shape = any(
            nvpr.attrib.get('name', '') == SHAPE_NAME
            for sp   in child.iter(f'{{{P}}}sp')
            for nvpr in [sp.find(f'.//{{{P}}}cNvPr')]
            if nvpr is not None
        )
        if tiene_shape:
            # leer Y del grpSpPr para ordenar
            grpSpPr = child.find(f'{{{P}}}grpSpPr')
            xfrm    = grpSpPr.find(f'{{{A}}}xfrm') if grpSpPr is not None else None
            off     = xfrm.find(f'{{{A}}}off') if xfrm is not None else None
            y       = int(off.attrib.get('y', 0)) if off is not None else 0
            grupos.append((y, child))

    grupos.sort(key=lambda g: g[0])
    return [g for _, g in grupos]


def _write_text_in_grupo(grpSp, texto):
    """Escribe texto en el shape redondeado dentro del grupo."""
    for sp in grpSp.iter(f'{{{P}}}sp'):
        nvpr = sp.find(f'.//{{{P}}}cNvPr')
        if nvpr is None or nvpr.attrib.get('name', '') != SHAPE_NAME:
            continue
        txb   = sp.find(f'{{{P}}}txBody')
        if txb is None:
            continue
        all_t = txb.findall(f'.//{{{A}}}t')
        if all_t:
            all_t[0].text = texto
            for t in all_t[1:]:
                t.text = ''
        break


def _remove_grupo(root, grpSp):
    """Elimina el grpSp completo del spTree."""
    spTree = root.find(f'.//{{{P}}}spTree')
    if spTree is not None:
        spTree.remove(grpSp)


def _duplicate_grupo(root, template_grpSp, y_offset):
    """
    Clona el template_grpSp, lo posiciona en y_offset y lo añade al spTree.
    Retorna el nuevo grpSp.
    """
    spTree  = root.find(f'.//{{{P}}}spTree')
    new_grp = copy.deepcopy(template_grpSp)

    # Actualizar Y del grpSpPr
    grpSpPr = new_grp.find(f'{{{P}}}grpSpPr')
    xfrm    = grpSpPr.find(f'{{{A}}}xfrm') if grpSpPr is not None else None
    off     = xfrm.find(f'{{{A}}}off') if xfrm is not None else None
    if off is not None:
        off.attrib['y'] = str(y_offset)

    # ID único para no colisionar con shapes existentes
    existing_ids = [
        int(el.attrib.get('id', 0))
        for el in root.iter()
        if 'id' in el.attrib
    ]
    next_id = max(existing_ids, default=0) + 1
    for el in new_grp.iter():
        if 'id' in el.attrib:
            el.attrib['id'] = str(next_id)
            next_id += 1

    spTree.append(new_grp)
    return new_grp


def _get_grupo_height(grpSp):
    """Retorna la altura (cy) del grupo en EMU."""
    grpSpPr = grpSp.find(f'{{{P}}}grpSpPr')
    xfrm    = grpSpPr.find(f'{{{A}}}xfrm') if grpSpPr is not None else None
    ext     = xfrm.find(f'{{{A}}}ext') if xfrm is not None else None
    return int(ext.attrib.get('cy', 871144)) if ext is not None else 871144


def _get_grupo_y(grpSp):
    """Retorna la posición Y del grupo en EMU."""
    grpSpPr = grpSp.find(f'{{{P}}}grpSpPr')
    xfrm    = grpSpPr.find(f'{{{A}}}xfrm') if grpSpPr is not None else None
    off     = xfrm.find(f'{{{A}}}off') if xfrm is not None else None
    return int(off.attrib.get('y', 0)) if off is not None else 0


# ═══════════════════════════════ localización y duplicación de slide ══════════

def _get_slide_order(pptx_bytes):
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        rels    = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        rid_map = {r.attrib['Id']: r.attrib['Target'] for r in rels}
        prs     = etree.fromstring(z.read('ppt/presentation.xml'))
        ns      = {'p': P, 'r': R}
        return ['ppt/' + rid_map[s.attrib[f'{{{R}}}id']]
                for s in prs.find('.//p:sldIdLst', ns)]


def _find_cons_slide(slides_order, files_dict):
    """Localiza el slide con >= 4 grpSp que contengan SHAPE_NAME."""
    for path in slides_order:
        root  = etree.fromstring(files_dict[path])
        count = sum(
            1
            for child in root.find(f'.//{{{P}}}spTree') or []
            if child.tag == f'{{{P}}}grpSp' and any(
                nvpr.attrib.get('name', '') == SHAPE_NAME
                for sp   in child.iter(f'{{{P}}}sp')
                for nvpr in [sp.find(f'.//{{{P}}}cNvPr')]
                if nvpr is not None
            )
        )
        if count >= 4:
            return path

    fallback_idx = min(9, len(slides_order) - 1)
    print(f'[CONSIDERACIONES] Advertencia: slide no encontrado. Usando índice {fallback_idx}.')
    return slides_order[fallback_idx]


def _duplicate_slide(files_dict, src_path, insert_after_path):
    """
    Duplica src_path e inserta la copia justo después de insert_after_path.
    Mismo patrón que fda_perfiles._duplicate_perf_slide.
    Retorna el path del nuevo slide.
    """
    NS_REL  = 'http://schemas.openxmlformats.org/package/2006/relationships'
    SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

    existing_nums = [
        int(re.search(r'slide(\d+)', f).group(1))
        for f in files_dict
        if re.match(r'ppt/slides/slide\d+\.xml$', f)
    ]
    new_num        = max(existing_nums) + 1
    new_slide_path = f'ppt/slides/slide{new_num}.xml'
    files_dict[new_slide_path] = files_dict[src_path]

    # Content_Types
    CT_NS   = 'http://schemas.openxmlformats.org/package/2006/content-types'
    ct_root = etree.fromstring(files_dict['[Content_Types].xml'])
    etree.SubElement(ct_root, f'{{{CT_NS}}}Override', {
        'PartName':    f'/ppt/slides/slide{new_num}.xml',
        'ContentType': SLIDE_CT,
    })
    files_dict['[Content_Types].xml'] = etree.tostring(
        ct_root, xml_declaration=True, encoding='UTF-8', standalone=True)

    # Rels del slide
    def _rels_path(p):
        parts = p.rsplit('/', 1)
        return f"{parts[0]}/_rels/{parts[1]}.rels"

    src_rels = _rels_path(src_path)
    new_rels = _rels_path(new_slide_path)
    if src_rels in files_dict:
        rels_root = etree.fromstring(files_dict[src_rels])
        for rel in rels_root.findall(f'{{{NS_REL}}}Relationship'):
            if 'notesSlide' in rel.attrib.get('Type', ''):
                rels_root.remove(rel)
        files_dict[new_rels] = etree.tostring(
            rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)

    # presentation.xml.rels
    prs_rels_path = 'ppt/_rels/presentation.xml.rels'
    prs_rels_root = etree.fromstring(files_dict[prs_rels_path])

    ref_target = insert_after_path.replace('ppt/', '')
    ref_rid    = next((
        r.attrib['Id']
        for r in prs_rels_root.findall(f'{{{NS_REL}}}Relationship')
        if r.attrib.get('Target', '') == ref_target
    ), None)

    rid_nums = [
        int(m.group(1))
        for r in prs_rels_root.findall(f'{{{NS_REL}}}Relationship')
        for m in [re.search(r'rId(\d+)', r.attrib.get('Id', ''))]
        if m
    ]
    new_rid = f'rId{max(rid_nums) + 1}'
    etree.SubElement(prs_rels_root, f'{{{NS_REL}}}Relationship', {
        'Id':     new_rid,
        'Type':   f'{R}/slide',
        'Target': f'slides/slide{new_num}.xml',
    })
    files_dict[prs_rels_path] = etree.tostring(
        prs_rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)

    # presentation.xml sldIdLst
    prs_root = etree.fromstring(files_dict['ppt/presentation.xml'])
    sldIdLst = prs_root.find(f'.//{{{P}}}sldIdLst')
    max_id   = max(int(s.attrib['id']) for s in sldIdLst)
    new_sld  = etree.Element(f'{{{P}}}sldId')
    new_sld.attrib['id']           = str(max_id + 1)
    new_sld.attrib[f'{{{R}}}id']   = new_rid

    children = list(sldIdLst)
    for child in children:
        sldIdLst.remove(child)
    inserted = False
    for child in children:
        sldIdLst.append(child)
        if child.attrib.get(f'{{{R}}}id') == ref_rid and not inserted:
            sldIdLst.append(new_sld)
            inserted = True
    if not inserted:
        sldIdLst.append(new_sld)

    files_dict['ppt/presentation.xml'] = etree.tostring(
        prs_root, xml_declaration=True, encoding='UTF-8', standalone=True)

    return new_slide_path


# ═══════════════════════════════ edición del slide ═══════════════════════════

def _edit_cons_slide(xml_bytes, chunk):
    """
    Edita un slide con hasta 4 consideraciones (máximo que cabe en el template).
    - Escribe texto en los grupos existentes sin tocar color ni estilo.
    - Elimina los grupos sobrantes si hay menos de 4 consideraciones.
    """
    root   = etree.fromstring(xml_bytes)
    grupos = _find_grupos(root)   # siempre 4 en el template

    for i, grpSp in enumerate(grupos):
        if i < len(chunk):
            _write_text_in_grupo(grpSp, chunk[i])
        else:
            _remove_grupo(root, grpSp)

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ═══════════════════════════════ entry point ═════════════════════════════════

def edit(pptx_bytes, config):
    """
    config = {
        filial: str,
        excel_data: {
            torres: [{ nombre: str }, ...],
            cliente: str,
            consideraciones: [str, ...],   # columna J parseada por el frontend
        },
        torres_seleccionadas: [...],
        opciones: { consideraciones: bool },  # true = pill ON → genéricos
    }
    """
    excel_data            = config.get('excel_data') or {}
    excel_torres          = excel_data.get('torres', [])
    torres_sel            = config.get('torres_seleccionadas', [])
    filial_key            = config.get('filial', 'corp')
    cliente               = excel_data.get('cliente') or config.get('cliente', '')
    opciones              = config.get('opciones', {})
    usar_genericos        = bool(opciones.get('consideraciones', False))
    excel_consideraciones = excel_data.get('consideraciones', [])

    torres_activas = [t['nombre'] for t in excel_torres] if excel_torres else torres_sel
    filial_nombre  = FILIAL_NOMBRES.get(filial_key, 'Periferia IT')

    print(f'[CONSIDERACIONES] Torres activas  : {torres_activas}')
    print(f'[CONSIDERACIONES] Cliente         : {cliente}')
    print(f'[CONSIDERACIONES] Filial          : {filial_nombre}')
    print(f'[CONSIDERACIONES] Pill genéricos  : {"ON" if usar_genericos else "OFF"}')

    if usar_genericos:
        consideraciones = _load_desde_generales(torres_activas, cliente, filial_nombre)
    else:
        consideraciones = _load_desde_excel(excel_consideraciones, cliente, filial_nombre)

    print(f'[CONSIDERACIONES] Total           : {len(consideraciones)}')

    # Paginar en chunks de MAX_POR_SLIDE (4 — máximo que cabe físicamente)
    chunks = [consideraciones[i:i+MAX_POR_SLIDE]
              for i in range(0, max(len(consideraciones), 1), MAX_POR_SLIDE)]

    # Cargar archivos del PPTX
    files_dict = {}
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zin:
        files_dict = {name: zin.read(name) for name in zin.namelist()}

    slides_order   = _get_slide_order(pptx_bytes)
    cons_slide_key = _find_cons_slide(slides_order, files_dict)
    template_xml   = files_dict[cons_slide_key]

    # Editar el primer slide
    files_dict[cons_slide_key] = _edit_cons_slide(template_xml, chunks[0])

    # Duplicar para slides adicionales si hay más de MAX_POR_SLIDE
    prev_path = cons_slide_key
    for chunk in chunks[1:]:
        new_path = _duplicate_slide(files_dict, cons_slide_key, prev_path)
        files_dict[new_path] = _edit_cons_slide(template_xml, chunk)
        prev_path = new_path

    # Reconstruir ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files_dict.items():
            zout.writestr(name, data)

    return buf.getvalue()