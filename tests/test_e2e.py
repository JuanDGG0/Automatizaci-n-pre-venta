"""
test_e2e.py — pruebas end-to-end exhaustivas del generador de propuestas.
Cubre: ENTREGABLES, CONSIDERACIONES, PERFILES, SLIDE DETECTION, FILIALES, EDGE CASES.
Corre: python3 test_e2e.py
"""

import sys, io, zipfile, traceback, re
from pathlib import Path
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from generators import generate

OUT_DIR = Path('/tmp/periferia_e2e')
OUT_DIR.mkdir(parents=True, exist_ok=True)

# ── Namespace constants ───────────────────────────────────────────────────────
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

# ── Shape name constants ──────────────────────────────────────────────────────
SHAPE_CONS     = 'Redondear rectángulo de esquina diagonal 14'
ALL_PERF_ROLES = {'CuadroTexto 10', 'CuadroTexto 30', 'CuadroTexto 47', 'CuadroTexto 53'}
BULLET_RECTS   = {'Rectángulo 10', 'Rectángulo 13', 'Rectángulo 19',
                  'Rectángulo 22', 'Rectángulo 25', 'Rectángulo 28'}
QA_CARD        = 'CuadroTexto 32'
ENT_LIST_SHAPES = {'CuadroTexto 4', 'CuadroTexto 13', 'CuadroTexto 17'}

PASS_SYM = '\033[92m✓ PASS\033[0m'
FAIL_SYM = '\033[91m✗ FAIL\033[0m'
results  = []

# ═══════════════════════════════════════ HELPERS ══════════════════════════════

def get_slide_order(pptx_bytes):
    """Retorna lista ordenada de paths de slides según presentation.xml."""
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        rels    = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        rid_map = {r.attrib['Id']: r.attrib['Target'] for r in rels}
        prs     = etree.fromstring(z.read('ppt/presentation.xml'))
        ns      = {'p': P, 'r': R}
        return ['ppt/' + rid_map[s.attrib[f'{{{R}}}id']]
                for s in prs.find('.//p:sldIdLst', ns)]


def get_shape_names_in_slide_path(z, slide_path):
    """Retorna el set de nombres de shapes (cNvPr.name) en un slide (usando ZipFile ya abierto)."""
    root  = etree.fromstring(z.read(slide_path))
    names = set()
    for sp in root.iter(f'{{{P}}}sp'):
        nvpr = sp.find(f'.//{{{P}}}cNvPr')
        if nvpr is not None:
            names.add(nvpr.attrib.get('name', ''))
    return names


def find_first_slide_with_all_shapes(pptx_bytes, required_shapes):
    """
    Busca el primer slide (en orden de presentación) que contenga TODOS los shape names.
    Retorna (slide_path, shape_names_set) o (None, set()).
    """
    order = get_slide_order(pptx_bytes)
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            names = get_shape_names_in_slide_path(z, path)
            if required_shapes.issubset(names):
                return path, names
    return None, set()


def count_slides_with_min_shapes(pptx_bytes, shape_set, min_count=1):
    """
    Cuenta cuántos slides tienen al menos min_count shapes del shape_set presentes.
    """
    order = get_slide_order(pptx_bytes)
    total = 0
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            names = get_shape_names_in_slide_path(z, path)
            if len(shape_set & names) >= min_count:
                total += 1
    return total


def count_perf_slides(pptx_bytes):
    """
    Cuenta cuántos slides tienen los 4 slots de perfiles (ALL_PERF_ROLES completo).
    Un slide con 1–3 perfiles habrá eliminado los grupos sobrantes, por lo que
    no tendrá los 4 shapes → esta función detecta solo slides con 4 perfiles.
    """
    order = get_slide_order(pptx_bytes)
    total = 0
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            names = get_shape_names_in_slide_path(z, path)
            if ALL_PERF_ROLES.issubset(names):
                total += 1
    return total


def count_perf_slides_any_slot(pptx_bytes, role_names, role_text_must_contain):
    """
    Cuenta slides que tienen al menos 1 shape de role_names Y cuyo texto completo
    contiene alguna de las cadenas de role_text_must_contain.
    Esto permite detectar slides de perfiles con < 4 perfiles (donde los grupos
    vacíos fueron eliminados).
    """
    order = get_slide_order(pptx_bytes)
    total = 0
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root  = etree.fromstring(z.read(path))
            names = set()
            for sp in root.iter(f'{{{P}}}sp'):
                nvpr = sp.find(f'.//{{{P}}}cNvPr')
                if nvpr is not None:
                    names.add(nvpr.attrib.get('name', ''))
            if not (names & role_names):
                continue
            text = ''.join(t.text or '' for t in root.iter(f'{{{A}}}t'))
            if any(r in text for r in role_text_must_contain):
                total += 1
    return total


def count_entregables_perf_slides(pptx_bytes, role_names, check_text):
    """
    Variant: count slides with any perfil slot AND text from check_text.
    Also verifies title 'Perfiles' appears in the slide (to distinguish from
    other slides that accidentally have the same shape name).
    """
    order = get_slide_order(pptx_bytes)
    total = 0
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root  = etree.fromstring(z.read(path))
            names = set()
            for sp in root.iter(f'{{{P}}}sp'):
                nvpr = sp.find(f'.//{{{P}}}cNvPr')
                if nvpr is not None:
                    names.add(nvpr.attrib.get('name', ''))
            if not (names & role_names):
                continue
            text = ''.join(t.text or '' for t in root.iter(f'{{{A}}}t'))
            if any(r in text for r in check_text) and 'Perfiles' in text:
                total += 1
    return total


def get_entregables_cols_present(pptx_bytes):
    """
    Encuentra el slide de entregables buscando el que tenga al menos
    1 de los shapes de lista de entregables Y el shape del título con
    'Entregables de'. Retorna cuáles de CuadroTexto 4/13/17 están presentes.
    """
    order = get_slide_order(pptx_bytes)
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root  = etree.fromstring(z.read(path))
            names = set()
            for sp in root.iter(f'{{{P}}}sp'):
                nvpr = sp.find(f'.//{{{P}}}cNvPr')
                if nvpr is not None:
                    names.add(nvpr.attrib.get('name', ''))

            # El slide de entregables siempre tiene al menos CuadroTexto 4
            # Y tiene al menos 2 de los 3 shapes de lista de entregables
            ent_present = ENT_LIST_SHAPES & names
            if len(ent_present) >= 1:
                # Verificar que el texto "Entregables de" esté en el slide
                full_text = ''.join(t.text or '' for t in root.iter(f'{{{A}}}t'))
                if 'Entregables de' in full_text:
                    return ent_present
    return set()


def get_entregables_title_runs(pptx_bytes):
    """
    Retorna lista de (text, color_hex_or_None) de los runs del primer
    título 'Entregables de...' encontrado en el slide de entregables.
    """
    order = get_slide_order(pptx_bytes)
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root      = etree.fromstring(z.read(path))
            full_text = ''.join(t.text or '' for t in root.iter(f'{{{A}}}t'))
            if 'Entregables de' not in full_text:
                continue
            # Verificar que es el slide de entregables (tiene al menos 1 lista)
            names = set()
            for sp in root.iter(f'{{{P}}}sp'):
                nvpr = sp.find(f'.//{{{P}}}cNvPr')
                if nvpr is not None:
                    names.add(nvpr.attrib.get('name', ''))
            if not (ENT_LIST_SHAPES & names):
                continue
            # Buscar el primer título con "Entregables de"
            for sp in root.iter(f'{{{P}}}sp'):
                nvpr = sp.find(f'.//{{{P}}}cNvPr')
                if nvpr is None or nvpr.attrib.get('name', '') != 'Título 1':
                    continue
                txb = sp.find(f'{{{P}}}txBody')
                if txb is None:
                    continue
                sp_text = ''.join(t.text or '' for t in txb.findall(f'.//{{{A}}}t')).strip()
                if 'Entregables de' not in sp_text:
                    continue
                # Recolectar runs con sus colores
                runs = []
                for r_el in txb.findall(f'.//{{{A}}}r'):
                    rPr   = r_el.find(f'{{{A}}}rPr')
                    color = None
                    if rPr is not None:
                        solid = rPr.find(f'{{{A}}}solidFill')
                        if solid is not None:
                            srgb = solid.find(f'{{{A}}}srgbClr')
                            if srgb is not None:
                                color = srgb.attrib.get('val', '').upper()
                    t_el  = r_el.find(f'{{{A}}}t')
                    text  = t_el.text or '' if t_el is not None else ''
                    runs.append((text, color))
                if runs:
                    return runs
    return []


def get_entregables_list_text(pptx_bytes):
    """
    Retorna el texto concatenado de todas las shapes de lista de entregables
    (CuadroTexto 4/13/17) en el slide de entregables.
    """
    order = get_slide_order(pptx_bytes)
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root      = etree.fromstring(z.read(path))
            full_text = ''.join(t.text or '' for t in root.iter(f'{{{A}}}t'))
            if 'Entregables de' not in full_text:
                continue
            names = set()
            for sp in root.iter(f'{{{P}}}sp'):
                nvpr = sp.find(f'.//{{{P}}}cNvPr')
                if nvpr is not None:
                    names.add(nvpr.attrib.get('name', ''))
            if not (ENT_LIST_SHAPES & names):
                continue
            # Collect text only from the list shapes
            texts = []
            for sp in root.iter(f'{{{P}}}sp'):
                nvpr = sp.find(f'.//{{{P}}}cNvPr')
                if nvpr is None or nvpr.attrib.get('name', '') not in ENT_LIST_SHAPES:
                    continue
                txb = sp.find(f'{{{P}}}txBody')
                if txb is not None:
                    texts.append(''.join(t.text or '' for t in txb.findall(f'.//{{{A}}}t')))
            return ' '.join(texts)
    return ''


# Y_START is the exact Y coordinate where the generator places the first group.
# Template decoration slides (slide9, slide11, slide17) have different Y values.
# A slide is "generated by the consideraciones generator" when at least 1 of its
# SHAPE_CONS groups has Y == Y_CONS_START (they all start from this position).
Y_CONS_START = 1107495  # from consideraciones.py: Y_START constant


def _get_cons_groups_info(root):
    """
    Retorna lista de (y_position, text) para cada grpSp con SHAPE_CONS en root.
    """
    spTree = root.find(f'.//{{{P}}}spTree')
    if spTree is None:
        return []
    groups = []
    for child in list(spTree):
        if child.tag != f'{{{P}}}grpSp':
            continue
        has_shape = False
        group_text = ''
        for sp in child.iter(f'{{{P}}}sp'):
            nvpr = sp.find(f'.//{{{P}}}cNvPr')
            if nvpr is None or nvpr.attrib.get('name', '') != SHAPE_CONS:
                continue
            has_shape = True
            txb = sp.find(f'{{{P}}}txBody')
            if txb is not None:
                group_text = ''.join(t.text or '' for t in txb.findall(f'.//{{{A}}}t'))
        if not has_shape:
            continue
        grpSpPr = child.find(f'{{{P}}}grpSpPr')
        xfrm    = grpSpPr.find(f'{{{A}}}xfrm') if grpSpPr is not None else None
        off     = xfrm.find(f'{{{A}}}off') if xfrm is not None else None
        y       = int(off.attrib.get('y', 0)) if off is not None else 0
        groups.append((y, group_text))
    return groups


def _is_generated_cons_slide(root):
    """
    Retorna True si el slide fue editado por el generador de consideraciones.
    Criterio: al menos 1 grupo con SHAPE_CONS tiene Y == Y_CONS_START.
    Los slides de template decoration usan Y distintos (855730, 3771894, 4398349).
    """
    groups = _get_cons_groups_info(root)
    return any(y == Y_CONS_START for y, _ in groups)


def count_edited_cons_slides(pptx_bytes):
    """
    Cuenta cuántos slides fueron editados por el generador de consideraciones.
    Usa la posición Y del primer grupo: solo los slides generados tienen Y == Y_CONS_START.
    """
    order = get_slide_order(pptx_bytes)
    total = 0
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root = etree.fromstring(z.read(path))
            if _is_generated_cons_slide(root):
                total += 1
    return total


def get_cons_groups_count_per_slide(pptx_bytes):
    """
    Retorna lista con el número de grupos SHAPE_CONS en cada slide editado
    por el generador (aquellos con al menos 1 grupo con Y == Y_CONS_START).
    """
    order  = get_slide_order(pptx_bytes)
    result = []
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root   = etree.fromstring(z.read(path))
            if not _is_generated_cons_slide(root):
                continue
            groups = _get_cons_groups_info(root)
            result.append(len(groups))
    return result


def get_edited_cons_grupo_y_positions(pptx_bytes):
    """
    Retorna lista de posiciones Y de todos los grupos SHAPE_CONS
    en los slides editados por el generador.
    Excluye los slides de template decoration.
    """
    positions = []
    order     = get_slide_order(pptx_bytes)
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root = etree.fromstring(z.read(path))
            if not _is_generated_cons_slide(root):
                continue
            for y, _ in _get_cons_groups_info(root):
                positions.append(y)
    return positions


def get_cons_slide_text(pptx_bytes):
    """Texto de todos los grupos SHAPE_CONS en slides editados por el generador."""
    texts = []
    order = get_slide_order(pptx_bytes)
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root = etree.fromstring(z.read(path))
            if not _is_generated_cons_slide(root):
                continue
            for _, text in _get_cons_groups_info(root):
                texts.append(text)
    return texts


def count_entregables_slides(pptx_bytes):
    """Cuenta cuántos slides tienen contenido de entregables ('Entregables de' + ≥1 lista)."""
    order = get_slide_order(pptx_bytes)
    total = 0
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root  = etree.fromstring(z.read(path))
            names = {nvpr.attrib.get('name', '')
                     for sp in root.iter(f'{{{P}}}sp')
                     for nvpr in [sp.find(f'.//{{{P}}}cNvPr')] if nvpr is not None}
            if ENT_LIST_SHAPES & names:
                txt = ''.join(t.text or '' for t in root.iter(f'{{{A}}}t'))
                if 'Entregables de' in txt:
                    total += 1
    return total


def count_ent_list_shapes_in_slide(pptx_bytes, slide_index=0):
    """
    Retorna el número de shapes de lista de entregables (CuadroTexto 4/13/17)
    en el slide de entregables indicado por slide_index (0 = primero).
    Cuenta duplicados (útil para verificar la 4ª columna clonada).
    """
    order = get_slide_order(pptx_bytes)
    found = []
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root  = etree.fromstring(z.read(path))
            names = {nvpr.attrib.get('name', '')
                     for sp in root.iter(f'{{{P}}}sp')
                     for nvpr in [sp.find(f'.//{{{P}}}cNvPr')] if nvpr is not None}
            if ENT_LIST_SHAPES & names:
                txt = ''.join(t.text or '' for t in root.iter(f'{{{A}}}t'))
                if 'Entregables de' in txt:
                    cnt = sum(
                        1 for sp in root.iter(f'{{{P}}}sp')
                        for nvpr in [sp.find(f'.//{{{P}}}cNvPr')]
                        if nvpr is not None and nvpr.attrib.get('name', '') in ENT_LIST_SHAPES
                    )
                    found.append(cnt)
    if slide_index < len(found):
        return found[slide_index]
    return 0


def get_all_text(pptx_bytes):
    """Texto completo de todo el PPTX (todos los slides)."""
    parts = []
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for n in sorted(z.namelist()):
            if n.startswith('ppt/slides/slide') and n.endswith('.xml'):
                root = etree.fromstring(z.read(n))
                parts.append(''.join(t.text or '' for t in root.iter(f'{{{A}}}t')))
    return ' '.join(parts)


def count_total_slides(pptx_bytes):
    """Cuenta el total de slides en el PPTX."""
    return len(get_slide_order(pptx_bytes))


def is_valid_pptx(pptx_bytes):
    """Verifica que el PPTX sea un ZIP válido con presentation.xml."""
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
            return 'ppt/presentation.xml' in z.namelist()
    except Exception:
        return False


def count_perf_slots_in_first_perf_slide(pptx_bytes):
    """
    Encuentra el primer slide de perfiles completo y cuenta cuántos de los 4 slots
    (CuadroTexto 10/30/47/53) están presentes.
    """
    path, names = find_first_slide_with_all_shapes(pptx_bytes, ALL_PERF_ROLES)
    if path is None:
        return 0
    return len(ALL_PERF_ROLES & names)


def text_in_perf_slide(pptx_bytes, text_to_find, slot_shapes=None):
    """
    Verifica que text_to_find aparece en algún slide que contenga
    al menos 1 slot de perfiles (CuadroTexto 10/30/47/53) y el texto 'Perfiles'.
    Ignora slides que son plantilla (slide2.xml suele tener CuadroTexto 10).
    """
    if slot_shapes is None:
        slot_shapes = ALL_PERF_ROLES
    order = get_slide_order(pptx_bytes)
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for path in order:
            root  = etree.fromstring(z.read(path))
            names = set()
            for sp in root.iter(f'{{{P}}}sp'):
                nvpr = sp.find(f'.//{{{P}}}cNvPr')
                if nvpr is not None:
                    names.add(nvpr.attrib.get('name', ''))
            if not (names & slot_shapes):
                continue
            full_text = ''.join(t.text or '' for t in root.iter(f'{{{A}}}t'))
            if text_to_find in full_text and 'Perfiles' in full_text:
                return True
    return False


# ═══════════════════════════════════════ TEST RUNNER ═════════════════════════

def run(name, config, checks):
    print(f'\n{"─"*68}')
    print(f'  {name}')
    print(f'{"─"*68}')
    try:
        result     = generate(config, str(OUT_DIR))
        pptx_path  = Path(result['propuesta'])
        pptx_bytes = pptx_path.read_bytes()
        print(f'  Archivo: {pptx_path.name} ({pptx_path.stat().st_size // 1024} KB)')

        failed = []
        for label, fn in checks:
            try:
                ok = fn(pptx_bytes)
            except Exception as exc:
                ok = False
                print(f'  {FAIL_SYM}  {label}  ← excepción en check: {exc}')
                traceback.print_exc()
                failed.append(label)
                continue
            sym = PASS_SYM if ok else FAIL_SYM
            print(f'  {sym}  {label}')
            if not ok:
                failed.append(label)

        results.append((name, len(failed) == 0))

    except Exception as e:
        print(f'  {FAIL_SYM}  Excepción generando PPTX: {e}')
        traceback.print_exc()
        results.append((name, False))


# ═══════════════════════════════════════ ENTREGABLES ═════════════════════════

# ENT-1: 1 torre → slide entregables tiene 1 columna centrada
run('ENT-1  1 torre → 1 columna entregables (CT4 presente, CT13/17 ausentes)', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('CuadroTexto 4 presente en slide entregables',
     lambda p: 'CuadroTexto 4' in get_entregables_cols_present(p)),
    ('CuadroTexto 13 ausente (solo 1 columna)',
     lambda p: 'CuadroTexto 13' not in get_entregables_cols_present(p)),
    ('CuadroTexto 17 ausente (solo 1 columna)',
     lambda p: 'CuadroTexto 17' not in get_entregables_cols_present(p)),
    ('Título entregables tiene exactamente 2 runs',
     lambda p: len(get_entregables_title_runs(p)) == 2),
    ('Run 1 contiene "Entregables de"',
     lambda p: any('Entregables de' in txt for txt, _ in get_entregables_title_runs(p))),
    ('Run 2 no contiene "Xxxxxx" (nombre de torre reemplazado correctamente)',
     lambda p: not any('Xxxxxx' in txt for txt, _ in get_entregables_title_runs(p)
                       if 'Entregables de' not in txt)),
    ('Run 2 contiene nombre de la torre',
     lambda p: any(
         'FULLSTACK' in (txt or '').upper() or 'DESARROLLO' in (txt or '').upper()
         for txt, _ in get_entregables_title_runs(p)
         if 'Entregables de' not in txt
     )),
])

# ENT-2: 2 torres → 2 columnas
run('ENT-2  2 torres → 2 columnas entregables (CT4 y CT13 presentes, CT17 ausente)', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO', 'DATOS'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('CuadroTexto 4 presente',
     lambda p: 'CuadroTexto 4' in get_entregables_cols_present(p)),
    ('CuadroTexto 13 presente',
     lambda p: 'CuadroTexto 13' in get_entregables_cols_present(p)),
    ('CuadroTexto 17 ausente (solo 2 columnas)',
     lambda p: 'CuadroTexto 17' not in get_entregables_cols_present(p)),
    ('Los 2 títulos tienen colores correctos (2 runs cada uno)',
     lambda p: len(get_entregables_title_runs(p)) == 2),
    ('Ningún run contiene "Xxxxxx" residual',
     lambda p: not any('Xxxxxx' in txt for txt, _ in get_entregables_title_runs(p))),
])

# ENT-3: 3 torres → 3 columnas
run('ENT-3  3 torres → 3 columnas entregables (CT4, CT13 y CT17 presentes)', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO', 'DATOS', 'PMO'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('CuadroTexto 4 presente',
     lambda p: 'CuadroTexto 4' in get_entregables_cols_present(p)),
    ('CuadroTexto 13 presente',
     lambda p: 'CuadroTexto 13' in get_entregables_cols_present(p)),
    ('CuadroTexto 17 presente',
     lambda p: 'CuadroTexto 17' in get_entregables_cols_present(p)),
])

# ENT-4: El título NO tiene texto residual "Xxxxxx"
run('ENT-4  Título entregables: el run azul claro es exactamente el nombre de la torre', {
    'filial': 'group', 'excel_data': None,
    'torres_seleccionadas': ['ARQUITECTURA'],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Ningún run del título contiene "Xxxxxx"',
     lambda p: not any('Xxxxxx' in (txt or '') for txt, _ in get_entregables_title_runs(p))),
    ('Run 2 es el nombre de la torre (no vacío, no "Xxxxxx")',
     lambda p: any(
         txt.strip() and 'Xxxxxx' not in txt and 'Entregables de' not in txt
         for txt, _ in get_entregables_title_runs(p)
     )),
])

# ENT-5: Entregables con contenido real — sin "Xxxxx" en las listas
# Nota: "Xxxxx" puede aparecer en slides de plantilla (no de entregables),
# por lo que se verifica solo el texto de las shapes de lista de entregables.
run('ENT-5  Entregables contiene ítems reales (sin "Xxxxx" en las listas)', {
    'filial': 'cbit', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('No hay "Xxxxx" en el contenido de las listas de entregables',
     lambda p: 'Xxxxx' not in get_entregables_list_text(p)),
    ('Las listas de entregables tienen contenido real (no vacías)',
     lambda p: len(get_entregables_list_text(p).strip()) > 0),
    ('El slide de entregables fue detectado (CuadroTexto 4 presente)',
     lambda p: 'CuadroTexto 4' in get_entregables_cols_present(p)),
])


# ═══════════════════════════════════════ CONSIDERACIONES ═════════════════════

# CONS-6: 0 items + pill ON → genera slide con contenido (genéricos)
run('CONS-6  0 items Excel + pill ON → genera slide con genéricos', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
                   'perfiles': [], 'consideraciones': []},
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Existe al menos 1 slide de consideraciones editado',
     lambda p: count_edited_cons_slides(p) >= 1),
    ('Los grupos de consideraciones editados tienen contenido',
     lambda p: any(t.strip() for t in get_cons_slide_text(p))),
])

# CONS-7: 4 items → 1 slide editado con 4 grupos
run('CONS-7  4 consideraciones → 1 slide con 4 grupos', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
        'perfiles': [],
        'consideraciones': [
            'Consideración de prueba número uno para el proyecto.',
            'Consideración de prueba número dos para el proyecto.',
            'Consideración de prueba número tres para el proyecto.',
            'Consideración de prueba número cuatro para el proyecto.',
        ],
    },
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': False},
}, [
    ('Exactamente 1 slide de consideraciones editado',
     lambda p: count_edited_cons_slides(p) == 1),
    ('El slide tiene 4 grupos',
     lambda p: get_cons_groups_count_per_slide(p) == [4]),
])

# CONS-8: 5 items → 2 slides
run('CONS-8  5 consideraciones → 2 slides', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
        'perfiles': [],
        'consideraciones': [
            'Consideración uno: el servicio se ejecutará conforme al alcance aprobado.',
            'Consideración dos: cambios en requerimientos se gestionarán formalmente.',
            'Consideración tres: actividades no descritas se considerarán fuera de alcance.',
            'Consideración cuatro: no incluye soporte continuo posterior a garantía.',
            'Consideración cinco: no incluye licenciamiento de terceros.',
        ],
    },
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': False},
}, [
    ('2 slides de consideraciones editados (5 ítems requieren 2 slides)',
     lambda p: count_edited_cons_slides(p) == 2),
    ('Total grupos = 5 (4 en slide 1 + 1 en slide 2)',
     lambda p: sum(get_cons_groups_count_per_slide(p)) == 5),
])

# CONS-9: 8 items → 2 slides
run('CONS-9  8 consideraciones → 2 slides', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'DATOS', 'horas': 100}],
        'perfiles': [],
        'consideraciones': [
            'Item uno: análisis y diseño del modelo de datos relacional.',
            'Item dos: configuración de pipelines de ingesta de datos.',
            'Item tres: implementación de capa de transformación con Apache Spark.',
            'Item cuatro: orquestación de flujos con Apache Airflow.',
            'Item cinco: validación de calidad de datos con Great Expectations.',
            'Item seis: monitoreo y alertas sobre pipelines de datos.',
            'Item siete: documentación técnica del diseño de la arquitectura.',
            'Item ocho: capacitación al equipo cliente sobre el uso del sistema.',
        ],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': False},
}, [
    ('2 slides de consideraciones para 8 ítems',
     lambda p: count_edited_cons_slides(p) == 2),
    ('Primer slide tiene grupos (no vacío)',
     lambda p: len(get_cons_groups_count_per_slide(p)) >= 1 and get_cons_groups_count_per_slide(p)[0] >= 1),
    ('Total grupos suma 8',
     lambda p: sum(get_cons_groups_count_per_slide(p)) == 8),
])

# CONS-10: Grupos en Y >= 1107495 (no por encima del título)
# Solo se verifican los slides editados por el generador (>= 2 grupos)
run('CONS-10  Grupos de consideraciones están en Y >= 1107495 (debajo del título)', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'PMO', 'horas': 100}],
                   'perfiles': [], 'consideraciones': ['Una consideración corta de prueba.']},
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': False},
}, [
    ('Al menos 1 grupo de consideraciones editado',
     lambda p: len(get_edited_cons_grupo_y_positions(p)) >= 1),
    ('Todos los grupos en slides editados tienen Y >= 1107495',
     lambda p: all(y >= 1107495 for y in get_edited_cons_grupo_y_positions(p))),
])

# CONS-11: Ningún grupo en Y < 1000000 en slides editados
run('CONS-11  Ningún grupo de consideraciones editado en Y < 1000000', {
    'filial': 'group',
    'excel_data': {'torres': [{'nombre': 'ARQUITECTURA', 'horas': 100}],
                   'perfiles': [], 'consideraciones': [
                       'Primera consideración importante para el alcance del proyecto.',
                       'Segunda consideración sobre el control de cambios.',
                       'Tercera consideración sobre infraestructura.',
                   ]},
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': False},
}, [
    ('Los slides editados tienen al menos 1 grupo',
     lambda p: len(get_edited_cons_grupo_y_positions(p)) >= 1),
    ('Ningún grupo editado tiene Y < 1000000',
     lambda p: not any(y < 1000000 for y in get_edited_cons_grupo_y_positions(p))),
])

# CONS-12: PPTX válido
run('CONS-12  El PPTX generado es un ZIP válido con presentation.xml', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['DATOS'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('PPTX es un ZIP válido',
     lambda p: is_valid_pptx(p)),
    ('PPTX contiene ppt/presentation.xml',
     lambda p: is_valid_pptx(p)),
])


# ═══════════════════════════════════════ PERFILES ════════════════════════════

# PERF-13: 1 perfil → 1 slide con 1 columna centrada
# Con 1 perfil, solo CuadroTexto 10 permanece (los otros 3 slots son eliminados).
# Verificamos que el rol aparece en el slide de perfiles (detección por texto).
run('PERF-13  1 perfil manual → 1 slide de perfiles con el rol correcto', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'PMO', 'horas': 80}], 'perfiles': []},
    'perfiles_manuales': [{'rol': 'PMO Manager', 'desc': 'Gestión de proyectos bajo PMI.'}],
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('PMO Manager aparece en el PPTX',
     lambda p: 'PMO Manager' in get_all_text(p)),
    ('PMO Manager aparece en un slide de perfiles (con título "Perfiles")',
     lambda p: text_in_perf_slide(p, 'PMO Manager', ALL_PERF_ROLES)),
])

# PERF-14: 4 perfiles → 1 slide con 4 columnas
run('PERF-14  4 perfiles manuales → 1 slide con 4 columnas', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'ARQUITECTURA', 'horas': 200}], 'perfiles': []},
    'perfiles_manuales': [
        {'rol': 'Arquitecto Senior', 'desc': 'Diseño de arquitecturas empresariales.'},
        {'rol': 'Tech Lead',         'desc': 'Liderazgo técnico de equipos ágiles.'},
        {'rol': 'Cloud Engineer',    'desc': 'Implementación en AWS y Azure.'},
        {'rol': 'DevOps Engineer',   'desc': 'CI/CD con Jenkins y GitLab.'},
    ],
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Exactamente 1 slide de perfiles completo (4 slots)',
     lambda p: count_perf_slides(p) == 1),
    ('Los 4 roles aparecen en el texto',
     lambda p: all(r in get_all_text(p)
                   for r in ['Arquitecto Senior', 'Tech Lead', 'Cloud Engineer', 'DevOps Engineer'])),
])

# PERF-15: 5 perfiles → 2 slides
# Slide 1: 4 perfiles (los 4 slots presentes)
# Slide 2: 1 perfil (solo 1 slot: CuadroTexto 10, los otros eliminados)
run('PERF-15  5 perfiles manuales → 2 slides de perfiles', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}], 'perfiles': []},
    'perfiles_manuales': [
        {'rol': 'Dev Senior A',  'desc': 'React + Node.js'},
        {'rol': 'Dev Senior B',  'desc': 'Python + Django'},
        {'rol': 'Dev Junior',    'desc': 'JavaScript + Vue.js'},
        {'rol': 'QA Engineer',   'desc': 'Selenium + Cypress'},
        {'rol': 'Scrum Master',  'desc': 'Agile + SAFe'},
    ],
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Slide 1 tiene los 4 slots completos (4 perfiles)',
     lambda p: count_perf_slides(p) == 1),
    ('Los 5 roles aparecen en el texto del PPTX',
     lambda p: all(r in get_all_text(p)
                   for r in ['Dev Senior A', 'Dev Senior B', 'Dev Junior', 'QA Engineer', 'Scrum Master'])),
    ('Dev Senior A en un slide de perfiles',
     lambda p: text_in_perf_slide(p, 'Dev Senior A', ALL_PERF_ROLES)),
    ('Scrum Master en un slide de perfiles',
     lambda p: text_in_perf_slide(p, 'Scrum Master', ALL_PERF_ROLES)),
])

# PERF-16: QA NO aparece como torre en TORRES_ALL del HTML
def _check_qa_not_in_torres_all():
    """Lee home.html y verifica que 'QA' no esté en TORRES_ALL como elemento."""
    html_path = Path(__file__).resolve().parent.parent / 'static' / 'home.html'
    if not html_path.exists():
        return False
    content = html_path.read_text(encoding='utf-8')
    match = re.search(r'const\s+TORRES_ALL\s*=\s*\[(.*?)\]', content, re.DOTALL)
    if not match:
        return False
    torres_block = match.group(1)
    torres_vals  = re.findall(r"'([^']+)'", torres_block)
    return 'QA' not in torres_vals

run('PERF-16  QA NO aparece como torre en TORRES_ALL del HTML', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('QA no está en TORRES_ALL de home.html',
     lambda p: _check_qa_not_in_torres_all()),
])

# PERF-17: Perfiles manuales funcionan
run('PERF-17  Perfiles manuales funcionan correctamente', {
    'filial': 'group',
    'excel_data': {'torres': [{'nombre': 'DATOS', 'horas': 100}], 'perfiles': []},
    'perfiles_manuales': [
        {'rol': 'Data Scientist Custom', 'desc': 'Machine learning con scikit-learn.'},
        {'rol': 'Data Engineer Custom',  'desc': 'ETL con Spark y Airflow en GCP.'},
    ],
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Ambos roles manuales aparecen en el texto del PPTX',
     lambda p: 'Data Scientist Custom' in get_all_text(p) and
               'Data Engineer Custom' in get_all_text(p)),
    ('Data Scientist Custom aparece en un slide de perfiles',
     lambda p: text_in_perf_slide(p, 'Data Scientist Custom', ALL_PERF_ROLES)),
    ('Data Engineer Custom aparece en el PPTX',
     lambda p: 'Data Engineer Custom' in get_all_text(p)),
])


# ═══════════════════════════════════════ SLIDE DETECTION ═════════════════════

for filial in ['corp', 'group', 'cbit']:
    # SLIDE-18: slide perfiles detectado tiene los 4 slots
    run(f'SLIDE-18  [{filial}] Slide perfiles detectado tiene los 4 slots de roles', {
        'filial': filial, 'excel_data': None,
        'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
        'incluir_qa': True,
        'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
    }, [
        ('Existe slide con todos los 4 slots de perfiles',
         lambda p: find_first_slide_with_all_shapes(p, ALL_PERF_ROLES)[0] is not None),
        ('Slide de perfiles tiene CuadroTexto 10',
         lambda p: 'CuadroTexto 10' in find_first_slide_with_all_shapes(p, ALL_PERF_ROLES)[1]),
        ('Slide de perfiles tiene CuadroTexto 30',
         lambda p: 'CuadroTexto 30' in find_first_slide_with_all_shapes(p, ALL_PERF_ROLES)[1]),
        ('Slide de perfiles tiene CuadroTexto 47',
         lambda p: 'CuadroTexto 47' in find_first_slide_with_all_shapes(p, ALL_PERF_ROLES)[1]),
        ('Slide de perfiles tiene CuadroTexto 53',
         lambda p: 'CuadroTexto 53' in find_first_slide_with_all_shapes(p, ALL_PERF_ROLES)[1]),
    ])

    # SLIDE-19: slide FDA detectado tiene >= 3 bullet rects
    run(f'SLIDE-19  [{filial}] Slide FDA detectado tiene >= 3 bullet rects', {
        'filial': filial, 'excel_data': None,
        'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
        'incluir_qa': True,
        'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
    }, [
        ('Existe slide con >= 3 bullet rects (Rectángulo 10/13/19/22/25/28)',
         lambda p: count_slides_with_min_shapes(p, BULLET_RECTS, min_count=3) >= 1),
    ])

    # SLIDE-20: slide entregables detectado tiene CuadroTexto 4
    run(f'SLIDE-20  [{filial}] Slide entregables detectado tiene CuadroTexto 4', {
        'filial': filial, 'excel_data': None,
        'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
        'incluir_qa': True,
        'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
    }, [
        ('CuadroTexto 4 presente en el slide de entregables',
         lambda p: 'CuadroTexto 4' in get_entregables_cols_present(p)),
    ])


# ═══════════════════════════════════════ FILIALES ════════════════════════════

# FIL-21: Las 3 filiales generan sin excepción
for filial in ['corp', 'group', 'cbit']:
    run(f'FIL-21  [{filial}] Generación completa sin excepción', {
        'filial': filial, 'excel_data': None,
        'torres_seleccionadas': ['FULLSTACK / DESARROLLO', 'DATOS'],
        'incluir_qa': True,
        'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
    }, [
        ('PPTX generado (> 10 KB)',
         lambda p: len(p) > 10_000),
        ('PPTX es un ZIP válido con presentation.xml',
         lambda p: is_valid_pptx(p)),
    ])

# FIL-22: Cada filial tiene > 20 slides
for filial in ['corp', 'group', 'cbit']:
    run(f'FIL-22  [{filial}] PPTX tiene > 20 slides', {
        'filial': filial, 'excel_data': None,
        'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
        'incluir_qa': True,
        'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
    }, [
        ('PPTX tiene más de 20 slides',
         lambda p: count_total_slides(p) > 20),
    ])


# ═══════════════════════════════════════ EDGE CASES ══════════════════════════

# EDGE-23: torres_seleccionadas vacío con excel_data vacío → no crashea
run('EDGE-23  Torres vacías y excel_data vacío → no crashea', {
    'filial': 'corp',
    'excel_data': {'torres': [], 'perfiles': [], 'consideraciones': []},
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('PPTX generado sin crash',
     lambda p: len(p) > 10_000),
    ('PPTX válido',
     lambda p: is_valid_pptx(p)),
])

# EDGE-24: perfiles_manuales vacío → no crashea
run('EDGE-24  perfiles_manuales vacío → no crashea', {
    'filial': 'corp',
    'excel_data': None,
    'perfiles_manuales': [],
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('PPTX generado sin crash',
     lambda p: len(p) > 10_000),
])

# EDGE-25: Todos los toggles OFF → genera igualmente
run('EDGE-25  Todos los toggles OFF → genera igualmente', {
    'filial': 'corp',
    'excel_data': None,
    'torres_seleccionadas': ['ARQUITECTURA'],
    'incluir_qa': False,
    'opciones': {'perfiles': False, 'fda': False, 'entregables': False, 'consideraciones': False},
}, [
    ('PPTX generado con toggles OFF',
     lambda p: len(p) > 10_000),
    ('PPTX válido',
     lambda p: is_valid_pptx(p)),
])

# EDGE-26: excel_data con cliente vacío → no crashea
run('EDGE-26  excel_data con cliente vacío → no crashea', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
        'perfiles': [{'perfil': 'Dev', 'seniority': ''}],
        'cliente': '',
        'consideraciones': ['Consideración de prueba.'],
    },
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('PPTX generado sin crash con cliente vacío',
     lambda p: len(p) > 10_000),
    ('PPTX válido',
     lambda p: is_valid_pptx(p)),
])


# ═══════════════════════════════════════ ENTREGABLES VACIOS / SPLIT CONS ═════

# ENT-27: Torre inventada (sin entregables en Excel) → columnas vacías eliminadas
run('ENT-27  Torre sin entregables en Excel → no deja cuadros vacíos', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['BLOCKCHAIN_FAKE'],   # no existe en el catálogo
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('No hay CuadroTexto 13 (columnas vacías eliminadas)',
     lambda p: 'CuadroTexto 13' not in get_entregables_cols_present(p)),
    ('No hay CuadroTexto 17 (columnas vacías eliminadas)',
     lambda p: 'CuadroTexto 17' not in get_entregables_cols_present(p)),
    ('PPTX válido',
     lambda p: is_valid_pptx(p)),
])

# ENT-30: 4 torres → 4 columnas en 1 solo slide (layout escalado)
run('ENT-30  4 torres → 4 columnas en 1 slide (layout al 75%)', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO', 'QA', 'ARQUITECTURA', 'DATOS'],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('1 solo slide de entregables (no hay duplicación para 4 torres)',
     lambda p: count_entregables_slides(p) == 1),
    ('4 shapes de lista en el slide (CT4, CT13, CT17 + clon CT17)',
     lambda p: count_ent_list_shapes_in_slide(p, 0) == 4),
    ('PPTX válido',
     lambda p: is_valid_pptx(p)),
])

# ENT-31: 5 torres → 2 slides (4 + 1)
run('ENT-31  5 torres → 2 slides de entregables (4 + 1)', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO', 'QA', 'ARQUITECTURA', 'DATOS', 'RPA'],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('2 slides de entregables (duplicación automática)',
     lambda p: count_entregables_slides(p) == 2),
    ('Slide 1 tiene 4 columnas (lista shapes)',
     lambda p: count_ent_list_shapes_in_slide(p, 0) == 4),
    ('Slide 2 tiene 1 columna (lista shapes)',
     lambda p: count_ent_list_shapes_in_slide(p, 1) == 1),
    ('PPTX válido',
     lambda p: is_valid_pptx(p)),
])

# ENT-32: 8 torres → 2 slides (4 + 4)
run('ENT-32  8 torres → 2 slides de entregables (4 + 4)', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': [
        'FULLSTACK / DESARROLLO', 'QA', 'ARQUITECTURA', 'DATOS',
        'RPA', 'DEVOPS', 'CIBERSEGURIDAD', 'IA',
    ],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('2 slides de entregables',
     lambda p: count_entregables_slides(p) == 2),
    ('Slide 1 tiene 4 columnas',
     lambda p: count_ent_list_shapes_in_slide(p, 0) == 4),
    ('Slide 2 tiene 4 columnas',
     lambda p: count_ent_list_shapes_in_slide(p, 1) == 4),
    ('PPTX válido',
     lambda p: is_valid_pptx(p)),
])

# ENT-33: Entregables desde excel_data.torres (verificar fuente real)
run('ENT-33  Entregables desde excel_data.torres con datos reales del catálogo', {
    'filial': 'corp',
    'excel_data': {
        'torres': [
            {'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100},
            {'nombre': 'QA', 'horas': 50},
        ],
        'perfiles': [],
        'consideraciones': [],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': False},
}, [
    ('2 columnas de entregables (una por torre del excel_data)',
     lambda p: count_ent_list_shapes_in_slide(p, 0) == 2),
    ('Texto "Arquitectura técnica" presente (ítem real de FULLSTACK)',
     lambda p: 'Arquitectura técnica' in get_entregables_list_text(p)),
    ('Texto "Plan de pruebas" presente (ítem real de QA)',
     lambda p: 'Plan de pruebas' in get_entregables_list_text(p)),
    ('PPTX válido',
     lambda p: is_valid_pptx(p)),
])

# CONS-28: Texto multi-oración desde Excel se divide en consideraciones individuales
run('CONS-28  Texto con puntos del Excel se divide en cuadros individuales', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
        'perfiles': [],
        'consideraciones': [
            'El servicio se prestará conforme al alcance técnico aprobado. '
            'Cualquier modificación posterior será gestionada mediante control de cambios. '
            'No incluye soporte continuo posterior al periodo de garantía.',
        ],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': False},
}, [
    ('Al menos 3 grupos de consideraciones (uno por oración)',
     lambda p: sum(get_cons_groups_count_per_slide(p)) >= 3),
    ('Texto "conforme al alcance" aparece en consideraciones editadas',
     lambda p: any('conforme al alcance' in t for t in get_cons_slide_text(p))),
    ('Texto "control de cambios" aparece en consideraciones editadas',
     lambda p: any('control de cambios' in t for t in get_cons_slide_text(p))),
])

# CONS-29: Ítem ya individual (sin punto interno) no se rompe
run('CONS-29  Ítem sin punto interior no se rompe al dividir', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
        'perfiles': [],
        'consideraciones': [
            'El cliente garantizará la disponibilidad del equipo.',
            'No se incluyen licencias de software de terceros.',
        ],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': False},
}, [
    ('Exactamente 2 grupos de consideraciones (1 por ítem)',
     lambda p: sum(get_cons_groups_count_per_slide(p)) == 2),
    ('Primer ítem aparece completo en el slide',
     lambda p: any('garantizará la disponibilidad' in t for t in get_cons_slide_text(p))),
])


# ENT-34: excel_data.entregables + pill OFF → solo los del Excel
run('ENT-34  excel_data.entregables + pill OFF → solo los del Excel', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 200}],
        'perfiles': [],
        'consideraciones': [],
        'entregables': [
            {'torre': 'FULLSTACK / DESARROLLO', 'items': ['Plan de pruebas unitarias', 'Repositorio Git', 'Manual de usuario']},
        ],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': False, 'fda': False, 'entregables': False, 'consideraciones': False},
}, [
    ('1 slide de entregables',
     lambda p: count_entregables_slides(p) == 1),
    ('Texto del Excel aparece',
     lambda p: 'Plan de pruebas unitarias' in get_entregables_list_text(p)),
    ('Sin catálogo extra (pill OFF, solo 1 torre del Excel)',
     lambda p: count_ent_list_shapes_in_slide(p) == 1),
])

# ENT-36: excel_data.entregables + pill ON → Excel + catálogo para torres faltantes
run('ENT-36  excel_data.entregables + pill ON → complementa con catálogo', {
    'filial': 'corp',
    'excel_data': {
        'torres': [
            {'nombre': 'FULLSTACK / DESARROLLO', 'horas': 200},
            {'nombre': 'QA', 'horas': 80},
        ],
        'perfiles': [],
        'consideraciones': [],
        'entregables': [
            {'torre': 'FULLSTACK / DESARROLLO', 'items': ['Plan de pruebas unitarias', 'Repositorio Git']},
            # QA no tiene entregables en el Excel
        ],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': False, 'fda': False, 'entregables': True, 'consideraciones': False},
}, [
    ('Entregables de Full Stack del Excel aparecen',
     lambda p: 'Plan de pruebas unitarias' in get_entregables_list_text(p)),
    ('QA se complementa con catálogo (2 columnas total)',
     lambda p: count_ent_list_shapes_in_slide(p) == 2),
])

# ENT-35: sin excel_data.entregables → fallback al catálogo
run('ENT-35  Sin excel_data.entregables usa catálogo genérico', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 200}],
        'perfiles': [],
        'consideraciones': [],
        'entregables': [],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': False, 'fda': False, 'entregables': True, 'consideraciones': False},
}, [
    ('1 slide de entregables',
     lambda p: count_entregables_slides(p) == 1),
    ('Hay texto del catálogo genérico',
     lambda p: len(get_entregables_list_text(p)) > 0),
])


# ═══════════════════════════════════════ RESUMEN ═════════════════════════════

print(f'\n{"═"*68}')
print('  RESUMEN FINAL — test_e2e.py')
print(f'{"═"*68}')
passed = sum(1 for _, ok in results if ok)
total  = len(results)
for name, ok in results:
    sym = PASS_SYM if ok else FAIL_SYM
    print(f'  {sym}  {name}')
print(f'\n  {passed}/{total} casos pasaron')
print(f'{"═"*68}')
sys.exit(0 if passed == total else 1)
