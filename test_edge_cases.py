"""
test_edge_cases.py — pruebas de casos límite y edge cases.
Corre: python3 test_edge_cases.py
"""

import sys, traceback, zipfile, io, json
from pathlib import Path
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent / 'periferia_v2'))
from generators import generate
from generators.fda_perfiles import _even_chunks, _find_desc_in_catalog, _load_generales

OUT_DIR = Path('/tmp/periferia_edge')
OUT_DIR.mkdir(parents=True, exist_ok=True)

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

PASS = '\033[92m✓ PASS\033[0m'
FAIL = '\033[91m✗ FAIL\033[0m'
results = []


# ── helpers ──────────────────────────────────────────────────────────────────

def all_text(pptx_bytes):
    parts = []
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for n in sorted(z.namelist()):
            if n.startswith('ppt/slides/slide') and n.endswith('.xml'):
                root = etree.fromstring(z.read(n))
                parts.append(''.join(t.text or '' for t in root.iter(f'{{{A}}}t')))
    return ' '.join(parts)

def count_slides(pptx_bytes):
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        return sum(1 for n in z.namelist()
                   if n.startswith('ppt/slides/slide') and n.endswith('.xml'))

def fda_slide_count(pptx_bytes):
    """
    Cuenta cuántos slides son slides FDA reales.
    Un slide FDA tiene Rectángulo 10 (bullets) Y CuadroTexto 32 (panel QA).
    Otros slides del template también tienen Rectángulo 10 pero no son FDA.
    """
    count = 0
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for n in sorted(z.namelist()):
            if not (n.startswith('ppt/slides/slide') and n.endswith('.xml')):
                continue
            root = etree.fromstring(z.read(n))
            names = {sp.find(f'.//{{{P}}}cNvPr').attrib.get('name','')
                     for sp in root.iter(f'{{{P}}}sp')
                     if sp.find(f'.//{{{P}}}cNvPr') is not None}
            if 'Rectángulo 10' in names and 'CuadroTexto 32' in names:
                count += 1
    return count

def perf_slide_count(pptx_bytes):
    """Cuenta cuántos slides tienen el marcador de perfiles (CuadroTexto 10)."""
    count = 0
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for n in sorted(z.namelist()):
            if not (n.startswith('ppt/slides/slide') and n.endswith('.xml')):
                continue
            root = etree.fromstring(z.read(n))
            names = {sp.find(f'.//{{{P}}}cNvPr').attrib.get('name','')
                     for sp in root.iter(f'{{{P}}}sp')
                     if sp.find(f'.//{{{P}}}cNvPr') is not None}
            if 'CuadroTexto 10' in names:
                count += 1
    return count

def run(name, config_or_fn, checks):
    print(f'\n{"─"*64}')
    print(f'  {name}')
    print(f'{"─"*64}')
    try:
        if callable(config_or_fn):
            # Caso que espera excepción
            config_or_fn()
            print(f'  {FAIL}  Se esperaba excepción pero no ocurrió')
            results.append((name, False))
            return

        result     = generate(config_or_fn, str(OUT_DIR))
        pptx_bytes = Path(result['propuesta']).read_bytes()
        txt        = all_text(pptx_bytes)
        print(f'  Slides totales: {count_slides(pptx_bytes)}  |  FDA slides: {fda_slide_count(pptx_bytes)}  |  Perf slides: {perf_slide_count(pptx_bytes)}')

        failed = []
        for label, cond in checks:
            try:
                ok  = cond(txt, pptx_bytes)
            except Exception as e:
                ok = False
                label = f'{label} [EXCEPCIÓN: {e}]'
            sym = PASS if ok else FAIL
            print(f'  {sym}  {label}')
            if not ok:
                failed.append(label)

        results.append((name, len(failed) == 0))

    except Exception as e:
        print(f'  {FAIL}  Excepción inesperada: {e}')
        traceback.print_exc()
        results.append((name, False))


# ════════════════════════ UNIT: _even_chunks ═════════════════════════════════

print('\n════ UNIT: _even_chunks ════')
tests_unit = [
    ([], 8,   [[]]),
    ([1], 8,  [[1]]),
    ([1,2,3,4,5,6], 6, [[1,2,3,4,5,6]]),
    ([1,2,3,4,5,6,7], 8, [[1,2,3,4,5,6,7]]),    # 7 cabe en 1 slide de 8
    ([1,2,3,4,5,6,7,8,9], 8, [[5,6,7,8,9], [1,2,3,4]]),  # 2 slides: 5+4... el orden puede variar
    (list(range(17)), 8, None),  # 3 slides ceil(17/8)=3
]
for items, mp, expected in tests_unit:
    result = _even_chunks(items, mp)
    total_in  = sum(len(c) for c in result)
    total_out = len(items)
    max_chunk = max(len(c) for c in result)
    n_chunks  = len(result)
    ok = (total_in == total_out) and (max_chunk <= mp)
    sym = PASS if ok else FAIL
    print(f'  {sym}  _even_chunks({len(items)} items, max={mp}) → {n_chunks} chunks, max_chunk={max_chunk}, total={total_in}/{total_out}')


# ════════════════════════ UNIT: _find_desc_in_catalog ════════════════════════

print('\n════ UNIT: _find_desc_in_catalog ════')
try:
    _, perf_db = _load_generales()
    # Búsqueda exacta de un rol que SÍ existe
    any_role = next((p['rol'] for profs in perf_db.values() for p in profs), None)
    if any_role:
        desc = _find_desc_in_catalog(any_role, perf_db)
        sym = PASS if desc else FAIL
        print(f'  {sym}  Rol existente "{any_role[:30]}" → desc encontrada: {bool(desc)}')
    # Búsqueda de rol inexistente
    desc_none = _find_desc_in_catalog('RolQueNoExisteJamas12345', perf_db)
    sym = PASS if desc_none == '' else FAIL
    print(f'  {sym}  Rol inexistente → desc vacía: "{desc_none}"')
    # Búsqueda case-insensitive/acento
    if any_role:
        desc_low = _find_desc_in_catalog(any_role.lower(), perf_db)
        sym = PASS if desc_low else FAIL
        print(f'  {sym}  Búsqueda lowercase "{any_role[:20].lower()}" → desc encontrada: {bool(desc_low)}')
except Exception as e:
    print(f'  {FAIL}  _find_desc_in_catalog lanzó excepción: {e}')


# ════════════════════ E2E: CASES ══════════════════════════════════════════════

# ── EDGE-1: perfiles_manuales override ───────────────────────────────────────
run('EDGE-1  perfiles_manuales tiene prioridad sobre catálogo y Excel', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
                   'perfiles': [{'perfil': 'Tech Lead', 'seniority': ''}]},
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': False, 'consideraciones': False},
    'perfiles_manuales': [
        {'rol': 'Mi Rol Manual Único', 'desc': 'Descripción manual de prueba.'},
    ],
}, [
    ('Rol manual aparece en el PPTX',
     lambda t, b: 'Mi Rol Manual Único' in t),
    ('Rol del Excel NO aparece (override completo)',
     lambda t, b: 'Tech Lead' not in t),
    ('Descripción manual aparece',
     lambda t, b: 'Descripción manual de prueba' in t),
])

# ── EDGE-2: perfiles_manuales vacío → cae al catálogo ────────────────────────
run('EDGE-2  perfiles_manuales=[] → usa catálogo normalmente', {
    'filial': 'corp',
    'excel_data': None,
    'torres_seleccionadas': ['DATOS'],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': False, 'consideraciones': False},
    'perfiles_manuales': [],
}, [
    ('Genera sin crash',
     lambda t, b: len(b) > 10000),
    ('Al menos 1 perfil renderizado',
     lambda t, b: perf_slide_count(b) >= 1),
])

# ── EDGE-3: FDA multi-slide desde excel_data.fda (9 items → 2 slides) ─────────
run('EDGE-3  FDA Excel 9 items → 2 slides FDA', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
        'perfiles': [],
        'fda': [
            'Item FDA 1.', 'Item FDA 2.', 'Item FDA 3.', 'Item FDA 4.',
            'Item FDA 5.', 'Item FDA 6.', 'Item FDA 7.', 'Item FDA 8.',
            'Item FDA 9.',
        ],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': False, 'entregables': False, 'consideraciones': False},
}, [
    ('Genera sin crash',
     lambda t, b: len(b) > 10000),
    ('Hay 2 slides FDA',
     lambda t, b: fda_slide_count(b) == 2),
    ('Item FDA 1 aparece',
     lambda t, b: 'Item FDA 1' in t),
    ('Item FDA 9 aparece',
     lambda t, b: 'Item FDA 9' in t),
])

# ── EDGE-4: FDA 7 items → 2 slides (max_per=6: [4,3]) ────────────────────────
run('EDGE-4  FDA 7 items → 2 slides (4+3, max_per=6)', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
        'perfiles': [],
        'fda': ['Item %d.' % i for i in range(1, 8)],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': False, 'entregables': False, 'consideraciones': False},
}, [
    ('Genera sin crash',
     lambda t, b: len(b) > 10000),
    ('2 slides FDA (7 items / max 6)',
     lambda t, b: fda_slide_count(b) == 2),
    ('Los 7 items aparecen',
     lambda t, b: all(f'Item {i}' in t for i in range(1, 8))),
])

# ── EDGE-5: FDA 8 items → 2 slides (max_per=6: [4,4]) ───────────────────────
run('EDGE-5  FDA 8 items → 2 slides (4+4, max_per=6)', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'DATOS', 'horas': 50}],
        'perfiles': [],
        'fda': ['Punto %d.' % i for i in range(1, 9)],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': False, 'entregables': False, 'consideraciones': False},
}, [
    ('Genera sin crash',      lambda t, b: len(b) > 10000),
    ('2 slides FDA (8 items / max 6)', lambda t, b: fda_slide_count(b) == 2),
    ('Los 8 puntos aparecen',   lambda t, b: all(f'Punto {i}' in t for i in range(1, 9))),
])

# ── EDGE-6: FDA 16 items → 3 slides (max_per=6: [6,6,4]) ────────────────────
run('EDGE-6  FDA 16 items → 3 slides (6+6+4, max_per=6)', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'DATOS', 'horas': 50}],
        'perfiles': [],
        'fda': ['FDA %d.' % i for i in range(1, 17)],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': False, 'entregables': False, 'consideraciones': False},
}, [
    ('Genera sin crash',    lambda t, b: len(b) > 10000),
    ('3 slides FDA',        lambda t, b: fda_slide_count(b) == 3),
    ('FDA 1 aparece',       lambda t, b: 'FDA 1' in t),
    ('FDA 16 aparece',      lambda t, b: 'FDA 16' in t),
])

# ── EDGE-7: torres_seleccionadas vacío y sin excel → catálogo vacío ──────────
run('EDGE-7  Sin torres y sin excel → no crashea, usa cláusula general FDA', {
    'filial': 'corp',
    'excel_data': None,
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': False, 'consideraciones': False},
}, [
    ('Genera sin crash',
     lambda t, b: len(b) > 10000),
    ('FDA usa cláusula general (sin torres)',
     lambda t, b: 'alcance técnico aprobado' in t),
])

# ── EDGE-8: Torre desconocida → cae a cláusula general ───────────────────────
run('EDGE-8  Torre desconocida en excel → cláusula general FDA', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'TORRE_INVENTADA_XYZ', 'horas': 100}],
                   'perfiles': []},
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': False, 'entregables': False, 'consideraciones': False},
}, [
    ('Genera sin crash',
     lambda t, b: len(b) > 10000),
    ('FDA tiene contenido (fallback a genéricos)',
     lambda t, b: fda_slide_count(b) >= 1),
])

# ── EDGE-9: QA como única torre activa ───────────────────────────────────────
run('EDGE-9  QA como única torre (excel vacío) → incluir_qa derivado', {
    'filial': 'corp',
    'excel_data': None,
    'torres_seleccionadas': ['QA'],
    'torres_qa': {'QA': True},
    'opciones': {'perfiles': True, 'fda': True, 'entregables': False, 'consideraciones': False},
}, [
    ('Genera sin crash',
     lambda t, b: len(b) > 10000),
    ('Al menos 1 perfil slide',
     lambda t, b: perf_slide_count(b) >= 1),
])

# ── EDGE-10: perfiles_manuales con caracteres especiales HTML ────────────────
run('EDGE-10  perfiles_manuales con caracteres especiales & < >', {
    'filial': 'corp',
    'excel_data': None,
    'torres_seleccionadas': ['DATOS'],
    'incluir_qa': False,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': False, 'consideraciones': False},
    'perfiles_manuales': [
        {'rol': 'Dev & Ops <Senior>', 'desc': 'Maneja CI/CD > 10 pipelines & más.'},
    ],
}, [
    ('Genera sin crash (sin excepción XML)',
     lambda t, b: len(b) > 10000),
    ('Rol con especiales aparece',
     lambda t, b: 'Dev & Ops' in t or 'Dev &amp; Ops' in t or 'Dev' in t),
])

# ── EDGE-11: Muchos perfiles manuales (más de 4, paginación) ─────────────────
run('EDGE-11  9 perfiles manuales → 3 slides de perfiles (4+4+1)', {
    'filial': 'corp',
    'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
    'incluir_qa': False,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': False, 'consideraciones': False},
    'perfiles_manuales': [{'rol': f'Rol Manual {i}', 'desc': f'Desc {i}.'} for i in range(1, 10)],
}, [
    ('Genera sin crash',
     lambda t, b: len(b) > 10000),
    ('3 slides de perfiles',
     lambda t, b: perf_slide_count(b) == 3),
    ('Rol Manual 1 aparece',   lambda t, b: 'Rol Manual 1' in t),
    ('Rol Manual 9 aparece',   lambda t, b: 'Rol Manual 9' in t),
])

# ── EDGE-12: Filial inválida → FileNotFoundError ─────────────────────────────
print(f'\n{"─"*64}')
print('  EDGE-12  Filial inválida → lanza FileNotFoundError')
print(f'{"─"*64}')
try:
    generate({
        'filial': 'FILIAL_INEXISTENTE',
        'excel_data': None,
        'torres_seleccionadas': ['DATOS'],
        'incluir_qa': False,
        'opciones': {'perfiles': True, 'fda': True, 'entregables': False, 'consideraciones': False},
    }, str(OUT_DIR))
    print(f'  {FAIL}  No lanzó excepción (debería haber fallado)')
    results.append(('EDGE-12  Filial inválida → FileNotFoundError', False))
except FileNotFoundError as e:
    print(f'  {PASS}  FileNotFoundError: {e}')
    results.append(('EDGE-12  Filial inválida → FileNotFoundError', True))
except Exception as e:
    print(f'  {FAIL}  Excepción inesperada ({type(e).__name__}): {e}')
    results.append(('EDGE-12  Filial inválida → FileNotFoundError', False))

# ── EDGE-13: excel_data con fda vacío [] vs None ─────────────────────────────
run('EDGE-13  excel_data.fda=[] → usa catálogo FDA, no crashea', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
                   'perfiles': [], 'fda': []},
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': False, 'entregables': False, 'consideraciones': False},
}, [
    ('Genera sin crash',
     lambda t, b: len(b) > 10000),
    ('FDA tiene contenido del catálogo FULLSTACK',
     lambda t, b: fda_slide_count(b) >= 1),
])

# ── EDGE-14: Excel con un solo perfil sin rol (campo vacío) ──────────────────
run('EDGE-14  excel_data.perfiles con entradas sin campo perfil → ignora vacíos', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'DATOS', 'horas': 100}],
                   'perfiles': [
                       {'perfil': '', 'seniority': ''},
                       {'perfil': None, 'seniority': ''},
                       {'perfil': 'Data Engineer', 'seniority': ''},
                   ]},
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': False, 'consideraciones': False},
}, [
    ('Genera sin crash',
     lambda t, b: len(b) > 10000),
    ('Solo "Data Engineer" renderizado',
     lambda t, b: 'Data Engineer' in t),
    ('1 slide de perfiles',
     lambda t, b: perf_slide_count(b) == 1),
])

# ── EDGE-15: Todas las torres (estrés máximo) ────────────────────────────────
TODAS_TORRES = [
    'FULLSTACK / DESARROLLO','QA','ARQUITECTURA','DATOS','RPA',
    'DEVOPS','CIBERSEGURIDAD','IA','SAP','PMO','MOBILE','PORTALES','INTEGRACIÓN'
]
run('EDGE-15  Todas las torres seleccionadas → no crashea', {
    'filial': 'corp',
    'excel_data': None,
    'torres_seleccionadas': TODAS_TORRES,
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': False, 'consideraciones': False},
}, [
    ('Genera sin crash',
     lambda t, b: len(b) > 10000),
    ('Hay slides de perfiles',
     lambda t, b: perf_slide_count(b) >= 1),
    ('Hay slides FDA',
     lambda t, b: fda_slide_count(b) >= 1),
])


# ════════════════════════════ RESUMEN ═════════════════════════════════════════

print(f'\n{"═"*64}')
print('  RESUMEN')
print(f'{"═"*64}')
passed = sum(1 for _, ok in results if ok)
for name, ok in results:
    sym = PASS if ok else FAIL
    print(f'  {sym}  {name}')
print(f'\n  {passed}/{len(results)} casos pasaron')
print(f'{"═"*64}')
