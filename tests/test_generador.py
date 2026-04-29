"""
test_generador.py — pruebas exhaustivas del generador de propuestas.
Corre: python3 test_generador.py
"""

import sys, json, traceback, zipfile, io, copy
from pathlib import Path
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from generators import generate
from generators.fda_perfiles import _load_generales, _DESC_CHARS_PER_LINE

OUT_DIR = Path('/tmp/periferia_test')
OUT_DIR.mkdir(parents=True, exist_ok=True)

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

FDA_MARKER   = 'Rectángulo 10'
DESC_SHAPES  = {'CuadroTexto 22', 'CuadroTexto 28', 'CuadroTexto 34', 'CuadroTexto 51'}
ROLE_SHAPES  = {'CuadroTexto 10', 'CuadroTexto 30', 'CuadroTexto 47', 'CuadroTexto 53'}
QA_CARD      = 'CuadroTexto 32'
BULLET_RECTS = {'Rectángulo 10','Rectángulo 13','Rectángulo 19',
                'Rectángulo 22','Rectángulo 25','Rectángulo 28'}

PASS = '\033[92m✓ PASS\033[0m'
FAIL = '\033[91m✗ FAIL\033[0m'
results = []

# ── helpers ──────────────────────────────────────────────────────────────────

def get_fda_slide_shapes(pptx_bytes):
    """Retorna {shape_name: text} de todos los shapes del slide FDA."""
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        slides = sorted(n for n in z.namelist()
                        if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
        for slide_name in slides:
            root = etree.fromstring(z.read(slide_name))
            names = {sp.find(f'.//{{{P}}}cNvPr').attrib.get('name','')
                     for sp in root.iter(f'{{{P}}}sp')
                     if sp.find(f'.//{{{P}}}cNvPr') is not None}
            if FDA_MARKER in names:
                result = {}
                for sp in root.iter(f'{{{P}}}sp'):
                    nvpr = sp.find(f'.//{{{P}}}cNvPr')
                    if nvpr is None: continue
                    name = nvpr.attrib.get('name','')
                    txb  = sp.find(f'{{{P}}}txBody')
                    text = ''.join(t.text or '' for t in txb.findall(f'.//{{{A}}}t')) if txb else ''
                    result[name] = text
                return result
    return {}

def get_all_profile_descs(pptx_bytes):
    """Retorna [(slide, shape_name, text, cy_emu)] de todas las descripciones de perfil."""
    found = []
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        slides = sorted(n for n in z.namelist()
                        if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
        for slide_name in slides:
            root = etree.fromstring(z.read(slide_name))
            for sp in root.iter(f'{{{P}}}sp'):
                nvpr = sp.find(f'.//{{{P}}}cNvPr')
                if nvpr is None: continue
                name = nvpr.attrib.get('name','')
                if name not in DESC_SHAPES: continue
                txb  = sp.find(f'{{{P}}}txBody')
                text = ''.join(t.text or '' for t in txb.findall(f'.//{{{A}}}t')) if txb else ''
                spPr = sp.find(f'{{{P}}}spPr')
                xfrm = spPr.find(f'{{{A}}}xfrm') if spPr is not None else None
                ext  = xfrm.find(f'{{{A}}}ext') if xfrm is not None else None
                cy   = int(ext.attrib.get('cy',0)) if ext is not None else 0
                found.append((slide_name, name, text, cy))
    return found

def get_all_text(pptx_bytes):
    """Texto completo de todo el PPTX."""
    parts = []
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        for n in sorted(z.namelist()):
            if n.startswith('ppt/slides/slide') and n.endswith('.xml'):
                root = etree.fromstring(z.read(n))
                parts.append(''.join(t.text or '' for t in root.iter(f'{{{A}}}t')))
    return ' '.join(parts)

def run(name, config, checks):
    print(f'\n{"─"*64}')
    print(f'  {name}')
    print(f'{"─"*64}')
    try:
        result  = generate(config, str(OUT_DIR))
        pptx_path  = Path(result['propuesta'])
        pptx_bytes = pptx_path.read_bytes()
        print(f'  Archivo: {pptx_path.name} ({pptx_path.stat().st_size//1024} KB)')

        full_text  = get_all_text(pptx_bytes)
        fda_shapes = get_fda_slide_shapes(pptx_bytes)
        descs      = get_all_profile_descs(pptx_bytes)

        failed = []
        for label, cond in checks:
            ok  = cond(full_text, fda_shapes, descs, pptx_bytes)
            sym = PASS if ok else FAIL
            print(f'  {sym}  {label}')
            if not ok:
                failed.append(label)

        results.append((name, len(failed) == 0))

    except Exception as e:
        print(f'  {FAIL}  Excepción: {e}')
        traceback.print_exc()
        results.append((name, False))


# ════════════════════════════ CASOS DE PRUEBA ════════════════════════════════

# ── FDA: QA=SÍ → ítems FDA de QA en cuadro verde ────────────────────────────
run('FDA-1  QA=SÍ, 1 torre → cuadro verde muestra ítems QA', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Cuadro QA tiene texto (no vacío)',
     lambda t,fda,d,p: bool(fda.get(QA_CARD,'').strip())),
    ('Cuadro QA NO tiene mensaje de exclusión',
     lambda t,fda,d,p: 'no hacen parte' not in fda.get(QA_CARD,'').lower()),
    ('Cuadro QA contiene ítems de QA del catálogo',
     lambda t,fda,d,p: 'pruebas' in fda.get(QA_CARD,'').lower() or 'qa' in fda.get(QA_CARD,'').lower()),
    ('Columna principal FDA tiene 6 ítems con texto',
     lambda t,fda,d,p: sum(1 for k in BULLET_RECTS if fda.get(k,'').strip()) == 6),
])

# ── FDA: QA=NO → mensaje exclusión en cuadro verde ──────────────────────────
run('FDA-2  QA=NO, 1 torre → cuadro verde muestra exclusión', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Cuadro QA muestra mensaje de exclusión',
     lambda t,fda,d,p: 'no hacen parte' in fda.get(QA_CARD,'').lower()),
    ('Columna principal FDA tiene ítems (no vacía)',
     lambda t,fda,d,p: sum(1 for k in BULLET_RECTS if fda.get(k,'').strip()) >= 1),
])

# ── FDA: múltiples torres QA=SÍ ─────────────────────────────────────────────
run('FDA-3  múltiples torres + QA=SÍ → cláusula general + QA ítems', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO', 'DATOS', 'PMO'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('FDA principal tiene cláusula general (múltiples torres + genéricos)',
     lambda t,fda,d,p: any('alcance' in fda.get(k,'').lower() for k in BULLET_RECTS)),
    ('Cuadro QA tiene ítems QA (no exclusión)',
     lambda t,fda,d,p: 'no hacen parte' not in fda.get(QA_CARD,'').lower()),
    ('Cuadro QA tiene contenido de pruebas',
     lambda t,fda,d,p: 'pruebas' in fda.get(QA_CARD,'').lower()),
])

# ── FDA: múltiples torres QA=NO ─────────────────────────────────────────────
run('FDA-4  múltiples torres + QA=NO → cláusula general + exclusión', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO', 'DATOS', 'PMO'],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Columna FDA tiene contenido',
     lambda t,fda,d,p: sum(1 for k in BULLET_RECTS if fda.get(k,'').strip()) >= 1),
    ('Cuadro QA muestra exclusión',
     lambda t,fda,d,p: 'no hacen parte' in fda.get(QA_CARD,'').lower()),
])

# ── FDA: pill FDA OFF, 1 torre, QA=SÍ ───────────────────────────────────────
run('FDA-5  pill FDA OFF, 1 torre (FULLSTACK), QA=SÍ → ítems específicos + QA', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': False, 'entregables': True, 'consideraciones': True},
}, [
    ('Columna FDA tiene ítems específicos FULLSTACK',
     lambda t,fda,d,p: sum(1 for k in BULLET_RECTS if fda.get(k,'').strip()) >= 1),
    ('Cuadro QA tiene ítems QA',
     lambda t,fda,d,p: 'pruebas' in fda.get(QA_CARD,'').lower()),
])

# ── FDA: pill FDA OFF, 1 torre, QA=NO ───────────────────────────────────────
run('FDA-6  pill FDA OFF, 1 torre (ARQUITECTURA), QA=NO', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['ARQUITECTURA'],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': False, 'entregables': True, 'consideraciones': True},
}, [
    ('Columna FDA tiene ítems de ARQUITECTURA',
     lambda t,fda,d,p: sum(1 for k in BULLET_RECTS if fda.get(k,'').strip()) >= 1),
    ('Cuadro QA muestra exclusión',
     lambda t,fda,d,p: 'no hacen parte' in fda.get(QA_CARD,'').lower()),
])

# ── PERFILES: descripción corta NO expande innecesariamente ──────────────────
run('PERF-1  Descripción corta → cuadro mantiene tamaño razonable', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
                   'perfiles': [{'perfil': 'Dev Junior', 'seniority': 'Desarrollador junior con 1 año de experiencia.'}]},
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Perfil renderizado',
     lambda t,fda,d,p: 'Dev Junior' in t),
    ('cy del cuadro de descripción es positivo',
     lambda t,fda,d,p: all(cy > 0 for _,_,txt,cy in d if txt.strip())),
    ('Sin "…" en descripciones',
     lambda t,fda,d,p: not any(txt.strip().endswith('…') for _,_,txt,cy in d if txt.strip())),
])

# ── PERFILES: descripción larga → se trunca a 2 oraciones, sin "…" ───────────
DESC_LARGA = ('Desarrollador backend senior con más de 8 años de experiencia '
              'diseñando e implementando APIs RESTful y microservicios en Java y Spring Boot. '
              'Ha liderado equipos de hasta 6 personas en proyectos de transformación digital '
              'en sectores financiero, retail y salud. Certificado en AWS y Google Cloud. '
              'Experiencia sólida en CI/CD, Docker, Kubernetes y metodologías ágiles Scrum/Kanban.')

run('PERF-2  Perfil de Excel → descripción del catálogo o placeholder si no se encuentra', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 100}],
                   'perfiles': [{'perfil': 'Senior Dev', 'seniority': DESC_LARGA}]},
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Perfil renderizado',
     lambda t,fda,d,p: 'Senior Dev' in t),
    ('Tiene descripción (catálogo o placeholder)',
     lambda t,fda,d,p: len(d) > 0),
    ('Sin "…" en descripciones',
     lambda t,fda,d,p: not any(txt.strip().endswith('…') for _,_,txt,cy in d if txt.strip())),
])

# ── PERFILES: genéricos desde catálogo (todas las torres) ────────────────────
run('PERF-3  Genéricos catálogo, 3 torres → múltiples slides, sin "…"', {
    'filial': 'corp', 'excel_data': None,
    'torres_seleccionadas': ['FULLSTACK / DESARROLLO', 'DATOS', 'PMO'],
    'incluir_qa': True,
    'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Hay descripciones de perfil en el PPTX',
     lambda t,fda,d,p: len(d) > 0),
    ('Sin "…" en ninguna descripción',
     lambda t,fda,d,p: not any(txt.strip().endswith('…') for _,_,txt,cy in d if txt.strip())),
    ('cy de descripciones no negativo',
     lambda t,fda,d,p: all(cy >= 0 for _,_,txt,cy in d)),
])

# ── PERFILES: 4 perfiles Excel con descripciones variadas ────────────────────
run('PERF-4  4 perfiles de Excel → descripción catálogo o placeholder por rol', {
    'filial': 'group',
    'excel_data': {'torres': [{'nombre': 'ARQUITECTURA', 'horas': 200}],
                   'perfiles': [
                       {'perfil': 'Arquitecto Senior', 'seniority': ''},
                       {'perfil': 'Tech Lead',          'seniority': ''},
                       {'perfil': 'Cloud Engineer',     'seniority': ''},
                       {'perfil': 'DevOps Engineer',    'seniority': ''},
                   ]},
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('4 roles renderizados',
     lambda t,fda,d,p: all(r in t for r in ['Arquitecto Senior','Tech Lead','Cloud Engineer','DevOps Engineer'])),
    ('Cada perfil tiene descripción (catálogo o placeholder)',
     lambda t,fda,d,p: len(d) >= 4),
    ('Sin "…"',
     lambda t,fda,d,p: not any(txt.strip().endswith('…') for _,_,txt,cy in d if txt.strip())),
])

# ── PERFILES: 1 sólo perfil (centrado) ───────────────────────────────────────
run('PERF-5  1 perfil solo → centrado, sin "…"', {
    'filial': 'cbit',
    'excel_data': {'torres': [{'nombre': 'PMO', 'horas': 80}],
                   'perfiles': [{'perfil': 'PMO Manager', 'seniority': 'Gestión de programas y portafolio de proyectos bajo PMI.'}]},
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Perfil PMO renderizado',
     lambda t,fda,d,p: 'PMO Manager' in t),
    ('Cuadro QA muestra exclusión (QA=No)',
     lambda t,fda,d,p: 'no hacen parte' in fda.get(QA_CARD,'').lower()),
    ('Sin "…"',
     lambda t,fda,d,p: not any(txt.strip().endswith('…') for _,_,txt,cy in d if txt.strip())),
])

# ── PERFILES: vacío / sin descripción ────────────────────────────────────────
run('PERF-6  Perfil sin descripción (campo vacío) → no crashea', {
    'filial': 'corp',
    'excel_data': {'torres': [{'nombre': 'DATOS', 'horas': 60}],
                   'perfiles': [
                       {'perfil': 'Data Scientist', 'seniority': ''},
                       {'perfil': 'Data Engineer', 'seniority': None},
                   ]},
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': False, 'fda': True, 'entregables': True, 'consideraciones': True},
}, [
    ('Roles renderizados sin crash',
     lambda t,fda,d,p: 'Data Scientist' in t or 'Data Engineer' in t),
])

# ── FILIALES: las 3 filiales generan correctamente ───────────────────────────
for filial in ['corp', 'group', 'cbit']:
    run(f'FILIAL-{filial}  Generación básica filial={filial}, QA=SÍ', {
        'filial': filial, 'excel_data': None,
        'torres_seleccionadas': ['FULLSTACK / DESARROLLO'],
        'incluir_qa': True,
        'opciones': {'perfiles': True, 'fda': True, 'entregables': True, 'consideraciones': True},
    }, [
        ('Archivo generado (>10 KB)',
         lambda t,fda,d,p: len(p) > 10_000),
        ('FDA tiene contenido',
         lambda t,fda,d,p: sum(1 for k in BULLET_RECTS if fda.get(k,'').strip()) >= 1),
        ('Cuadro QA tiene ítems',
         lambda t,fda,d,p: bool(fda.get(QA_CARD,'').strip())),
        ('Sin "…" en perfiles',
         lambda t,fda,d,p: not any(txt.strip().endswith('…') for _,_,txt,cy in d if txt.strip())),
    ])

# ── Excel con datos reales (torres del excel) ─────────────────────────────────
run('EXCEL-1  excel_data con 2 torres, pill perfiles OFF, QA=SÍ', {
    'filial': 'corp',
    'excel_data': {
        'torres': [{'nombre': 'FULLSTACK / DESARROLLO', 'horas': 200},
                   {'nombre': 'DATOS', 'horas': 120}],
        'perfiles': [
            {'perfil': 'Full Stack Dev', 'seniority': 'React + Node.js con 4 años de experiencia.'},
            {'perfil': 'Data Engineer',  'seniority': 'Pipelines con Spark, Airflow y Snowflake.'},
            {'perfil': 'Tech Lead',      'seniority': 'Liderazgo técnico y arquitectura de soluciones en proyectos financieros.'},
        ],
    },
    'torres_seleccionadas': [],
    'incluir_qa': True,
    'opciones': {'perfiles': False, 'fda': False, 'entregables': True, 'consideraciones': True},
}, [
    ('Perfiles del Excel aparecen',
     lambda t,fda,d,p: 'Full Stack Dev' in t and 'Data Engineer' in t),
    ('Cuadro QA tiene ítems QA',
     lambda t,fda,d,p: 'pruebas' in fda.get(QA_CARD,'').lower()),
    ('Sin "…"',
     lambda t,fda,d,p: not any(txt.strip().endswith('…') for _,_,txt,cy in d if txt.strip())),
])

run('EXCEL-2  excel_data con 1 torre, pill FDA OFF, QA=NO', {
    'filial': 'group',
    'excel_data': {
        'torres': [{'nombre': 'ARQUITECTURA', 'horas': 160}],
        'perfiles': [],
    },
    'torres_seleccionadas': [],
    'incluir_qa': False,
    'opciones': {'perfiles': True, 'fda': False, 'entregables': True, 'consideraciones': True},
}, [
    ('FDA principal tiene ítems ARQUITECTURA',
     lambda t,fda,d,p: sum(1 for k in BULLET_RECTS if fda.get(k,'').strip()) >= 1),
    ('Cuadro QA muestra exclusión',
     lambda t,fda,d,p: 'no hacen parte' in fda.get(QA_CARD,'').lower()),
])

# ════════════════════════════════ RESUMEN ════════════════════════════════════

print(f'\n{"═"*64}')
print('  RESUMEN')
print(f'{"═"*64}')
passed = sum(1 for _,ok in results if ok)
total  = len(results)
for name, ok in results:
    sym = PASS if ok else FAIL
    print(f'  {sym}  {name}')
print(f'\n  {passed}/{total} casos pasaron')
print(f'{"═"*64}')
sys.exit(0 if passed == total else 1)
