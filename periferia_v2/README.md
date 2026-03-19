# 🚀 Generador Automático de Propuestas Comerciales
### Periferia IT Group — Equipo de Preventa

---

## ¿Qué hace esto?

Genera presentaciones PowerPoint de propuestas comerciales automáticamente a partir del Excel de estimación del proyecto. Solo subes el Excel, eliges la filial y el sistema arma el PPTX listo para entregar al cliente.

---

## Requisitos previos

- **Python 3.10 o superior**

Verifica tu versión:
```bash
python3 --version
```

---

## Instalación (primera vez)

**1. Clona el repositorio**
```bash
git clone <url-del-repo>
cd periferia_preventa
```

**2. Instala las dependencias**
```bash
pip install -r requirements.txt
```

Si estás en Linux/WSL y da error de permisos:
```bash
pip install -r requirements.txt --break-system-packages
```

### Librerías usadas

| Librería | Para qué |
|----------|----------|
| `lxml` | Leer y editar el XML interno de los PPTX |
| `openpyxl` | Leer el Excel de catálogo de generales |

---

## Cómo correr el proyecto

**1. Inicia el servidor**
```bash
python3 server.py
```
Verás: `Servidor listo en http://localhost:8090/generate`

**2. Abre el frontend**

Abre `static/home.html` directamente en el navegador (doble clic desde el explorador de archivos).

> ⚠️ El servidor debe estar corriendo antes de generar el documento.

---

## Estructura del proyecto

```
periferia_preventa/
├── server.py                        ← Servidor HTTP (no tocar)
├── requirements.txt
├── static/
│   └── home.html                    ← Interfaz web
├── data/
│   ├── Generales_para_todos.xlsx    ← Catálogo de contenido genérico
│   └── FOR-CA-CUADRO_BASE_ESTIMACIÓN.xlsx
├── templates/
│   ├── CS-FR-012-...-CORP.pptx
│   ├── CS-FR-005-...-GROUP.pptx
│   └── CS-FR-011-...-CBIT.pptx
└── generators/
    ├── __init__.py                  ← Orquestador
    ├── fda_perfiles.py              ← Heidy: slides 8 (Perfiles) y 11 (FDA)
    ├── consideraciones.py           ← Juan: slide 10
    └── cronograma_entregables.py    ← José: slides 7 y 9
```

---

## Slides por responsable

| Slide | Sección | Quién |
|-------|---------|-------|
| 7 | Entregables | José |
| 8 | Perfiles | Heidy |
| 9 | Cronograma | José |
| 10 | Consideraciones | Juan |
| 11 | Fuera del Alcance | Heidy |

---

## Cómo implementar tu generator

Tu función debe llamarse `edit(pptx_bytes, config)`, recibir el PPTX como bytes, editar tus slides y retornar los bytes modificados.

```python
# generators/mi_generator.py
import io, zipfile
from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def _get_slide_order(pptx_bytes):
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        rid_map = {r.attrib['Id']: r.attrib['Target'] for r in rels}
        prs = etree.fromstring(z.read('ppt/presentation.xml'))
        ns = {'p': P, 'r': R}
        return ['ppt/' + rid_map[s.attrib[f'{{{R}}}id']]
                for s in prs.find('.//p:sldIdLst', ns)]

def edit(pptx_bytes, config):
    slides_order = _get_slide_order(pptx_bytes)

    files_dict = {}
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zin:
        files_dict = {name: zin.read(name) for name in zin.namelist()}

    # Editar TU slide (ej: slide 10 = índice 9)
    slide_key = slides_order[9]
    root = etree.fromstring(files_dict[slide_key])

    # ... tu lógica aquí ...

    files_dict[slide_key] = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files_dict.items():
            zout.writestr(name, data)
    return buf.getvalue()
```

### El config que recibes

```python
config = {
    'filial': 'corp',           # 'corp' | 'group' | 'cbit'
    'excel_data': {
        'cliente':  'Empresa XYZ',
        'proyecto': 'Proyecto ABC',
        'torres': [
            {'nombre': 'Full Stack', 'horas': 480, 'semanas': 12},
            {'nombre': 'QA',         'horas': 160, 'semanas': 4},
        ],
        'perfiles': [
            {'torre': 'Full Stack', 'perfil': 'Desarrollador Backend Java'},
        ]
    },
    'torres_seleccionadas': [],  # si el Excel estaba vacío
    'opciones': {
        'perfiles':        'genericos',  # 'excel' | 'genericos'
        'consideraciones': 'genericos',  # 'genericos' | 'manual'
        'entregables':     'genericos',  # 'genericos' | 'manual'
    }
}
```

### Probar solo tu generator

```python
# test_mi_generator.py
from generators.mi_generator import edit

pptx = open('templates/CS-FR-012-PROPUESTA_COMERCIAL_PERIFERIA_IT_CORP.pptx', 'rb').read()

config = {
    'filial': 'corp',
    'excel_data': {
        'cliente': 'Cliente Test',
        'torres': [{'nombre': 'Full Stack', 'horas': 480, 'semanas': 12}]
    },
    'torres_seleccionadas': [],
    'opciones': {}
}

open('test_output.pptx', 'wb').write(edit(pptx, config))
print('Listo! Abre test_output.pptx')
```

```bash
python3 test_mi_generator.py
```

---

## Activar tu generator en producción

Cuando termines, dile a Heidy para que descomente tu línea en `generators/__init__.py`:

```python
# pptx_bytes = edit_consideraciones(pptx_bytes, config)      ← Juan
# pptx_bytes = edit_cronograma_entregables(pptx_bytes, config) ← José
```

---

## Inspeccionar shapes de un slide

Para saber los nombres de los shapes que debes editar:

```python
import zipfile
from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

with zipfile.ZipFile('templates/CS-FR-012-PROPUESTA_COMERCIAL_PERIFERIA_IT_CORP.pptx') as z:
    root = etree.fromstring(z.read('ppt/slides/slide10.xml'))
    for sp in root.iter(f'{{{P}}}sp'):
        nvpr = sp.find(f'.//{{{P}}}cNvPr')
        name = nvpr.attrib.get('name', '') if nvpr is not None else ''
        txb  = sp.find(f'{{{P}}}txBody')
        if txb is not None:
            txt = ''.join(t.text or '' for t in txb.findall(f'.//{{{A}}}t')).strip()
            if txt:
                print(f'[{name}]: {txt[:80]}')
```

> Cambia `slide10.xml` por el slide que te corresponde.

---

## Contacto

Dudas → **Heidy Romero** (Preventa)