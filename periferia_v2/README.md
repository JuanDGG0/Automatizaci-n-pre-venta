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
git clone https://github.com/JuanDGG0/Automatizaci-n-pre-venta.git
```

**2. Entra a la carpeta del repositorio**
```bash
cd Automatizaci-n-pre-venta
```

**3. Abre en VSCode**
```bash
code .
```

**4. Entra a la carpeta del proyecto**
```bash
cd periferia_v2
```

**5. Instala las dependencias**
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

**1. Asegúrate de estar en la carpeta correcta**
```bash
cd ~/Automatizaci-n-pre-venta/periferia_v2
```

**2. Inicia el servidor**
```bash
python3 server.py
```
Verás: `Servidor listo en http://localhost:8090/generate`

**3. Abre el frontend**

Abre `static/home.html` directamente en el navegador (doble clic desde el explorador de archivos de Windows).

> ⚠️ El servidor debe estar corriendo **antes** de hacer clic en "Generar documento".
> 
> ⚠️ Cada vez que hagas cambios en el código, reinicia el servidor con **Ctrl+C** y vuelve a correr `python3 server.py`.

---

## Estructura del proyecto

```
Automatizaci-n-pre-venta/
└── periferia_v2/
    ├── server.py                        ← Servidor HTTP (no tocar)
    ├── requirements.txt                 ← Dependencias Python
    ├── static/
    │   └── home.html                    ← Interfaz web (abrir en navegador)
    ├── data/
    │   ├── Generales_para_todos.xlsx    ← Catálogo de contenido genérico
    │   └── FOR-CA-CUADRO_BASE_ESTIMACIÓN.xlsx
    ├── templates/
    │   ├── CS-FR-012-...-CORP.pptx      ← Plantilla Periferia IT Corp
    │   ├── CS-FR-005-...-GROUP.pptx     ← Plantilla Periferia IT Group
    │   └── CS-FR-011-...-CBIT.pptx      ← Plantilla CBIT
    └── generators/
        ├── __init__.py                  ← Orquestador (llama a los 3 generators)
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
    """Retorna el orden real de slides según presentation.xml"""
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        rid_map = {r.attrib['Id']: r.attrib['Target'] for r in rels}
        prs = etree.fromstring(z.read('ppt/presentation.xml'))
        ns = {'p': P, 'r': R}
        return ['ppt/' + rid_map[s.attrib[f'{{{R}}}id']]
                for s in prs.find('.//p:sldIdLst', ns)]

def edit(pptx_bytes, config):
    slides_order = _get_slide_order(pptx_bytes)

    # Leer todos los archivos del ZIP
    files_dict = {}
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zin:
        files_dict = {name: zin.read(name) for name in zin.namelist()}

    # Editar TU slide (ej: slide 10 = índice 9)
    slide_key = slides_order[9]  # ← cambia el índice según tu slide
    root = etree.fromstring(files_dict[slide_key])

    # ... tu lógica aquí ...

    files_dict[slide_key] = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    # Reconstruir el ZIP y retornar bytes
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
# test_mi_generator.py  ← crea este archivo en periferia_v2/
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

Corre con:
```bash
python3 test_mi_generator.py
```

---

## Activar tu generator en producción

Cuando termines, dile a Heidy para que descomente tu línea en `generators/__init__.py`:

```python
# pptx_bytes = edit_consideraciones(pptx_bytes, config)        ← Juan
# pptx_bytes = edit_cronograma_entregables(pptx_bytes, config)  ← José
```

---

## Subir cambios al repositorio

Cuando hagas cambios en tu generator y quieras subirlos:

```bash
cd ~/Automatizaci-n-pre-venta
git add .
git commit -m "descripción de lo que hiciste"
git push
```

> Para el `git push` necesitas un token de GitHub (no la contraseña).  
> Generalo en: **https://github.com/settings/tokens** → **Tokens (classic)** → marcar **repo**

---

## Inspeccionar shapes de un slide

Para saber los nombres de los shapes que debes editar en tu slide:

```python
# inspeccionar_slide.py ← crea este archivo en periferia_v2/
import zipfile
from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

# Cambia slide10.xml por el slide que te corresponde
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

```bash
python3 inspeccionar_slide.py
```

---

## Preguntas frecuentes

**¿Por qué no se descarga el PPTX?**  
Verifica que el servidor esté corriendo (`python3 server.py`) y que no haya errores en la terminal.

**¿Puedo probar con el Excel vacío?**  
Sí. Si el Excel no tiene horas en ninguna torre, el sistema te pide que selecciones las torres manualmente y usa el contenido de `data/Generales_para_todos.xlsx`.

**¿Dónde están los datos genéricos?**  
En `data/Generales_para_todos.xlsx`. Tiene 4 hojas: `Fuera del Alcance`, `Perfiles`, `Consideraciones`, `Entregables`.

**¿Por qué el slide X quedó con las X del template?**  
Porque ese generator aún no está implementado. Los slides sin generator quedan con el contenido placeholder del template original.

---

## Contacto

Dudas sobre el proyecto → **Heidy Romero** (Preventa)