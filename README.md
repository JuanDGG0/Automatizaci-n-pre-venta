# Generador de Propuestas Comerciales — Periferia IT

Genera presentaciones PowerPoint de propuestas comerciales a partir del Excel de estimación del proyecto. Subes el Excel, eliges la filial, y el sistema arma el PPTX listo para entregar al cliente.

## Requisitos

- Python 3.10 o superior

## Instalación

```bash
git clone https://github.com/JuanDGG0/Automatizaci-n-pre-venta.git
cd Automatizaci-n-pre-venta
pip install -r requirements.txt --break-system-packages
```

## Uso

```bash
python3 server.py
```

Luego abre `http://localhost:8090` en el navegador. Para detener el servidor presiona `Ctrl+C`.

> Cada vez que modifiques código, reinicia el servidor con `Ctrl+C` y vuelve a correr `python3 server.py`.

## Estructura

```
Automatizaci-n-pre-venta/
├── server.py                     # Servidor HTTP (multihilo, puerto 8090)
├── requirements.txt
├── static/
│   └── home.html                 # Interfaz web
├── data/
│   ├── Generales_para_todos.xlsx # Catálogo de contenido genérico
│   └── FOR-CA-CUADRO_BASE_ESTIMACIÓN_PROPUESTAS.xlsx
├── templates/
│   ├── CS-FR-012-...-CORP.pptx
│   ├── CS-FR-005-...-GROUP.pptx
│   └── CS-FR-011-...-CBIT.pptx
├── generators/
│   ├── __init__.py               # Orquestador
│   ├── fda_perfiles.py           # Slides Perfiles y Fuera del Alcance
│   ├── consideraciones.py        # Slide Consideraciones
│   └── cronograma_entregables.py # Slide Entregables
└── tests/
    └── test_e2e.py
```

## Cómo funciona el flujo

El frontend parsea el Excel del cliente en tres hojas:

- **RESUMEN** → nombre del cliente y lista de torres con horas
- **Estimación** → consideraciones (col J), fuera del alcance (col K), entregables por torre (col M)
- **Anexos** → perfiles del equipo

Con esos datos construye un payload que envía al servidor (`POST /generate`), el cual llama a los tres generators en cadena sobre el template PPTX de la filial elegida.

Cada sección tiene una pill en el frontend que controla si se agregan ítems genéricos del catálogo encima de los datos del Excel:

- **Pill ON** → datos del Excel + genéricos del catálogo para complementar
- **Pill OFF** → solo datos del Excel (o catálogo completo si el Excel no trae datos de esa sección)

## Generators

### `fda_perfiles.py` — Perfiles y Fuera del Alcance

**Perfiles:** toma los datos del Excel (hoja Anexos), busca la descripción de cada rol en el catálogo, y pagina en slides de máximo 4 tarjetas. Si el Excel no trae perfiles, el usuario puede elegirlos manualmente desde el buscador del frontend (`GET /api/perfiles-catalog`).

**Fuera del Alcance:** si la pill está ON muestra la cláusula general; si está OFF usa los ítems de col K del Excel, o el catálogo por torre si el Excel no tiene datos. Pagina automáticamente si hay más de 6 ítems.

### `consideraciones.py` — Consideraciones

Toma los ítems de col J del Excel de estimación (siempre los incluye). Si la pill está ON agrega además los genéricos del catálogo filtrados por torres activas. Pagina en grupos que caben en cada slide según la altura real del texto — los textos largos agrandan el grupo automáticamente. Reemplaza el nombre del cliente y la filial en el texto.

### `cronograma_entregables.py` — Entregables

Toma los entregables de col M del Excel de estimación agrupados por torre. Si la pill está ON complementa con el catálogo para las torres que no tengan entregables en el Excel. Si el Excel no trae ningún entregable usa el catálogo completo. Soporta hasta 4 torres por slide (1–2 centradas, 3 en posición template, 4 a escala 75%); si hay más de 4 torres duplica el slide.

## Catálogo genérico — `Generales_para_todos.xlsx`

Tiene cuatro hojas: `Fuera del Alcance`, `Perfiles`, `Consideraciones`, `Entregables`. Para agregar o modificar contenido genérico edita este archivo directamente, sin tocar código.

## Subir cambios

```bash
git add .
git commit -m "descripción del cambio"
git push
```

> Para el push necesitas un token de GitHub (no la contraseña). Generalo en **GitHub → Settings → Developer settings → Tokens (classic)** con scope `repo`.

## Inspeccionar un slide

```python
import zipfile
from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

with zipfile.ZipFile('templates/CS-FR-012-PROPUESTA_COMERCIAL_PERIFERIA_IT_CORP.pptx') as z:
    root = etree.fromstring(z.read('ppt/slides/slide7.xml'))
    for sp in root.iter(f'{{{P}}}sp'):
        nvpr = sp.find(f'.//{{{P}}}cNvPr')
        name = nvpr.attrib.get('name', '') if nvpr is not None else ''
        txb  = sp.find(f'{{{P}}}txBody')
        if txb is not None:
            txt = ''.join(t.text or '' for t in txb.findall(f'.//{{{A}}}t')).strip()
            if txt:
                print(f'[{name}]: {txt[:80]}')
```

## Tests

```bash
python3 tests/test_e2e.py
```

Corre 46 casos que cubren entregables, consideraciones, perfiles, detección de slides y casos borde para las tres filiales.
