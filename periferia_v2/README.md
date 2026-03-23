# Generador Automático de Propuestas Comerciales
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
Verás: `Servidor listo en http://localhost:8090`

**3. Abre el frontend en el navegador**
```
http://localhost:8090
```

> El servidor debe estar corriendo **antes** de hacer clic en "Generar documento".
>
> Cada vez que hagas cambios en el código, reinicia el servidor con **Ctrl+C** y vuelve a correr `python3 server.py`.

---

## Estructura del proyecto

```
periferia_v2/
├── server.py                        ← Servidor HTTP
├── requirements.txt                 ← Dependencias Python
├── static/
│   └── home.html                    ← Interfaz web
├── data/
│   ├── Generales_para_todos.xlsx    ← Catálogo de contenido genérico
│   └── FOR-CA-CUADRO_BASE_ESTIMACIÓN_PROPUESTAS.xlsx
├── templates/
│   ├── CS-FR-012-...-CORP.pptx      ← Plantilla Periferia IT Corp
│   ├── CS-FR-005-...-GROUP.pptx     ← Plantilla Periferia IT Group
│   └── CS-FR-011-...-CBIT.pptx      ← Plantilla CBIT
└── generators/
    ├── __init__.py                  ← Orquestador (llama a los 3 generators en cadena)
    ├── fda_perfiles.py              ← Heidy: slides Perfiles y Fuera del Alcance
    ├── consideraciones.py           ← Juan: slide Consideraciones
    └── cronograma_entregables.py    ← José: slide Entregables
```

---

## Slides por responsable

| Slide | Sección | Quién | Estado |
|-------|---------|-------|--------|
| 7 | Entregables | José | Pendiente |
| 8 | Perfiles | Heidy | Activo |
| 9 | Cronograma | José | Pendiente |
| 10 | Consideraciones | Juan | Pendiente |
| 11 | Fuera del Alcance | Heidy | Activo |

---

## Lógica de cada generator

### Heidy — `fda_perfiles.py`

**Slide Perfiles:**
- Toma perfiles del Excel (hoja `Perfiles` de `Generales_para_todos.xlsx`) filtrados por torres activas, o del Excel del cliente (hoja Anexos) si el usuario subió uno.
- Pagina en slides de máximo 4 tarjetas. Si hay 5+ perfiles, crea slides adicionales.
- Slides con menos de 4 tarjetas se centran automáticamente.
- Pill "Genéricos" ON → usa `Generales_para_todos.xlsx`. Pill OFF + Excel con perfiles → usa los del Excel.

**Slide Fuera del Alcance:**
- Pill ON → muestra la cláusula general (aplica para cualquier combinación de torres).
- Pill OFF + 1 torre → muestra los ítems específicos de esa torre.
- Pill OFF + múltiples torres → combina ítems específicos de cada torre (hasta 6 en total).
- Si QA está entre las torres activas → oculta la card verde de QA (ya está cubierto).
- Si QA no está entre las torres → la card QA muestra "Las pruebas de calidad no hacen parte de esta propuesta."

### Juan — `consideraciones.py`

- Toma consideraciones de `Generales_para_todos.xlsx` hoja `Consideraciones`.
- Filtra por las torres activas del proyecto.
- Muestra máximo 4 consideraciones (una por shape del slide).
- Reemplaza `XXXXXXXXXX` por el nombre del cliente y `Filial` por el nombre de la filial.

### José — `cronograma_entregables.py`

- **Entregables (slide 7):** toma entregables de `Generales_para_todos.xlsx` hoja `Entregables`, agrupa por torres activas, muestra máximo 3 torres (una por columna).
- **Cronograma (slide 9):** pendiente de implementar.

---

## El config que reciben los generators

```python
config = {
    'filial': 'corp',           # 'corp' | 'group' | 'cbit'
    'excel_data': {
        'cliente':  'Empresa XYZ',
        'proyecto': 'Proyecto ABC',
        'torres': [
            {'nombre': 'FULLSTACK / DESARROLLO', 'horas': 480, 'semanas': 12},
            {'nombre': 'QA',                     'horas': 160, 'semanas': 4},
        ],
        'perfiles': [
            {'torre': 'FULLSTACK / DESARROLLO', 'perfil': 'Desarrollador Backend Java',
             'seniority': 'Senior, 5+ años en Java Spring Boot', 'horas_mes': 160, 'cantidad': 1},
        ]
    },
    'torres_seleccionadas': [],   # usado si excel_data.torres está vacío
    'opciones': {
        'perfiles':        True,  # True = usar genéricos, False = usar datos del Excel
        'fda':             False, # True = cláusula general, False = específico por torre
        'consideraciones': False, # True = usar genéricos
        'entregables':     False, # True = usar genéricos
    }
}
```

> **Nombres de torres:** deben usar los nombres canónicos definidos en `Generales_para_todos.xlsx`.
> El frontend ya mapea automáticamente los nombres del Excel del cliente a estos nombres canónicos.

---

## Datos genéricos — `Generales_para_todos.xlsx`

| Hoja | Contenido | Usado por |
|------|-----------|-----------|
| `Fuera del Alcance` | Ítems FDA por torre + cláusula general | Heidy |
| `Perfiles` | Roles y descripciones por torre | Heidy |
| `Consideraciones` | Consideraciones por torre | Juan |
| `Entregables` | Lista de entregables por torre | José |

Para agregar o modificar contenido genérico, edita directamente este Excel. No se necesita cambiar código.

---

## Subir cambios al repositorio

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

Para saber los nombres de los shapes de un slide específico:

```python
# inspeccionar_slide.py  ← crea este archivo en periferia_v2/ (no commitear)
import zipfile
from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

with zipfile.ZipFile('templates/CS-FR-012-PROPUESTA_COMERCIAL_PERIFERIA_IT_CORP.pptx') as z:
    root = etree.fromstring(z.read('ppt/slides/slide10.xml'))  # cambia el número
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

**¿Por qué el slide de Cronograma quedó sin datos?**
El slide de Cronograma (slide 9) está pendiente de implementar. Queda con el contenido placeholder del template.

---

