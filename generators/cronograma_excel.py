"""
cronograma_excel.py
Genera el cronograma como XLSX usando shapes roundRect (drawingML),
replicando el estilo de la plantilla FABRICA_PR-FR-004.
"""

import io
import math
import zipfile
from lxml import etree
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# ─── Constantes ──────────────────────────────────────────────────────────────
HORAS_POR_SEMANA   = 40
SEMANAS_POR_MES    = 4
SEMANAS_POR_SPRINT = 2

# Colores de barras por torre (ciclo)
BARRA_COLORES = [
    "0DC56D",  # verde
    "2FF195",  # verde claro
    "24BABA",  # teal
    "2F2BCB",  # azul/morado
    "1D1B80",  # azul oscuro
    "FF6D00",  # naranja
    "D50000",  # rojo
]

# Colores del header (derivados del tema de la plantilla)
COLOR_MES     = "757070"   # gris oscuro
COLOR_SEMANA  = "D0CECE"   # gris claro
COLOR_SPRINT  = "798EA9"   # azul grisáceo
COLOR_KICKOFF = "757070"

PADDING        = 38_100
COL_A_EMU      = 1_781_175
COL_SEMANA_EMU =   414_337
ROW_HDR1_EMU   =   304_800
ROW_HDR2_EMU   =   254_000
ROW_ACT_EMU    =   279_400


# ─── API pública ─────────────────────────────────────────────────────────────
def generate_cronograma(config: dict) -> bytes:
    actividades = config.get("actividades", [])
    if not actividades:
        raise ValueError("No hay actividades para generar cronograma")

    for act in actividades:
        act["semanas"] = max(1, math.ceil(act["horas"] / HORAS_POR_SEMANA))

    total_semanas = max(a["semanas"] for a in actividades)
    total_cols = total_semanas + 1  # col 0-indexed 1 = Kick Off, 2.. = semanas

    wb = Workbook()
    ws = wb.active
    ws.title = "Cronograma"
    ws.sheet_view.showGridLines = False

    _configurar_dimensiones(ws, total_cols)

    buffer = io.BytesIO()
    wb.save(buffer)

    return _inyectar_drawing(buffer.getvalue(), actividades, total_cols)


# ─── Dimensiones ─────────────────────────────────────────────────────────────
def _configurar_dimensiones(ws, total_cols):
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 20

    ws.column_dimensions['A'].width = 26
    for i in range(1, total_cols + 2):
        ws.column_dimensions[get_column_letter(i + 1)].width = 5.5

    for row_idx in range(4, 4 + 30):
        ws.row_dimensions[row_idx].height = 22


# ─── Inyección del drawing ────────────────────────────────────────────────────
def _inyectar_drawing(xlsx_bytes: bytes, actividades: list, total_cols: int) -> bytes:
    drawing_xml = _build_drawing_xml(actividades, total_cols)

    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), 'r') as zin:
        files = {n: zin.read(n) for n in zin.namelist()}

    files['xl/drawings/drawing1.xml'] = drawing_xml.encode('utf-8')

    sheet_rel_path = 'xl/worksheets/_rels/sheet1.xml.rels'
    if sheet_rel_path not in files:
        files[sheet_rel_path] = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
            b'</Relationships>'
        )
    else:
        rel_xml = files[sheet_rel_path].decode('utf-8')
        if 'drawing1.xml' not in rel_xml:
            rel_xml = rel_xml.replace(
                '</Relationships>',
                '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'
            )
            files[sheet_rel_path] = rel_xml.encode('utf-8')

    sheet_xml = files['xl/worksheets/sheet1.xml'].decode('utf-8')
    if '<drawing ' not in sheet_xml:
        # openpyxl omits xmlns:r when no relationships exist — add it so r:id is valid
        if 'xmlns:r=' not in sheet_xml:
            sheet_xml = sheet_xml.replace(
                '<worksheet ',
                '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
                1
            )
        sheet_xml = sheet_xml.replace('</worksheet>', '<drawing r:id="rId10"/></worksheet>')
        files['xl/worksheets/sheet1.xml'] = sheet_xml.encode('utf-8')

    ct = files['[Content_Types].xml'].decode('utf-8')
    if 'drawing1.xml' not in ct:
        ct = ct.replace(
            '</Types>',
            '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
        )
        files['[Content_Types].xml'] = ct.encode('utf-8')

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    return out.getvalue()


# ─── Construcción del XML de shapes ──────────────────────────────────────────
def _build_drawing_xml(actividades: list, total_cols: int) -> str:
    XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    root = etree.Element(f'{{{XDR}}}wsDr', nsmap={
        'xdr': XDR, 'a': A,
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    })

    sid = 1
    total_semanas = total_cols - 1

    # ── HEADER ───────────────────────────────────────────────────────────
    # Kick Off (col 1, 1 sola col) — single-col: pass col_width_emu
    sid = _shape(root, sid, "Kick Off", 1, 0, 1, 1,
                 COLOR_KICKOFF, "default", "FFFFFF", 900, True,
                 col_width_emu=COL_SEMANA_EMU, row_height_emu=ROW_HDR1_EMU)

    # Meses — multi-col, offsets work normally
    col_cur = 2
    for m in range(math.ceil(total_semanas / SEMANAS_POR_MES)):
        end = min(col_cur + SEMANAS_POR_MES, 2 + total_semanas)
        sid = _shape(root, sid, f"Mes {m+1}", col_cur, 0, end - 1, 1,
                     COLOR_MES, "default", "FFFFFF", 900, True,
                     row_height_emu=ROW_HDR1_EMU)
        col_cur = end

    # Semanas — S0 Kick Off, S1..Sn cada semana, todas de 1 col
    sid = _shape(root, sid, "S0", 1, 1, 1, 2,
                 COLOR_SEMANA, "default", "44546A", 800,
                 col_width_emu=COL_SEMANA_EMU, row_height_emu=ROW_HDR2_EMU)
    for s in range(total_semanas):
        col = 2 + s
        sid = _shape(root, sid, f"S{s+1}", col, 1, col, 2,
                     COLOR_SEMANA, "default", "44546A", 800,
                     col_width_emu=COL_SEMANA_EMU, row_height_emu=ROW_HDR2_EMU)

    # Sprints — Sprint 0 de 1 col, el resto de 2 cols
    sid = _shape(root, sid, "Sprint 0", 1, 2, 1, 3,
                 COLOR_SPRINT, "default", "FFFFFF", 800,
                 col_width_emu=COL_SEMANA_EMU, row_height_emu=ROW_HDR2_EMU)
    for s in range(0, total_semanas, SEMANAS_POR_SPRINT):
        sprint_n  = s // SEMANAS_POR_SPRINT + 1
        start_col = 2 + s
        end_col   = min(start_col + SEMANAS_POR_SPRINT - 1, 1 + total_semanas)
        sid = _shape(root, sid, f"Sprint {sprint_n}", start_col, 2, end_col, 3,
                     COLOR_SPRINT, "default", "FFFFFF", 800,
                     row_height_emu=ROW_HDR2_EMU)

    # ── ACTIVIDADES ──────────────────────────────────────────────────────
    for i, act in enumerate(actividades):
        row   = 3 + i
        color = BARRA_COLORES[i % len(BARRA_COLORES)]

        # Etiqueta de torre — col A (0), 1 sola col, ancho especial
        sid = _shape(root, sid, act["torre"], 0, row, 0, row + 1,
                     color, "17948", "FFFFFF", 900, True,
                     col_width_emu=COL_A_EMU, row_height_emu=ROW_ACT_EMU)

        # Barra horizontal — multi-col, offsets normales
        bar_end = 1 + act["semanas"]
        sid = _shape(root, sid, "", 2, row, bar_end, row + 1,
                     color, "50000", "FFFFFF", 900,
                     row_height_emu=ROW_ACT_EMU)

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True).decode('utf-8')


# ─── Helper: una shape twoCellAnchor ─────────────────────────────────────────
def _shape(parent, sid: int, text: str,
           col_from: int, row_from: int,
           col_to: int,   row_to: int,
           color: str, adj: str = "default",
           text_color: str = "FFFFFF",
           font_size: int = 900, bold: bool = False,
           col_width_emu: int = COL_SEMANA_EMU,
           row_height_emu: int = ROW_ACT_EMU) -> int:

    XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    anchor = etree.SubElement(parent, f'{{{XDR}}}twoCellAnchor', editAs='oneCell')

    frm = etree.SubElement(anchor, f'{{{XDR}}}from')
    etree.SubElement(frm, f'{{{XDR}}}col').text    = str(col_from)
    etree.SubElement(frm, f'{{{XDR}}}colOff').text = str(PADDING)
    etree.SubElement(frm, f'{{{XDR}}}row').text    = str(row_from)
    etree.SubElement(frm, f'{{{XDR}}}rowOff').text = str(PADDING)

    # When col_from == col_to the shape is inside one column.
    # colOff must reach (col_width_emu - PADDING), NOT -PADDING (which would be negative).
    # Same logic for rows spanning a single row.
    to_col_off = (col_width_emu - PADDING) if col_from == col_to else -PADDING
    to_row_off = (row_height_emu - PADDING) if row_from == row_to else -PADDING

    to = etree.SubElement(anchor, f'{{{XDR}}}to')
    etree.SubElement(to, f'{{{XDR}}}col').text    = str(col_to)
    etree.SubElement(to, f'{{{XDR}}}colOff').text = str(to_col_off)
    etree.SubElement(to, f'{{{XDR}}}row').text    = str(row_to)
    etree.SubElement(to, f'{{{XDR}}}rowOff').text = str(to_row_off)

    sp      = etree.SubElement(anchor, f'{{{XDR}}}sp', macro='', textlink='')
    nvSpPr  = etree.SubElement(sp, f'{{{XDR}}}nvSpPr')
    etree.SubElement(nvSpPr, f'{{{XDR}}}cNvPr', id=str(sid), name=f"shape{sid}")
    etree.SubElement(nvSpPr, f'{{{XDR}}}cNvSpPr')

    spPr     = etree.SubElement(sp, f'{{{XDR}}}spPr')
    prstGeom = etree.SubElement(spPr, f'{{{A}}}prstGeom', prst='roundRect')
    avLst    = etree.SubElement(prstGeom, f'{{{A}}}avLst')
    if adj != "default":
        etree.SubElement(avLst, f'{{{A}}}gd', name='adj', fmla=f'val {adj}')

    sf = etree.SubElement(spPr, f'{{{A}}}solidFill')
    etree.SubElement(sf, f'{{{A}}}srgbClr', val=color)

    ln = etree.SubElement(spPr, f'{{{A}}}ln')
    etree.SubElement(ln, f'{{{A}}}noFill')

    txBody = etree.SubElement(sp, f'{{{XDR}}}txBody')
    etree.SubElement(txBody, f'{{{A}}}bodyPr', wrap='square', rtlCol='0', anchor='ctr')
    etree.SubElement(txBody, f'{{{A}}}lstStyle')

    p   = etree.SubElement(txBody, f'{{{A}}}p')
    etree.SubElement(p, f'{{{A}}}pPr', algn='ctr')

    if text:
        r    = etree.SubElement(p, f'{{{A}}}r')
        attrs = {'lang': 'es-CO', 'sz': str(font_size), 'dirty': '0'}
        if bold:
            attrs['b'] = '1'
        rPr  = etree.SubElement(r, f'{{{A}}}rPr', **attrs)
        sf2  = etree.SubElement(rPr, f'{{{A}}}solidFill')
        etree.SubElement(sf2, f'{{{A}}}srgbClr', val=text_color)
        etree.SubElement(rPr, f'{{{A}}}latin', typeface='Calibri')
        etree.SubElement(r, f'{{{A}}}t').text = text

    etree.SubElement(anchor, f'{{{XDR}}}clientData')
    return sid + 1