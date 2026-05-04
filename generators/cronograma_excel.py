"""
cronograma_excel.py
Genera cronograma XLSX con drawingML roundRect shapes.

Layout:
  Fila 0         : Roles (una sola fila horizontal, cards a lo largo del cronograma)
  Fila 1         : Kick Off | Meses
  Fila 2         : S0       | Semanas S1..Sn  (omitida si > 6 meses)
  Fila 3 (o 2)   : Sprint 0 | Sprints 1..N
  Fila 4+ (o 3+) : Actividades (primera = "Preparar Ambientes" siempre, Sprint 0 completo)
"""

import io
import math
import zipfile
from lxml import etree
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# ─── Constantes ──────────────────────────────────────────────────────────────
HORAS_SEMANALES    = 43
SEMANAS_POR_MES    = 4
SEMANAS_POR_SPRINT = 2
UMBRAL_MESES       = 24   # si total_semanas > esto → ocultar fila de semanas

BARRA_COLORES = [
    "0DC56D", "2FF195", "24BABA",
    "2F2BCB", "1D1B80", "FF6D00", "D50000",
]
COLOR_PREP_AMBIENTES = "4A4A8A"   # azul oscuro fijo para "Preparar Ambientes"

COLOR_MES     = "757070"
COLOR_SEMANA  = "D0CECE"
COLOR_SPRINT  = "798EA9"
COLOR_KICKOFF = "757070"

ROLE_COLORS = [
    "2F2BCB", "1D1B80", "24BABA",
    "0DC56D", "FF6D00", "D50000",
    "798EA9", "757070",
]

PADDING        = 38_100
COL_A_EMU      = 1_781_175
COL_SEMANA_EMU =   414_337
ROW_HDR_EMU    =   304_800
ROW_SUB_EMU    =   254_000
ROW_ACT_EMU    =   279_400
ROW_ROLES_EMU  =   600_000   # altura fija de la fila de roles

# ─── API pública ─────────────────────────────────────────────────────────────
def generate_cronograma(config: dict) -> bytes:
    actividades = config.get("actividades", [])
    roles_raw   = config.get("roles", [])

    if not actividades:
        raise ValueError("No hay actividades")

    # Normalizar roles: solo los que tienen perfil asignado (col[1] no vacío)
    roles = []
    for r in roles_raw:
        if isinstance(r, dict):
            nombre = str(r.get("perfil", r.get("rol", ""))).strip()
            if nombre:
                roles.append({
                    "perfil":    nombre,
                    "seniority": str(r.get("seniority", "")).strip(),
                    "personas":  max(1, int(r.get("personas", 1))),
                    "torre":     str(r.get("torre", "")).strip(),
                })
        else:
            nombre = str(r).strip()
            if nombre:
                roles.append({"perfil": nombre, "seniority": "", "personas": 1, "torre": ""})

    # Calcular semanas por actividad
    for act in actividades:
        p = max(1, int(act.get("personas", 1)))
        act["semanas"] = max(1, math.ceil(act["horas"] / p / HORAS_SEMANALES))

    total_semanas = max(a["semanas"] for a in actividades)
    sin_semanas   = total_semanas > UMBRAL_MESES

    # Filas de header: roles(1) + mes(1) + semana?(1) + sprint(1)
    n_hdr_sin_roles = 2 if sin_semanas else 3   # mes + sprint  OR  mes + semana + sprint
    n_hdr = 1 + n_hdr_sin_roles                 # +1 for roles row

    wb = Workbook()
    ws = wb.active
    ws.title = "Cronograma"
    ws.sheet_view.showGridLines = False
    _configurar_dimensiones(ws, total_semanas, n_hdr, sin_semanas)

    buf = io.BytesIO()
    wb.save(buf)

    return _inyectar_drawing(buf.getvalue(), actividades, roles,
                             total_semanas, sin_semanas, n_hdr)

# ─── Dimensiones ─────────────────────────────────────────────────────────────
def _configurar_dimensiones(ws, total_semanas, n_hdr, sin_semanas):
    ws.row_dimensions[1].height = 45    # fila 0 → roles (en puntos, ~600k EMU)
    ws.row_dimensions[2].height = 24    # fila 1 → mes / kick off
    if not sin_semanas:
        ws.row_dimensions[3].height = 20   # fila 2 → semanas / S0
        ws.row_dimensions[4].height = 20   # fila 3 → sprints / Sprint 0
        act_start = 5
    else:
        ws.row_dimensions[3].height = 20   # fila 2 → sprints / Sprint 0
        act_start = 4

    for r in range(act_start, act_start + 40):
        ws.row_dimensions[r].height = 22

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 5.5
    for i in range(total_semanas):
        ws.column_dimensions[get_column_letter(i + 3)].width = 5.5

# ─── Inyección ZIP ───────────────────────────────────────────────────────────
def _inyectar_drawing(xlsx_bytes, actividades, roles,
                      total_semanas, sin_semanas, n_hdr):

    drawing_xml = _build_drawing_xml(actividades, roles,
                                     total_semanas, sin_semanas, n_hdr)

    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as zin:
        files = {n: zin.read(n) for n in zin.namelist()}

    files["xl/drawings/drawing1.xml"] = drawing_xml.encode("utf-8")

    rel_path = "xl/worksheets/_rels/sheet1.xml.rels"
    if rel_path not in files:
        files[rel_path] = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"'
            b' Target="../drawings/drawing1.xml"/></Relationships>'
        )
    else:
        rel = files[rel_path].decode("utf-8")
        if "drawing1.xml" not in rel:
            rel = rel.replace("</Relationships>",
                              '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
            files[rel_path] = rel.encode("utf-8")

    sheet = files["xl/worksheets/sheet1.xml"].decode("utf-8")
    if "<drawing " not in sheet:
        if "xmlns:r=" not in sheet:
            sheet = sheet.replace("<worksheet ",
                '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
        sheet = sheet.replace("</worksheet>", '<drawing r:id="rId10"/></worksheet>')
        files["xl/worksheets/sheet1.xml"] = sheet.encode("utf-8")

    ct = files["[Content_Types].xml"].decode("utf-8")
    if "drawing1.xml" not in ct:
        ct = ct.replace("</Types>",
            '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
        files["[Content_Types].xml"] = ct.encode("utf-8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    return out.getvalue()

# ─── Drawing XML principal ───────────────────────────────────────────────────
def _build_drawing_xml(actividades, roles, total_semanas, sin_semanas, n_hdr):
    XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
    root = etree.Element(f"{{{XDR}}}wsDr",
                         nsmap={"xdr": XDR, "a": A,
                                "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"})
    sid = 1

    # ── Índices de filas (0-based para drawingML) ─────────────────────────
    ROW_ROLES  = 0
    ROW_MES    = 1
    if sin_semanas:
        ROW_SEMANA = None
        ROW_SPRINT = 2
    else:
        ROW_SEMANA = 2
        ROW_SPRINT = 3

    COL_KO = 1    # col B = Kick Off / S0 / Sprint 0
    COL_S1 = 2    # primera semana real

    # ── 1. FILA DE ROLES (fila 0, una sola fila horizontal) ───────────────
    # Cards distribuidas uniformemente a lo largo del ancho total del cronograma:
    # col A (0) + col B (KO, 1) + cols semanas (2..N+1)
    # Ancho total disponible = COL_A_EMU + (total_semanas + 1) * COL_SEMANA_EMU
    TOTAL_W_EMU = COL_A_EMU + (total_semanas + 1) * COL_SEMANA_EMU
    n_roles = len(roles)

    if n_roles > 0:
        # Distribuir uniformemente: cada card ocupa TOTAL_W / n_roles de ancho
        # pero con un mínimo razonable para que el texto quepa
        CARD_GAP    = 25_000
        BADGE_SIZE  = 160_000
        card_w      = max(400_000, TOTAL_W_EMU // n_roles)  # mínimo ~400k EMU

        for i, rol in enumerate(roles):
            color = ROLE_COLORS[i % len(ROLE_COLORS)]

            nombre    = rol["perfil"]
            seniority = rol["seniority"]
            personas  = rol["personas"]
            label     = f"{nombre}\n{seniority}" if seniority else nombre

            # Posición X absoluta del card en el sheet
            # Las shapes de drawingML usan col + colOffset
            # Calculamos en qué columna cae x_start y cuál es el offset
            x_start = i * card_w
            x_end   = x_start + card_w - CARD_GAP * 2

            col_from, x_off_from = _x_to_col_offset(x_start + CARD_GAP, total_semanas)
            col_to,   x_off_to   = _x_to_col_offset(x_end, total_semanas)

            # Card principal
            sid = _shape_span(root, sid, label,
                              col_from, ROW_ROLES, x_off_from,
                              col_to,   ROW_ROLES, x_off_to,
                              ROW_ROLES_EMU - CARD_GAP,
                              color, "17948", "FFFFFF", 700, True)

            # Badge circular con nº personas (top-right de la card)
            badge_x = x_end - BADGE_SIZE // 2
            badge_col, badge_x_off = _x_to_col_offset(badge_x, total_semanas)
            sid = _shape_span(root, sid, str(personas),
                              badge_col, ROW_ROLES, badge_x_off,
                              badge_col, ROW_ROLES, badge_x_off + BADGE_SIZE,
                              BADGE_SIZE,
                              "FFFFFF", "default", color, 700, True,
                              geom="ellipse")

    # ── 2. KICK OFF (col B, filas MES → SPRINT, span vertical) ───────────
    sid = _shape(root, sid, "Kick Off",
                 COL_KO, ROW_MES, COL_KO, ROW_SPRINT,
                 COLOR_KICKOFF, "default", "FFFFFF", 900, True)

    # ── S0 (debajo de Kick Off, en fila SEMANA si existe) ─────────────────
    if ROW_SEMANA is not None:
        sid = _shape(root, sid, "S0",
                     COL_KO, ROW_SEMANA, COL_KO, ROW_SEMANA,
                     COLOR_SEMANA, "default", "44546A", 800)

    # ── Sprint 0 (debajo de Kick Off, en fila SPRINT) ─────────────────────
    sid = _shape(root, sid, "Sprint 0",
                 COL_KO, ROW_SPRINT, COL_KO, ROW_SPRINT,
                 COLOR_SPRINT, "default", "FFFFFF", 800)

    # ── 3. MESES ──────────────────────────────────────────────────────────
    col_cur = COL_S1
    for m in range(math.ceil(total_semanas / SEMANAS_POR_MES)):
        col_end = min(col_cur + SEMANAS_POR_MES - 1, COL_S1 + total_semanas - 1)
        sid = _shape(root, sid, f"Mes {m+1}",
                     col_cur, ROW_MES, col_end, ROW_MES,
                     COLOR_MES, "default", "FFFFFF", 900, True,
                     row_height_emu=ROW_HDR_EMU)
        col_cur += SEMANAS_POR_MES

    # ── 4. SEMANAS S1..Sn ────────────────────────────────────────────────
    if ROW_SEMANA is not None:
        for s in range(total_semanas):
            col = COL_S1 + s
            sid = _shape(root, sid, f"S{s+1}",
                         col, ROW_SEMANA, col, ROW_SEMANA,
                         COLOR_SEMANA, "default", "44546A", 800,
                         row_height_emu=ROW_SUB_EMU)

    # ── 5. SPRINTS 1..N ──────────────────────────────────────────────────
    sprint_n = 1
    for s in range(0, total_semanas, SEMANAS_POR_SPRINT):
        col_start = COL_S1 + s
        col_end   = min(col_start + SEMANAS_POR_SPRINT - 1, COL_S1 + total_semanas - 1)
        sid = _shape(root, sid, f"Sprint {sprint_n}",
                     col_start, ROW_SPRINT, col_end, ROW_SPRINT,
                     COLOR_SPRINT, "default", "FFFFFF", 800,
                     row_height_emu=ROW_SUB_EMU)
        sprint_n += 1

    # ── 6. "PREPARAR AMBIENTES" — siempre primera actividad en Sprint 0 ───
    # Ocupa toda la duración del Sprint 0 = SEMANAS_POR_SPRINT semanas desde COL_S1
    # (pero en realidad la barra arranca en col B = COL_KO, que es donde está Sprint 0)
    prep_row = n_hdr   # primera fila de actividades (0-based)
    prep_bar_end = COL_KO  # solo ocupa la columna de Kick Off (Sprint 0 = S0 = 1 col)

    # Etiqueta "Preparar Ambientes"
    sid = _shape(root, sid, "Preparar Ambientes",
                 0, prep_row, 0, prep_row,
                 COLOR_PREP_AMBIENTES, "17948", "FFFFFF", 900, True,
                 col_width_emu=COL_A_EMU, row_height_emu=ROW_ACT_EMU)

    # Barra en la columna de Sprint 0 (col B = col 1)
    sid = _shape(root, sid, "",
                 COL_KO, prep_row, COL_KO, prep_row,
                 COLOR_PREP_AMBIENTES, "50000", "FFFFFF", 900,
                 row_height_emu=ROW_ACT_EMU)

    # ── 7. ACTIVIDADES del Excel ──────────────────────────────────────────
    for i, act in enumerate(actividades):
        row   = n_hdr + 1 + i   # +1 porque Preparar Ambientes ocupa la primera
        color = BARRA_COLORES[i % len(BARRA_COLORES)]

        sid = _shape(root, sid, act["torre"],
                     0, row, 0, row,
                     color, "17948", "FFFFFF", 900, True,
                     col_width_emu=COL_A_EMU, row_height_emu=ROW_ACT_EMU)

        # La barra empieza en COL_S1 (después de Kick Off)
        bar_end = COL_S1 + act["semanas"] - 1
        sid = _shape(root, sid, "",
                     COL_S1, row, bar_end, row,
                     color, "50000", "FFFFFF", 900,
                     row_height_emu=ROW_ACT_EMU)

    return etree.tostring(root, xml_declaration=True,
                          encoding="UTF-8", standalone=True).decode("utf-8")

# ─── Helper: convertir X absoluto en EMU a (col, offset) ────────────────────
def _x_to_col_offset(x_emu: int, total_semanas: int):
    """
    Dado un X absoluto en EMU desde el borde izquierdo del sheet,
    devuelve (col_index_0based, offset_dentro_de_esa_col).
    Col 0 = A (COL_A_EMU), col 1 = B (COL_SEMANA_EMU), col 2+ = semanas (COL_SEMANA_EMU).
    """
    if x_emu < COL_A_EMU:
        return 0, x_emu
    x_emu -= COL_A_EMU
    col = 1 + x_emu // COL_SEMANA_EMU
    off = x_emu % COL_SEMANA_EMU
    return int(col), int(off)

# ─── Shape con span horizontal usando col+offset para from y to ─────────────
def _shape_span(parent, sid, text,
                col_from, row_from, x_off_from,
                col_to,   row_from2, x_off_to,
                row_h_emu,
                color, adj="default",
                text_color="FFFFFF", font_size=800, bold=False,
                geom="roundRect"):
    """Shape donde from y to se especifican como (col, row, x_offset_emu)."""
    XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    A   = "http://schemas.openxmlformats.org/drawingml/2006/main"

    anchor = etree.SubElement(parent, f"{{{XDR}}}twoCellAnchor", editAs="oneCell")

    frm = etree.SubElement(anchor, f"{{{XDR}}}from")
    etree.SubElement(frm, f"{{{XDR}}}col").text    = str(col_from)
    etree.SubElement(frm, f"{{{XDR}}}colOff").text = str(int(x_off_from))
    etree.SubElement(frm, f"{{{XDR}}}row").text    = str(row_from)
    etree.SubElement(frm, f"{{{XDR}}}rowOff").text = str(PADDING)

    to = etree.SubElement(anchor, f"{{{XDR}}}to")
    etree.SubElement(to, f"{{{XDR}}}col").text    = str(col_to)
    etree.SubElement(to, f"{{{XDR}}}colOff").text = str(int(x_off_to))
    etree.SubElement(to, f"{{{XDR}}}row").text    = str(row_from)
    etree.SubElement(to, f"{{{XDR}}}rowOff").text = str(int(row_h_emu))

    return _fill_shape(anchor, sid, text, color, adj, text_color, font_size, bold, geom)

# ─── Shape estándar con col_from/col_to ──────────────────────────────────────
def _shape(parent, sid, text,
           col_from, row_from, col_to, row_to,
           color, adj="default", text_color="FFFFFF",
           font_size=900, bold=False,
           col_width_emu=COL_SEMANA_EMU,
           row_height_emu=ROW_ACT_EMU):

    XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    A   = "http://schemas.openxmlformats.org/drawingml/2006/main"

    anchor = etree.SubElement(parent, f"{{{XDR}}}twoCellAnchor", editAs="oneCell")

    frm = etree.SubElement(anchor, f"{{{XDR}}}from")
    etree.SubElement(frm, f"{{{XDR}}}col").text    = str(col_from)
    etree.SubElement(frm, f"{{{XDR}}}colOff").text = str(PADDING)
    etree.SubElement(frm, f"{{{XDR}}}row").text    = str(row_from)
    etree.SubElement(frm, f"{{{XDR}}}rowOff").text = str(PADDING)

    to = etree.SubElement(anchor, f"{{{XDR}}}to")
    etree.SubElement(to, f"{{{XDR}}}col").text    = str(col_to)
    etree.SubElement(to, f"{{{XDR}}}colOff").text = str(col_width_emu - PADDING)
    etree.SubElement(to, f"{{{XDR}}}row").text    = str(row_to)
    etree.SubElement(to, f"{{{XDR}}}rowOff").text = str(row_height_emu - PADDING)

    return _fill_shape(anchor, sid, text, color, adj, text_color, font_size, bold)

# ─── Rellena el interior de un anchor con sp, spPr, txBody ──────────────────
def _fill_shape(anchor, sid, text, color, adj, text_color, font_size, bold,
                geom="roundRect"):
    XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    A   = "http://schemas.openxmlformats.org/drawingml/2006/main"

    sp     = etree.SubElement(anchor, f"{{{XDR}}}sp", macro="", textlink="")
    nvSpPr = etree.SubElement(sp, f"{{{XDR}}}nvSpPr")
    etree.SubElement(nvSpPr, f"{{{XDR}}}cNvPr", id=str(sid), name=f"shape{sid}")
    etree.SubElement(nvSpPr, f"{{{XDR}}}cNvSpPr")

    spPr     = etree.SubElement(sp, f"{{{XDR}}}spPr")
    prstGeom = etree.SubElement(spPr, f"{{{A}}}prstGeom", prst=geom)
    avLst    = etree.SubElement(prstGeom, f"{{{A}}}avLst")
    if adj != "default":
        etree.SubElement(avLst, f"{{{A}}}gd", name="adj", fmla=f"val {adj}")

    sf = etree.SubElement(spPr, f"{{{A}}}solidFill")
    etree.SubElement(sf, f"{{{A}}}srgbClr", val=color)
    ln = etree.SubElement(spPr, f"{{{A}}}ln")
    etree.SubElement(ln, f"{{{A}}}noFill")

    txBody = etree.SubElement(sp, f"{{{XDR}}}txBody")
    etree.SubElement(txBody, f"{{{A}}}bodyPr", wrap="square", rtlCol="0", anchor="ctr")
    etree.SubElement(txBody, f"{{{A}}}lstStyle")

    p = etree.SubElement(txBody, f"{{{A}}}p")
    etree.SubElement(p, f"{{{A}}}pPr", algn="ctr")

    if text:
        lines = str(text).split("\n")
        for li, line in enumerate(lines):
            if li > 0:
                etree.SubElement(p, f"{{{A}}}br")
            r    = etree.SubElement(p, f"{{{A}}}r")
            attrs = {"lang": "es-CO", "sz": str(font_size), "dirty": "0"}
            if bold:
                attrs["b"] = "1"
            rPr = etree.SubElement(r, f"{{{A}}}rPr", **attrs)
            sf2 = etree.SubElement(rPr, f"{{{A}}}solidFill")
            etree.SubElement(sf2, f"{{{A}}}srgbClr", val=text_color)
            etree.SubElement(rPr, f"{{{A}}}latin", typeface="Calibri")
            etree.SubElement(r, f"{{{A}}}t").text = line

    etree.SubElement(anchor, f"{{{XDR}}}clientData")
    return sid + 1
