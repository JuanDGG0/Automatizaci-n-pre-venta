"""
generator.py — Generador de propuestas Periferia IT v5
=======================================================
v5: Nueva plantilla FDA de un solo slide con dos secciones:
  - Izquierda: 6 grupos (bullets) para FDA técnico de la torre
  - Card verde derecha: sección QA (CuadroTexto 32)
    · Si include_qa=False → texto fijo "Las pruebas de calidad no hacen parte de esta propuesta"
    · Si include_qa=True  → items de QA editables
Sigue soportando "fda", "perfiles", "ambos"
"""
import io, re, json, zipfile, unicodedata
from copy import deepcopy
from pathlib import Path
from lxml import etree

BASE_DIR = Path(__file__).resolve().parent
FDA_TPL  = BASE_DIR / "templates" / "FUERA_DEL_ALCANCE.pptx"
PERF_TPL = BASE_DIR / "templates" / "PERFILES_POR_TORRE.pptx"

A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
P  = "http://schemas.openxmlformats.org/presentationml/2006/main"
R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

STRIP_PATTERNS = [
    r"ppt/revisionInfo\.xml",
    r"ppt/_rels/.*revisionInfo",
    r"customXml/",
    r"docMetadata/",
]

NO_QA_TEXT = "Las pruebas de calidad no hacen parte de esta propuesta"

# Nombres de los grupos de bullets izquierda (en orden de aparición)
BULLET_GROUPS = ["Grupo 2", "Grupo 16", "Grupo 29", "Grupo 38", "Grupo 41", "Grupo 44"]
# Nombre del cuadro de texto QA (card verde derecha)
QA_TEXTBOX    = "CuadroTexto 32"
# Nombre del título técnico (izquierda) y QA (derecha)
TITLE_TECNICO = "Fuera del Alcance Técnico"
TITLE_QA      = "Fuera del Alcance Pruebas de Calidad"

def _norm(s):
    s = (s or "").strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip()

def _esc(t):
    return (t or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def _load(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}

def _should_strip(name):
    return any(re.search(pat, name) for pat in STRIP_PATTERNS)

def _clean_root_rels(rels_xml):
    root = etree.fromstring(rels_xml)
    for rel in list(root):
        rtype = rel.attrib.get("Type","")
        keep_types = ("officeDocument", "extended-properties", "core-properties", "custom-properties")
        if not any(k in rtype for k in keep_types):
            root.remove(rel)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

# ── Editor del slide FDA (nueva plantilla) ────────────────────────────────────

def _edit_fda_slide(xml_bytes, torre_name, fda_items, include_qa, qa_items):
    """
    Edita el slide único de la nueva plantilla FDA:
      - Cambia el título 'Fuera del Alcance Técnico' por el nombre de la torre
      - Llena los 6 grupos de bullets con fda_items
      - Llena el card QA con qa_items o texto de "no aplica"
    """
    root = etree.fromstring(xml_bytes)

    # 1. Cambiar título técnico (izquierda)
    _replace_title_text(root, TITLE_TECNICO, f"Fuera del Alcance {torre_name}")

    # 2. Llenar bullets izquierda
    _fill_bullets(root, fda_items)

    # 3. Llenar card QA derecha
    if not include_qa:
        _fill_qa_card(root, [NO_QA_TEXT])
    else:
        # Si no hay items custom, usar los del Excel para QA
        items_qa = qa_items if qa_items else _get_fda_items("QA", [])
        _fill_qa_card(root, items_qa if items_qa else [NO_QA_TEXT])

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _replace_title_text(root, old_text, new_text):
    """Busca un shape con texto que contenga old_text y lo reemplaza."""
    for sp in root.iter(f"{{{P}}}sp"):
        txb = sp.find(f"{{{P}}}txBody")
        if txb is None: continue
        full = "".join(t.text or "" for t in txb.findall(f".//{{{A}}}t")).strip()
        if old_text.lower() in full.lower():
            # Reemplazar texto en el primer run del primer párrafo
            ts = txb.findall(f".//{{{A}}}t")
            if ts:
                ts[0].text = new_text
                for t in ts[1:]: t.text = ""
            break


def _fill_bullets(root, items):
    """
    Llena los 6 grupos de bullets (izquierda) con los items.
    Cada grupo tiene 2 shapes: [0] rectángulo blanco con texto, [1] palito verde decorativo.
    Si el grupo está vacío, oculta AMBOS shapes.
    """
    bullet_shapes = []  # lista de (sp_texto, txb, sp_palito_o_None)
    for grp_name in BULLET_GROUPS:
        for grpsp in root.iter(f"{{{P}}}grpSp"):
            nvpr = grpsp.find(f".//{{{P}}}cNvPr")
            gname = nvpr.attrib.get("name","") if nvpr is not None else ""
            if gname == grp_name:
                children_sp = list(grpsp.findall(f"{{{P}}}sp"))
                sp_texto = None
                sp_palito = None
                txb_texto = None
                for child_sp in children_sp:
                    txb = child_sp.find(f"{{{P}}}txBody")
                    if txb is not None and sp_texto is None:
                        sp_texto = child_sp
                        txb_texto = txb
                    elif sp_texto is not None:
                        sp_palito = child_sp  # el segundo es el palito decorativo
                if sp_texto is not None:
                    bullet_shapes.append((sp_texto, txb_texto, sp_palito))
                break

    if not bullet_shapes:
        return

    n_boxes = len(bullet_shapes)
    n_items = len(items) if items else 0

    if n_items == 0:
        # Sin items: dejar el texto original de la plantilla (no ocultar)
        return

    base, extra = divmod(n_items, n_boxes)
    idx = 0
    for i, (sp, txb, sp_palito) in enumerate(bullet_shapes):
        for p_el in list(txb.findall(f"{{{A}}}p")):
            txb.remove(p_el)

        take = base + (1 if i < extra else 0)
        chunk = items[idx:idx+take]
        idx += take

        if chunk:
            for item in chunk:
                p_xml = (
                    f'<a:p xmlns:a="{A}">'
                    f'<a:pPr algn="just"/>'
                    f'<a:r>'
                    f'<a:rPr lang="es-MX" sz="1465" dirty="0">'
                    f'<a:solidFill><a:srgbClr val="333333"/></a:solidFill>'
                    f'<a:cs typeface="Calibri"/>'
                    f'</a:rPr>'
                    f'<a:t>{_esc(item)}</a:t>'
                    f'</a:r>'
                    f'</a:p>'
                )
                txb.append(etree.fromstring(p_xml))
        else:
            _clear_and_set_invisible(sp, txb)
            if sp_palito is not None:
                _hide_shape(sp_palito)


def _hide_shape(sp):
    """Oculta un shape poniéndolo transparente."""
    spPr = sp.find(f"{{{P}}}spPr")
    if spPr is None:
        spPr = etree.SubElement(sp, f"{{{P}}}spPr")
    for tag in [f"{{{A}}}solidFill", f"{{{A}}}noFill", f"{{{A}}}gradFill"]:
        for el in spPr.findall(tag):
            spPr.remove(el)
    solid = etree.SubElement(spPr, f"{{{A}}}solidFill")
    srgb  = etree.SubElement(solid, f"{{{A}}}srgbClr")
    srgb.attrib["val"] = "FFFFFF"
    alpha = etree.SubElement(srgb, f"{{{A}}}alpha")
    alpha.attrib["val"] = "0"
    style = sp.find(f"{{{P}}}style")
    if style is not None:
        fillRef = style.find(f"{{{A}}}fillRef")
        if fillRef is not None:
            fillRef.attrib["idx"] = "0"
            for child in list(fillRef): fillRef.remove(child)
            schClr = etree.SubElement(fillRef, f"{{{A}}}srgbClr")
            schClr.attrib["val"] = "FFFFFF"


def _clear_and_set_invisible(sp, txb):
    """Limpia el texto y hace el shape completamente invisible (transparente)."""
    for p_el in list(txb.findall(f"{{{A}}}p")):
        txb.remove(p_el)
    p_xml = (
        f'<a:p xmlns:a="{A}">'
        f'<a:r><a:rPr lang="es-MX" sz="1465" dirty="0">'
        f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>'
        f'<a:t> </a:t></a:r></a:p>'
    )
    txb.append(etree.fromstring(p_xml))

    spPr = sp.find(f"{{{P}}}spPr")
    if spPr is None:
        spPr = etree.SubElement(sp, f"{{{P}}}spPr")

    # Remover cualquier fill existente y poner solidFill transparente
    for tag in [f"{{{A}}}solidFill", f"{{{A}}}noFill", f"{{{A}}}gradFill", f"{{{A}}}pattFill"]:
        for el in spPr.findall(tag):
            spPr.remove(el)
    # Usar solidFill con alpha=0 (transparente) en lugar de noFill
    # para que el style.fillRef no tome control
    solid = etree.SubElement(spPr, f"{{{A}}}solidFill")
    srgb  = etree.SubElement(solid, f"{{{A}}}srgbClr")
    srgb.attrib["val"] = "FFFFFF"
    alpha = etree.SubElement(srgb, f"{{{A}}}alpha")
    alpha.attrib["val"] = "0"  # 0% opacidad = invisible

    # También neutralizar el style.fillRef para que no interfiera
    style = sp.find(f"{{{P}}}style")
    if style is not None:
        fillRef = style.find(f"{{{A}}}fillRef")
        if fillRef is not None:
            # Cambiar a idx=0 (sin fill de estilo) y color transparente
            fillRef.attrib["idx"] = "0"
            for child in list(fillRef):
                fillRef.remove(child)
            schClr = etree.SubElement(fillRef, f"{{{A}}}srgbClr")
            schClr.attrib["val"] = "FFFFFF"

    # Sin borde
    ln = spPr.find(f"{{{A}}}ln")
    if ln is None:
        ln = etree.SubElement(spPr, f"{{{A}}}ln")
    for nf in ln.findall(f"{{{A}}}noFill"):
        ln.remove(nf)
    etree.SubElement(ln, f"{{{A}}}noFill")


def _fill_qa_card(root, qa_items):
    """
    Llena el CuadroTexto 32 (card verde derecha) con los qa_items.
    Cada item es un bullet blanco sobre el fondo verde.
    """
    for sp in root.iter(f"{{{P}}}sp"):
        nvpr = sp.find(f".//{{{P}}}cNvPr")
        name = nvpr.attrib.get("name","") if nvpr is not None else ""
        if name == QA_TEXTBOX:
            txb = sp.find(f"{{{P}}}txBody")
            if txb is None: break

            # Limpiar párrafos existentes
            for p_el in list(txb.findall(f"{{{A}}}p")):
                txb.remove(p_el)

            for item in qa_items:
                p_xml = (
                    f'<a:p xmlns:a="{A}">'
                    f'<a:pPr marL="228389" indent="-228389">'
                    f'<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>'
                    f'<a:buChar char="&#8226;"/>'
                    f'</a:pPr>'
                    f'<a:r>'
                    f'<a:rPr lang="es-ES" sz="1465" dirty="0">'
                    f'<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'
                    f'<a:latin typeface="Arial"/>'
                    f'<a:cs typeface="Arial"/>'
                    f'</a:rPr>'
                    f'<a:t>{_esc(item)}</a:t>'
                    f'</a:r>'
                    f'</a:p>'
                )
                txb.append(etree.fromstring(p_xml))
            break


# ── Construcción del ZIP ──────────────────────────────────────────────────────

def _build_simple_zip(files, slide_key, edited_xml, out_path):
    """
    Genera un PPTX con el slide único de la plantilla FDA, con el XML editado.
    Como la plantilla solo tiene 1 slide, no necesitamos filtrar slides.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            if _should_strip(name):
                continue
            if name == "_rels/.rels":
                zout.writestr(name, _clean_root_rels(data))
                continue
            if name == slide_key:
                zout.writestr(name, edited_xml)
                continue
            zout.writestr(name, data)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(buf.getvalue())


# ── Perfiles (sin cambios respecto a v4) ─────────────────────────────────────

def _slide_title(xml_bytes):
    try:
        root = etree.fromstring(xml_bytes)
        paras = []
        for para in root.findall(f".//{{{A}}}p"):
            texts = [t.text or "" for t in para.findall(f".//{{{A}}}t")]
            full = "".join(texts).strip()
            if full: paras.append(full)
        if not paras: return ""
        for txt in paras:
            if re.search(r"Perfiles\s*[-]\s*TORRE", txt, re.IGNORECASE): return txt
        for txt in paras:
            if re.search(r"FUERA DEL ALCANCE.{0,10}TORRE", txt, re.IGNORECASE): return txt
        for txt in paras:
            u = txt.upper()
            if any(k in u for k in ["FUERA DEL ALCANCE - TORRES","LEVANTAMIENTO DE INFORMACION",
                                     "ROADMAP","GRACIAS","GENERAL APLICABLE"]): return txt
        return paras[0]
    except Exception: return ""

def _index(files):
    prs  = etree.fromstring(files["ppt/presentation.xml"])
    rels = etree.fromstring(files["ppt/_rels/presentation.xml.rels"])
    rid_to_target = {el.attrib["Id"]: el.attrib["Target"]
                     for el in rels.findall(f"{{{PR}}}Relationship")}
    slides = []
    for sld_el in prs.findall(f".//{{{P}}}sldId"):
        rid    = sld_el.attrib.get(f"{{{R}}}id")
        target = rid_to_target.get(rid, "")
        m = re.search(r"slide(\d+)\.xml$", target)
        if not m: continue
        num  = int(m.group(1))
        key  = f"ppt/{target}"
        title = _slide_title(files.get(key, b""))
        slides.append({"num": num, "rid": rid, "target": target,
                        "key": key, "title": title, "title_norm": _norm(title)})
    return slides

def _find_fixed(slides, keys):
    result = []
    for key in keys:
        kn = _norm(key)
        for i, s in enumerate(slides):
            if kn in s["title_norm"] or s["title_norm"] in kn:
                result.append(i); break
    return result

def _find_torre(slides, name, prefix):
    pn = _norm(prefix)
    nn = _norm(name).split("/")[0].strip()
    return [i for i, s in enumerate(slides)
            if pn in s["title_norm"] and nn in s["title_norm"]]

def _set_text(txb, text):
    ts = txb.findall(f".//{{{A}}}t")
    if ts:
        ts[0].text = text
        for t in ts[1:]: t.text = ""

def _edit_perf(xml_bytes, profiles):
    if not profiles: return xml_bytes
    root = etree.fromstring(xml_bytes)
    main_slots, shadow_nomes = [], []
    for grp in root.iter(f"{{{P}}}grpSp"):
        ntb = dtb = None
        for sp in grp.iter(f"{{{P}}}sp"):
            tb = sp.find(f"{{{P}}}txBody")
            if tb is None: continue
            txt = "".join(t.text or "" for t in tb.findall(f".//{{{A}}}t")).strip()
            if not txt: continue
            if len(txt) < 70 and txt == txt.upper() and ntb is None: ntb = tb
            elif len(txt) >= 30 and dtb is None: dtb = tb
        if ntb is not None and dtb is not None: main_slots.append((ntb, dtb))
        elif ntb is not None: shadow_nomes.append(ntb)
    for i, (ntb, dtb) in enumerate(main_slots):
        if i >= len(profiles): break
        p   = profiles[i] or {}
        rol = (p.get("rol") or "").strip()
        dsc = (p.get("desc") or "").strip()
        if rol and ntb is not None: _set_text(ntb, rol.upper())
        if dsc and dtb is not None: _set_text(dtb, dsc)
        if i < len(shadow_nomes) and rol: _set_text(shadow_nomes[i], rol.upper())
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

def _build_zip(files, slides, selected_indices, out_path):
    selected  = [slides[i] for i in selected_indices]
    sel_nums  = {s["num"] for s in selected}

    valid_notes = set()
    for num in sel_nums:
        rels_key = f"ppt/slides/_rels/slide{num}.xml.rels"
        if rels_key in files:
            try:
                rels = etree.fromstring(files[rels_key])
                for r in rels:
                    t = r.attrib.get("Target", "")
                    if "notesSlide" in t:
                        valid_notes.add(t.split("/")[-1])
            except Exception:
                pass

    prs_root = etree.fromstring(files["ppt/presentation.xml"])
    sldIdLst = prs_root.find(f"{{{P}}}sldIdLst")
    orig_slds = {el.attrib.get(f"{{{R}}}id"): deepcopy(el) for el in list(sldIdLst)}
    for child in list(sldIdLst): sldIdLst.remove(child)
    for s in selected:
        orig = orig_slds.get(s["rid"])
        if orig is not None: sldIdLst.append(orig)
    new_prs = etree.tostring(prs_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    prs_rels_root = etree.fromstring(files["ppt/_rels/presentation.xml.rels"])
    for rel in list(prs_rels_root):
        target = rel.attrib.get("Target","")
        if _should_strip(target) or "revisionInfo" in target or "customXml" in target:
            prs_rels_root.remove(rel); continue
        m = re.search(r"slide(\d+)\.xml$", target)
        if m and "slideLayout" not in target and "slideMaster" not in target:
            if int(m.group(1)) not in sel_nums:
                prs_rels_root.remove(rel)
    new_prs_rels = etree.tostring(prs_rels_root, xml_declaration=True,
                                   encoding="UTF-8", standalone=True)

    ct_root = etree.fromstring(files["[Content_Types].xml"])
    for override in list(ct_root.findall(f"{{{CT}}}Override")):
        part = override.attrib.get("PartName","")
        part_clean = part.lstrip("/")
        if _should_strip(part_clean): ct_root.remove(override); continue
        mn = re.search(r"notesSlides/(notesSlide[^/]+\.xml)$", part)
        if mn:
            if mn.group(1) not in valid_notes: ct_root.remove(override)
            continue
        m = re.search(r"slide(\d+)\.xml$", part)
        if m and "slideLayout" not in part and "slideMaster" not in part and "notes" not in part.lower():
            if int(m.group(1)) not in sel_nums:
                ct_root.remove(override)
    new_ct = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            if _should_strip(name): continue
            if name == "ppt/presentation.xml":
                zout.writestr(name, new_prs); continue
            if name == "ppt/_rels/presentation.xml.rels":
                zout.writestr(name, new_prs_rels); continue
            if name == "[Content_Types].xml":
                zout.writestr(name, new_ct); continue
            if name == "_rels/.rels":
                zout.writestr(name, _clean_root_rels(data)); continue
            m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
            if m:
                if int(m.group(1)) in sel_nums: zout.writestr(name, data)
                continue
            m = re.match(r"ppt/slides/_rels/slide(\d+)\.xml\.rels$", name)
            if m:
                if int(m.group(1)) in sel_nums: zout.writestr(name, data)
                continue
            mn = re.match(r"ppt/notesSlides/(notesSlide[^/]+\.xml)$", name)
            if mn:
                if mn.group(1) in valid_notes: zout.writestr(name, data)
                continue
            mn = re.match(r"ppt/notesSlides/_rels/(notesSlide[^/]+\.xml)\.rels$", name)
            if mn:
                if mn.group(1) in valid_notes: zout.writestr(name, data)
                continue
            zout.writestr(name, data)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(buf.getvalue())

def _select_perf(files_pf, slides_pf, torres):
    fs = _find_fixed(slides_pf, ["LEVANTAMIENTO DE INFORMACION","ROADMAP DEL SERVICIO"])
    fe = _find_fixed(slides_pf, ["GRACIAS"])
    fa = set(fs + fe)
    tidxs = []
    for tc in torres:
        name     = tc.get("name","")
        profiles = tc.get("perfiles") or []
        cands = [i for i in _find_torre(slides_pf, name, "PERFILES - TORRE") if i not in fa]
        if not cands: print(f"[WARN PERF] no encontrado: {name!r}"); continue
        if profiles:
            PPS = 4
            files_pf = dict(files_pf)
            n_slides_needed = max(1, -(-len(profiles) // PPS))
            slides_to_use = cands[:n_slides_needed]
            while len(slides_to_use) < n_slides_needed:
                slides_to_use.append(cands[0])
            for si, idx in enumerate(slides_to_use):
                chunk = profiles[si*PPS:(si+1)*PPS]
                if chunk:
                    files_pf[slides_pf[idx]["key"]] = _edit_perf(files_pf[slides_pf[idx]["key"]], chunk)
                tidxs.append(idx)
        else:
            tidxs.extend(cands)
    if not tidxs: raise ValueError("Ninguna torre Perfiles encontrada.")
    return files_pf, fs + tidxs + fe


# ── Generación FDA (nueva lógica de slide único) ──────────────────────────────

def _get_fda_slide_key(files):
    """Devuelve la key del primer (y único) slide del template FDA."""
    prs  = etree.fromstring(files["ppt/presentation.xml"])
    rels = etree.fromstring(files["ppt/_rels/presentation.xml.rels"])
    rid_to_target = {el.attrib["Id"]: el.attrib["Target"]
                     for el in rels.findall(f"{{{PR}}}Relationship")}
    for sld_el in prs.findall(f".//{{{P}}}sldId"):
        rid    = sld_el.attrib.get(f"{{{R}}}id")
        target = rid_to_target.get(rid, "")
        if target:
            return f"ppt/{target}"
    return None

def _gen_fda(torres, out_path):
    """
    Genera un PPTX FDA por cada torre seleccionada.
    Nueva lógica: un slide por torre, editando el slide único de la plantilla.
    Si hay múltiples torres, genera múltiples slides (uno por torre) en el mismo PPTX.
    """
    files = _load(FDA_TPL)
    slide_key = _get_fda_slide_key(files)
    if not slide_key:
        raise ValueError("No se encontró slide en la plantilla FDA")

    original_xml = files[slide_key]

    # Si hay múltiples torres: duplicar el slide para cada una
    if len(torres) > 1:
        _gen_fda_multi(files, slide_key, original_xml, torres, out_path)
    else:
        # Una sola torre: editar el slide único directamente
        tc = torres[0] if torres else {}
        torre_name  = tc.get("name", "")
        fda_items   = _get_fda_items(torre_name, tc.get("fda_items") or [])
        include_qa  = tc.get("include_qa", True)
        qa_items    = tc.get("qa_items") or []

        edited_xml = _edit_fda_slide(original_xml, torre_name, fda_items, include_qa, qa_items)
        files = dict(files)
        files[slide_key] = edited_xml
        _build_simple_zip(files, slide_key, edited_xml, out_path)

    return str(out_path)


def _gen_fda_multi(files_orig, slide_key, original_xml, torres, out_path):
    """
    Genera múltiples slides FDA (uno por torre) en un solo PPTX.
    Duplica el slide de la plantilla tantas veces como torres haya.
    """
    # Parsear la presentación base
    prs_xml   = files_orig["ppt/presentation.xml"]
    prs_rels  = files_orig["ppt/_rels/presentation.xml.rels"]

    prs_root      = etree.fromstring(prs_xml)
    prs_rels_root = etree.fromstring(prs_rels)

    # Encontrar el slide original
    original_match = re.search(r"slide(\d+)\.xml$", slide_key)
    if not original_match:
        raise ValueError(f"slide_key inválido: {slide_key}")
    orig_num = int(original_match.group(1))

    # Encontrar el rId del slide original en presentation.xml.rels
    orig_rid = None
    for rel in prs_rels_root.findall(f"{{{PR}}}Relationship"):
        target = rel.attrib.get("Target","")
        if f"slide{orig_num}.xml" in target:
            orig_rid = rel.attrib.get("Id")
            break

    # Obtener rId máximo existente
    existing_rids = [int(re.sub(r'\D','',el.attrib.get("Id","0")))
                     for el in prs_rels_root if el.attrib.get("Id","").startswith("rId")]
    next_rid_num = max(existing_rids, default=10) + 1

    # Obtener sldId máximo
    sldIdLst = prs_root.find(f"{{{P}}}sldIdLst")
    existing_ids = [int(el.attrib.get("id","0")) for el in list(sldIdLst)]
    next_sld_id = max(existing_ids, default=256) + 1

    # Obtener rels del slide original para copiarlos
    orig_rels_key = f"ppt/slides/_rels/slide{orig_num}.xml.rels"

    # Limpiar el sldIdLst existente (vamos a reconstruirlo)
    orig_sld_el = None
    for el in list(sldIdLst):
        rid = el.attrib.get(f"{{{R}}}id")
        if rid == orig_rid:
            orig_sld_el = deepcopy(el)
        sldIdLst.remove(el)

    # Remover rel original de prs_rels (lo reemplazaremos con uno por torre)
    for rel in list(prs_rels_root):
        target = rel.attrib.get("Target","")
        if f"slide{orig_num}.xml" in target and "slideLayout" not in target:
            prs_rels_root.remove(rel)

    files_out = dict(files_orig)
    ct_root   = etree.fromstring(files_orig["[Content_Types].xml"])

    # Eliminar Override del slide original en content_types
    for override in list(ct_root.findall(f"{{{CT}}}Override")):
        part = override.attrib.get("PartName","")
        if f"slide{orig_num}.xml" in part and "Layout" not in part and "Master" not in part:
            ct_root.remove(override)

    new_slide_nums = []

    for i, tc in enumerate(torres):
        torre_name  = tc.get("name", "")
        fda_items   = _get_fda_items(torre_name, tc.get("fda_items") or [])
        include_qa  = tc.get("include_qa", True)
        qa_items    = tc.get("qa_items") or []

        # Usar OFFSET alto para evitar colisión con el slide original
        new_num = 500 + i
        new_rid = f"rId{next_rid_num}"
        new_sld_id = next_sld_id

        next_rid_num += 1
        next_sld_id  += 1
        new_slide_nums.append(new_num)

        # Editar XML del slide para esta torre
        edited_xml = _edit_fda_slide(original_xml, torre_name, fda_items, include_qa, qa_items)
        new_key = f"ppt/slides/slide{new_num}.xml"
        files_out[new_key] = edited_xml

        # Copiar rels del slide
        if orig_rels_key in files_orig:
            new_rels_key = f"ppt/slides/_rels/slide{new_num}.xml.rels"
            files_out[new_rels_key] = files_orig[orig_rels_key]

        # Agregar sldId en presentation
        new_sld_el = etree.SubElement(sldIdLst, f"{{{P}}}sldId")
        new_sld_el.attrib["id"] = str(new_sld_id)
        new_sld_el.attrib[f"{{{R}}}id"] = new_rid

        # Agregar rel en presentation.xml.rels
        new_rel = etree.SubElement(prs_rels_root, f"{{{PR}}}Relationship")
        new_rel.attrib["Id"]     = new_rid
        new_rel.attrib["Type"]   = f"{R}/slide"
        new_rel.attrib["Target"] = f"slides/slide{new_num}.xml"

        # Agregar a content_types
        ct_el = etree.SubElement(ct_root, f"{{{CT}}}Override")
        ct_el.attrib["PartName"]    = f"/ppt/slides/slide{new_num}.xml"
        ct_el.attrib["ContentType"] = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

    # Actualizar XMLs en files_out
    files_out["ppt/presentation.xml"]          = etree.tostring(prs_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    files_out["ppt/_rels/presentation.xml.rels"] = etree.tostring(prs_rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    files_out["[Content_Types].xml"]            = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # Remover el slide original del dict (ya no lo necesitamos)
    files_out.pop(slide_key, None)
    files_out.pop(orig_rels_key, None)

    # Escribir ZIP
    sel_nums = set(new_slide_nums)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files_out.items():
            if _should_strip(name): continue
            if name == "_rels/.rels":
                zout.writestr(name, _clean_root_rels(data)); continue
            zout.writestr(name, data)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(buf.getvalue())


def _gen_perf(torres, out_path):
    files  = _load(PERF_TPL)
    slides = _index(files)
    files, sel = _select_perf(files, slides, torres)
    _build_zip(files, slides, sel, out_path)
    return str(out_path)


# ── Combinado: FDA (nueva plantilla) + Perfiles en un solo PPTX ──────────────

# ── Carga de items FDA desde Excel ───────────────────────────────────────────

FDA_EXCEL = BASE_DIR / "Fuera_del_Alcance_Estructura_Actualizada.xlsx"

def _load_fda_db():
    """Carga el Excel y devuelve dict {TORRE: [items...]}"""
    try:
        import openpyxl, unicodedata as ud
        wb = openpyxl.load_workbook(FDA_EXCEL)
        ws = wb.active
        db = {}
        current = None
        for row in ws.iter_rows(values_only=True):
            torre, item = (row[0] or "").strip(), (row[1] or "").strip() if len(row) > 1 else ""
            if torre and torre.upper() != "TORRE":
                current = torre.upper()
            if current and item:
                db.setdefault(current, []).append(item)
        return db
    except Exception:
        return {}

def _get_fda_items(torre_name, custom_items):
    """Retorna items: custom si los hay, sino 6 aleatorios del Excel."""
    import random
    if custom_items:
        return custom_items
    db = _load_fda_db()
    # Normalizar nombre de torre para buscar en el Excel
    name_norm = torre_name.strip().upper()
    candidates = db.get(name_norm, [])
    if not candidates:
        # Buscar coincidencia parcial
        for key in db:
            if name_norm in key or key in name_norm:
                candidates = db[key]
                break
    if not candidates:
        return []
    n = min(6, len(candidates))
    return random.sample(candidates, n)


def _gen_combinado(torres, out_path):
    """
    Genera un único PPTX:
      [slides FDA uno por torre] + [slides Perfiles por torre]
    Usa el template de Perfiles como base y agrega los slides FDA al inicio.
    Los slides FDA se renombran con OFFSET=2000 para no colisionar.
    """
    OFFSET = 2000

    # — Preparar slides Perfiles PRIMERO (necesitamos sel_pf_nums y files_pf) —
    files_pf  = _load(PERF_TPL)
    slides_pf = _index(files_pf)
    files_pf, sel_pf = _select_perf(files_pf, slides_pf, torres)
    sel_pf_slides = [slides_pf[i] for i in sel_pf]
    sel_pf_nums   = {s["num"] for s in sel_pf_slides}

    # — Preparar slides FDA editados —
    files_fda = _load(FDA_TPL)
    fda_slide_key = _get_fda_slide_key(files_fda)
    if not fda_slide_key:
        raise ValueError("No se encontró slide en la plantilla FDA")
    orig_fda_xml = files_fda[fda_slide_key]
    orig_fda_match = re.search(r"slide(\d+)\.xml$", fda_slide_key)
    orig_fda_num = int(orig_fda_match.group(1))
    orig_fda_rels_key = f"ppt/slides/_rels/slide{orig_fda_num}.xml.rels"

    # Averiguar qué slideLayout usa el primer slide de Perfiles (para reutilizarlo)
    pf_first_slide_num = min(sel_pf_nums) if sel_pf_nums else None
    pf_layout_target = None
    if pf_first_slide_num:
        pf_rels_key = f"ppt/slides/_rels/slide{pf_first_slide_num}.xml.rels"
        if pf_rels_key in files_pf:
            try:
                pf_rels_root_tmp = etree.fromstring(files_pf[pf_rels_key])
                for rel in pf_rels_root_tmp:
                    if "slideLayout" in rel.attrib.get("Target",""):
                        pf_layout_target = rel.attrib["Target"]
                        break
            except Exception:
                pass

    def _make_fda_rels(pf_layout_target):
        """Genera rels mínimos para un slide FDA apuntando al layout de Perfiles."""
        if pf_layout_target:
            return (
                f'<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
                f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
                f'Target="{pf_layout_target}"/>'
                f'</Relationships>'
            ).encode("utf-8")
        return b""

    fda_slides_data = []  # lista de (new_num, xml_bytes, rels_bytes)
    for i, tc in enumerate(torres):
        new_num    = orig_fda_num + OFFSET + i
        edited_xml = _edit_fda_slide(
            orig_fda_xml,
            tc.get("name",""),
            _get_fda_items(tc.get("name",""), tc.get("fda_items") or []),
            tc.get("include_qa", True),
            tc.get("qa_items") or []
        )
        rels_bytes = _make_fda_rels(pf_layout_target)
        fda_slides_data.append((new_num, edited_xml, rels_bytes))

    # — Construir ZIP usando Perfiles como base —
    # presentation.xml de Perfiles ya tiene los slides seleccionados; vamos a
    # reconstruirlo metiendo los FDA al inicio.
    prs_root      = etree.fromstring(files_pf["ppt/presentation.xml"])
    prs_rels_root = etree.fromstring(files_pf["ppt/_rels/presentation.xml.rels"])
    ct_root       = etree.fromstring(files_pf["[Content_Types].xml"])

    sldIdLst = prs_root.find(f"{{{P}}}sldIdLst")

    # Guardar sldId actuales de Perfiles seleccionados
    pf_sld_elements = []
    for el in list(sldIdLst):
        rid = el.attrib.get(f"{{{R}}}id")
        # Verificar que es un slide seleccionado
        target = next((r.attrib.get("Target","") for r in prs_rels_root
                       if r.attrib.get("Id") == rid), "")
        m = re.search(r"slide(\d+)\.xml$", target)
        if m and int(m.group(1)) in sel_pf_nums:
            pf_sld_elements.append(deepcopy(el))
        sldIdLst.remove(el)

    # Calcular rId y sldId máximos existentes en Perfiles
    existing_rids = [int(re.sub(r'\D','',el.attrib.get("Id","0")))
                     for el in prs_rels_root if re.match(r"rId\d+", el.attrib.get("Id",""))]
    next_rid_num = max(existing_rids, default=50) + 1

    existing_ids = [int(el.attrib.get("id","256")) for el in pf_sld_elements]
    next_sld_id  = max(existing_ids, default=256) + 1

    # — Insertar slides FDA al inicio del sldIdLst —
    fda_rids = []
    for (new_num, _, _) in fda_slides_data:
        new_rid = f"rId{next_rid_num}"
        fda_rids.append(new_rid)
        next_rid_num += 1

        new_sld_el = etree.SubElement(sldIdLst, f"{{{P}}}sldId")
        new_sld_el.attrib["id"] = str(next_sld_id)
        new_sld_el.attrib[f"{{{R}}}id"] = new_rid
        next_sld_id += 1

        new_rel = etree.SubElement(prs_rels_root, f"{{{PR}}}Relationship")
        new_rel.attrib["Id"]     = new_rid
        new_rel.attrib["Type"]   = f"{R}/slide"
        new_rel.attrib["Target"] = f"slides/slide{new_num}.xml"

        ct_el = etree.SubElement(ct_root, f"{{{CT}}}Override")
        ct_el.attrib["PartName"]    = f"/ppt/slides/slide{new_num}.xml"
        ct_el.attrib["ContentType"] = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

    # — Agregar slides Perfiles seleccionados después de los FDA —
    for el in pf_sld_elements:
        sldIdLst.append(el)

    # Limpiar rels de Perfiles no seleccionados y metadata
    for rel in list(prs_rels_root):
        target = rel.attrib.get("Target","")
        if _should_strip(target) or "revisionInfo" in target or "customXml" in target:
            prs_rels_root.remove(rel); continue
        m = re.search(r"slides/slide(\d+)\.xml$", target)
        if m and "slideLayout" not in target and "slideMaster" not in target:
            num = int(m.group(1))
            fda_new_nums = {d[0] for d in fda_slides_data}
            if num not in sel_pf_nums and num not in fda_new_nums:
                prs_rels_root.remove(rel)

    # Limpiar content_types de slides PF no seleccionados
    fda_new_nums = {d[0] for d in fda_slides_data}
    for override in list(ct_root.findall(f"{{{CT}}}Override")):
        part = override.attrib.get("PartName","").lstrip("/")
        if _should_strip(part): ct_root.remove(override); continue
        if "notesSlide" in part.lower() and "notesMaster" not in part.lower():
            ct_root.remove(override); continue
        m = re.search(r"slides/slide(\d+)\.xml$", part)
        if m and "Layout" not in part and "Master" not in part and "notes" not in part.lower():
            num = int(m.group(1))
            if num not in sel_pf_nums and num not in fda_new_nums:
                ct_root.remove(override)

    new_prs      = etree.tostring(prs_root,      xml_declaration=True, encoding="UTF-8", standalone=True)
    new_prs_rels = etree.tostring(prs_rels_root,  xml_declaration=True, encoding="UTF-8", standalone=True)
    new_ct       = etree.tostring(ct_root,        xml_declaration=True, encoding="UTF-8", standalone=True)

    # — Media FDA: siempre prefijamos con fda_ para evitar colisiones con Perfiles —
    fda_media = {k: v for k, v in files_fda.items() if k.startswith("ppt/media/")}
    fda_media_remap = {}  # old_path -> new_path
    for k in fda_media:
        fname = k.split("/")[-1]
        fda_media_remap[k] = f"ppt/media/fda_{fname}"

    # Actualizar rels de cada slide FDA para apuntar a las imágenes renombradas
    updated_fda_slides = []
    orig_rels_key = f"ppt/slides/_rels/slide{orig_fda_num}.xml.rels"
    for (new_num, xml_bytes, _) in fda_slides_data:
        if orig_rels_key in files_fda:
            rels_root = etree.fromstring(files_fda[orig_rels_key])
            for rel in rels_root:
                rtype = rel.attrib.get("Type", "")
                target = rel.attrib.get("Target", "")
                if "image" in rtype and "../media/" in target:
                    fname = target.split("/")[-1]
                    rel.attrib["Target"] = f"../media/fda_{fname}"
                elif "slideLayout" in rtype and pf_layout_target:
                    rel.attrib["Target"] = pf_layout_target
                elif "notesSlide" in rtype:
                    rels_root.remove(rel)
            new_rels = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
        else:
            # Rels mínimos si no existe el original
            new_rels = (
                f'<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>'
                f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + (f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="{pf_layout_target}"/>' if pf_layout_target else "")
                + f'</Relationships>'
            ).encode("utf-8")
        updated_fda_slides.append((new_num, xml_bytes, new_rels))
    fda_slides_data = updated_fda_slides

    # — Escribir ZIP —
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        # Archivos base de Perfiles
        for name, data in files_pf.items():
            if _should_strip(name): continue
            if name == "ppt/presentation.xml":
                zout.writestr(name, new_prs); continue
            if name == "ppt/_rels/presentation.xml.rels":
                zout.writestr(name, new_prs_rels); continue
            if name == "[Content_Types].xml":
                zout.writestr(name, new_ct); continue
            if name == "_rels/.rels":
                zout.writestr(name, _clean_root_rels(data)); continue
            m = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
            if m:
                if int(m.group(1)) in sel_pf_nums: zout.writestr(name, data)
                continue
            m = re.match(r"ppt/slides/_rels/slide(\d+)\.xml\.rels$", name)
            if m:
                if int(m.group(1)) in sel_pf_nums: zout.writestr(name, data)
                continue
            if re.match(r"ppt/notesSlides/", name): continue
            zout.writestr(name, data)

        # Slides FDA con rels actualizados
        for (new_num, xml_bytes, rels_bytes) in fda_slides_data:
            zout.writestr(f"ppt/slides/slide{new_num}.xml", xml_bytes)
            if rels_bytes:
                zout.writestr(f"ppt/slides/_rels/slide{new_num}.xml.rels", rels_bytes)

        # Media FDA renombrada (fda_imageX.png)
        written = set()
        for old_k, new_k in fda_media_remap.items():
            if new_k not in written:
                zout.writestr(new_k, fda_media[old_k])
                written.add(new_k)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(buf.getvalue())
    return str(out_path)


# ── Punto de entrada ──────────────────────────────────────────────────────────

def _check():
    missing = []
    if not FDA_TPL.exists():  missing.append(str(FDA_TPL))
    if not PERF_TPL.exists(): missing.append(str(PERF_TPL))
    if missing:
        raise FileNotFoundError("Plantillas faltantes:\n" +
            "\n".join(f"  x {m}" for m in missing) +
            f"\n\nColocalas en: {BASE_DIR / 'templates'}")

def generate(config, out_dir):
    _check()
    mode   = (config.get("mode") or "fda").strip().lower()
    torres = config.get("torres") or []
    out    = Path(out_dir)
    result = {}

    if mode == "fda":
        result["fda"] = _gen_fda(torres, out / "Periferia_IT_Fuera_del_Alcance.pptx")
    elif mode == "perfiles":
        result["perfiles"] = _gen_perf(torres, out / "Periferia_IT_Perfiles.pptx")
    elif mode == "ambos":
        result["ambos"] = _gen_combinado(torres, out / "Periferia_IT_Propuesta_Completa.pptx")

    return result

if __name__ == "__main__":
    import sys
    cfg = json.loads(sys.argv[1]) if len(sys.argv) > 1 else {}
    out = sys.argv[2] if len(sys.argv) > 2 else "/tmp/out"
    print(json.dumps(generate(cfg, out), ensure_ascii=False))