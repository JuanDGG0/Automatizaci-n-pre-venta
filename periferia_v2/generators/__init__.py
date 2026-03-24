"""
generators/__init__.py
Orquestador principal — llama a los 3 generators en cadena sobre el mismo PPTX.
"""

import io
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

TEMPLATES = {
    'corp':  BASE_DIR / 'templates' / 'CS-FR-012-PROPUESTA_COMERCIAL_PERIFERIA_IT_CORP.pptx',
    'group': BASE_DIR / 'templates' / 'CS-FR-005-PROPUESTA_COMERCIAL_PERIFERIA_IT_GROUP.pptx',
    'cbit':  BASE_DIR / 'templates' / 'CS-FR-011-PROPUESTA_COMERCIAL_CBIT.pptx',
}

from .fda_perfiles           import edit as edit_fda_perfiles
from .consideraciones        import edit as edit_consideraciones
from .cronograma_entregables import edit as edit_cronograma_entregables


def generate(config, out_dir):
    """
    config = {
        filial: 'corp' | 'group' | 'cbit',
        excel_data: { torres: [...], perfiles: [...], cliente: str, proyecto: str },
        torres_seleccionadas: [...],   # si excel vacío
        incluir_qa: True | False,          # ¿QA hace parte de esta propuesta?
        opciones: {
            perfiles: 'excel' | 'genericos',
            fda: 'excel' | 'genericos',
            entregables: 'genericos' | 'manual',
            consideraciones: 'genericos' | 'manual',
            cronograma: 'excel' | 'manual'
        },
        entregables: [...],        # si manual
        consideraciones: [...],    # si manual
    }
    Retorna: { 'propuesta': '/path/to/propuesta.pptx' }
    """
    filial = config.get('filial', 'corp')
    template_path = TEMPLATES.get(filial)

    if not template_path or not template_path.exists():
        raise FileNotFoundError(f'Plantilla no encontrada para filial: {filial} en {template_path}')

    # Cargar plantilla como bytes
    pptx_bytes = template_path.read_bytes()

    # ── Heidy: FDA + Perfiles ──────────────────────────────────────────
    pptx_bytes = edit_fda_perfiles(pptx_bytes, config)

    # ── Juan: Consideraciones ──────────────────────────────────────────
    pptx_bytes = edit_consideraciones(pptx_bytes, config)

    # ── José: Cronograma + Entregables ─────────────────────────────────
    pptx_bytes = edit_cronograma_entregables(pptx_bytes, config)

    # ── Guardar resultado ──────────────────────────────────────────────
    out_path = Path(out_dir) / f'Propuesta_Periferia_{filial.upper()}.pptx'
    out_path.write_bytes(pptx_bytes)

    return {'propuesta': str(out_path)}
