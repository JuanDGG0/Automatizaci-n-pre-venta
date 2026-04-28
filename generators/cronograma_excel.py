import io
import math
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

HORAS_POR_SEMANA = 40
SEMANAS_POR_MES = 4
SEMANAS_POR_SPRINT = 2


def generate_cronograma(config) -> bytes:
    actividades = config.get("actividades", [])

    if not actividades:
        raise ValueError("No hay actividades para generar cronograma")

    # Convertir horas → semanas
    for act in actividades:
        semanas = max(1, math.ceil(act["horas"] / HORAS_POR_SEMANA))
        act["start"] = 0
        act["end"] = semanas

    max_semanas = max(a["end"] for a in actividades)

    wb = Workbook()
    ws = wb.active
    ws.title = "Cronograma"

    construir_header(ws, max_semanas)
    pintar_actividades(ws, actividades)

    buffer = io.BytesIO()
    wb.save(buffer)

    return buffer.getvalue()


# ─────────────────────────────────────────────
# 🧱 HEADER (Meses + Semanas + Sprints)
# ─────────────────────────────────────────────
def construir_header(ws, total_semanas):
    # Fila 1 → Meses
    # Fila 2 → Semanas
    # Fila 3 → Sprints

    ws.cell(row=1, column=1, value="")
    ws.cell(row=2, column=1, value="TORRE")

    col_offset = 2

    # 🔹 Meses
    mes = 1
    for i in range(0, total_semanas, SEMANAS_POR_MES):
        start_col = col_offset + i
        end_col = min(start_col + SEMANAS_POR_MES - 1, col_offset + total_semanas - 1)

        ws.merge_cells(start_row=1, start_column=start_col,
                       end_row=1, end_column=end_col)

        cell = ws.cell(row=1, column=start_col)
        cell.value = f"Mes {mes}"
        cell.alignment = Alignment(horizontal="center")
        cell.font = Font(bold=True)

        mes += 1

    # 🔹 Semanas
    for i in range(total_semanas):
        col = col_offset + i
        ws.cell(row=2, column=col, value=f"S{i+1}")

    # 🔹 Sprints
    sprint = 0
    for i in range(0, total_semanas, SEMANAS_POR_SPRINT):
        start_col = col_offset + i
        end_col = min(start_col + SEMANAS_POR_SPRINT - 1, col_offset + total_semanas - 1)

        ws.merge_cells(start_row=3, start_column=start_col,
                       end_row=3, end_column=end_col)

        cell = ws.cell(row=3, column=start_col)
        cell.value = f"Sprint {sprint}"
        cell.alignment = Alignment(horizontal="center")

        sprint += 1

    # Estilos
    ws.column_dimensions["A"].width = 30

    for col in range(2, total_semanas + 2):
        ws.column_dimensions[get_column_letter(col)].width = 4  # ← fix: chr() → get_column_letter()


# ─────────────────────────────────────────────
# 🎨 PINTAR BARRAS
# ─────────────────────────────────────────────
def pintar_actividades(ws, actividades):
    colores = generar_colores_por_torre(actividades)

    row = 4  # empieza después del header

    for act in actividades:
        ws.cell(row=row, column=1, value=act["torre"])

        fill = PatternFill(
            start_color=colores[act["torre"]],
            end_color=colores[act["torre"]],
            fill_type="solid"
        )

        for col in range(act["start"], act["end"]):
            ws.cell(row=row, column=col + 2).fill = fill

        row += 1


# ─────────────────────────────────────────────
# 🎨 COLORES
# ─────────────────────────────────────────────
def generar_colores_por_torre(actividades):
    base_colors = [
        "00C853",  # verde
        "2962FF",  # azul
        "AA00FF",  # morado
        "FF6D00",  # naranja
        "D50000",  # rojo
    ]

    torres = list({a["torre"] for a in actividades})

    return {
        torre: base_colors[i % len(base_colors)]
        for i, torre in enumerate(torres)
    }