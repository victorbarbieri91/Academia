"""Gera registro-treino.xlsx — planilha de log de treinos.

Estrutura:
  - Aba "Registro": uma linha por série feita (Data, Treino, Exercício, Série, Reps, Carga, Obs)
  - Aba "Treino A" e "Treino B": configuração de referência (exercícios padrão, séries x reps)
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "/home/user/Academia/registro-treino.xlsx"

# -------- Treino A (Posterior) --------
TREINO_A = [
    ("Costas",      "Puxada Alta (Pulldown)",                 "Máquina", "4 x 10-12"),
    ("Costas",      "Remada Sentada (Cabo)",                  "Polia",   "3 x 12"),
    ("Costas",      "Remada Máquina (Articulada)",            "Máquina", "3 x 10"),
    ("Costas",      "Pulldown Braço Estendido",               "Polia",   "3 x 15"),
    ("Posterior",   "Mesa Flexora (Deitada)",                 "Máquina", "3 x 15 leve"),
    ("Posterior",   "Cadeira Flexora Sentada (Unilateral)",   "Máquina", "3 x 12 cada"),
    ("Glúteo",      "Elevação Pélvica Máquina (Glute Drive)", "Máquina", "3 x 12-15"),
    ("Bíceps",      "Rosca Direta (Halter ou Barra W)",       "Livre",   "3 x 10-12"),
    ("Bíceps",      "Rosca Martelo (Hammer)",                 "Livre",   "3 x 12"),
    ("Panturrilha", "Panturrilha em Pé (Máquina)",            "Máquina", "4 x 15"),
]

# -------- Treino B (Anterior) --------
TREINO_B = [
    ("Peito",       "Supino Máquina (Chest Press)",           "Máquina", "4 x 10"),
    ("Peito",       "Voador (Peck Deck)",                     "Máquina", "3 x 12"),
    ("Peito",       "Crossover na Polia",                     "Polia",   "3 x 12"),
    ("Quadríceps",  "Leg Press 45°",                          "Máquina", "4 x 10-12"),
    ("Quadríceps",  "Cadeira Extensora",                      "Máquina", "3 x 12"),
    ("Adutor",      "Cadeira Adutora",                        "Máquina", "3 x 15"),
    ("Ombro",       "Desenvolvimento Máquina",                "Máquina", "3 x 10-12"),
    ("Ombro",       "Elevação Lateral (Máquina/Polia)",       "Máquina", "3 x 12-15"),
    ("Tríceps",     "Tríceps Testa (Halter/Barra W)",         "Livre",   "3 x 10-12"),
    ("Tríceps",     "Tríceps Corda (Polia)",                  "Polia",   "3 x 12-15"),
    ("Panturrilha", "Panturrilha Sentado",                    "Máquina", "4 x 15"),
]


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
A_FILL      = PatternFill("solid", fgColor="FFF2CC")
B_FILL      = PatternFill("solid", fgColor="DDEBF7")
THIN        = Side(border_style="thin", color="BFBFBF")
BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def fmt_header(ws, row, cols):
    for c, val in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=c, value=val)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_registro(wb):
    ws = wb.active
    ws.title = "Registro"

    headers = ["Data", "Treino", "Grupo", "Exercício", "Série", "Reps", "Carga (kg)", "RPE", "Observações"]
    fmt_header(ws, 1, headers)
    autosize(ws, [12, 8, 14, 38, 8, 8, 12, 6, 40])

    # Data validations: Treino A/B
    dv_treino = DataValidation(type="list", formula1='"A,B"', allow_blank=True)
    dv_treino.add("B2:B2000")
    ws.add_data_validation(dv_treino)

    # RPE 1-10
    dv_rpe = DataValidation(type="list", formula1='"6,7,8,9,10"', allow_blank=True)
    dv_rpe.add("H2:H2000")
    ws.add_data_validation(dv_rpe)

    # Format data column
    for r in range(2, 2001):
        ws.cell(row=r, column=1).number_format = "DD/MM/YYYY"
        ws.cell(row=r, column=7).number_format = "0.0"

    ws.freeze_panes = "A2"

    # Exemplo (linha 2) - mostra como preencher
    sample = ["2026-05-25", "A", "Costas", "Puxada Alta (Pulldown)", 1, 12, 50.0, 8, "Aquecimento"]
    for c, val in enumerate(sample, start=1):
        cell = ws.cell(row=2, column=c, value=val)
        cell.font = Font(italic=True, color="888888")
    ws.cell(row=2, column=1).number_format = "DD/MM/YYYY"


def build_treino_sheet(wb, title, exercicios, fill):
    ws = wb.create_sheet(title)
    headers = ["Grupo", "Exercício", "Tipo", "Séries x Reps"]
    fmt_header(ws, 1, headers)
    autosize(ws, [16, 40, 12, 18])

    for i, (grupo, ex, tipo, sxr) in enumerate(exercicios, start=2):
        ws.cell(row=i, column=1, value=grupo).fill = fill
        ws.cell(row=i, column=2, value=ex)
        ws.cell(row=i, column=3, value=tipo).alignment = Alignment(horizontal="center")
        ws.cell(row=i, column=4, value=sxr).alignment = Alignment(horizontal="center")
        for c in range(1, 5):
            ws.cell(row=i, column=c).border = BORDER

    ws.freeze_panes = "A2"


def build_resumo(wb):
    """Aba de resumo: melhor carga por exercício (com fórmulas que leem do Registro)."""
    ws = wb.create_sheet("Resumo")
    fmt_header(ws, 1, ["Exercício", "Carga Máx (kg)", "Última Carga (kg)", "Última Data"])
    autosize(ws, [40, 16, 18, 14])

    todos_ex = sorted({ex for _, ex, _, _ in TREINO_A + TREINO_B})
    for i, ex in enumerate(todos_ex, start=2):
        ws.cell(row=i, column=1, value=ex)
        # MÁXIMO de carga p/ esse exercício
        ws.cell(row=i, column=2,
                value=f'=IFERROR(MAXIFS(Registro!G:G,Registro!D:D,A{i}),"")')
        # Última carga registrada (procura a última linha desse exercício)
        ws.cell(row=i, column=3,
                value=f'=IFERROR(LOOKUP(2,1/(Registro!D:D=A{i}),Registro!G:G),"")')
        # Última data registrada
        last_date = f'=IFERROR(LOOKUP(2,1/(Registro!D:D=A{i}),Registro!A:A),"")'
        cell = ws.cell(row=i, column=4, value=last_date)
        cell.number_format = "DD/MM/YYYY"
        for c in range(1, 5):
            ws.cell(row=i, column=c).border = BORDER

    ws.freeze_panes = "A2"


def main():
    wb = Workbook()
    build_registro(wb)
    build_treino_sheet(wb, "Treino A", TREINO_A, A_FILL)
    build_treino_sheet(wb, "Treino B", TREINO_B, B_FILL)
    build_resumo(wb)
    wb.save(OUT)
    print(f"OK: {OUT}")


if __name__ == "__main__":
    main()
