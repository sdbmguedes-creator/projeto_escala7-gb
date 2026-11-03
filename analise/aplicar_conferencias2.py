"""Conferencias P10/P10b/P11/P12 - versao refinada, aplicada no arquivo."""
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_corrigida_01-11.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_conferencias_01-11.xlsx"

AJUDA = 100
VERM = PatternFill(start_color="FFE06666", end_color="FFE06666", fill_type="solid")
AMAR = PatternFill(start_color="FFFFD966", end_color="FFFFD966", fill_type="solid")
AZUL = PatternFill(start_color="FF9FC5E8", end_color="FF9FC5E8", fill_type="solid")

ABAS = {"1ªSB": ("G", "AK", 3, 30), "2ªSB": ("F", "AJ", 3, 32),
        "3ªSB": ("F", "AJ", 3, 20), "4ªSB": ("F", "AJ", 3, 33)}

wb = load_workbook(SRC)
saida = []

for sh, (c1, c2, r1, r2) in ABAS.items():
    ws = wb[sh]
    col_ini = ws[f"{c1}1"].column
    col_fim = ws[f"{c2}1"].column
    ncols = col_fim - col_ini + 1

    ws.cell(AJUDA, 4).value = "CONFERENCIA - escala do dia (linha oculta, nao apagar)"
    for i in range(ncols):
        col = get_column_letter(col_ini + i)
        ws[f"{col}{AJUDA}"].value = (
            "=IFERROR(INDEX('ESCALAS 2026'!$B$6:$AF$17,"
            f'MATCH(UPPER(TEXT({col}$2,"mmmm")),\'ESCALAS 2026\'!$A$6:$A$17,0),DAY({col}$2)),"")'
        )
    ws.row_dimensions[AJUDA].hidden = True

    rng = f"{c1}{r1}:{c2}{r2}"
    # P10 — X em dia que não é o da escala (só para quem tem escala fixa)
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND({c1}{r1}="X",{c1}${AJUDA}<>"",{c1}${AJUDA}<>$B{r1},LEFT($B{r1},3)<>"X -",LEFT($B{r1},3)<>"ADM")'],
        fill=VERM))
    # P10b — afastado (X - ***) com marcação de serviço
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(LEFT($B{r1},3)="X -",{c1}{r1}<>"",{c1}{r1}<>"N")'],
        fill=AZUL))
    # P11 — PJES seguido de ordinário (proibido: faltam 12h de descanso)
    viz = get_column_letter(col_ini + 1)
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(OR({c1}{r1}="PJ",{c1}{r1}="PJ12",{c1}{r1}="PJ12H"),{viz}{r1}="X")'],
        fill=AMAR))
    saida.append(sh)

res = wb["RESUMO"]
res.conditional_formatting.add("F2:F400", FormulaRule(
    formula=['AND($F2<>"",COUNTIF($F$2:$F$400,$F2)>1)'], fill=VERM))
res.conditional_formatting.add("F2:F400", FormulaRule(
    formula=['AND($F2<>"",TRIM($F2)<>$F2)'], fill=AMAR))

wb.save(OUT)
print("salvo:", OUT)
print("abas com conferencia:", saida)
