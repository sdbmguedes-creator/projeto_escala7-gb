"""Adiciona as conferencias P10, P11 e P12 ao arquivo corrigido."""
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_corrigida_01-11.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_conferencias_01-11.xlsx"

AJUDA = 100
FILL_ERRO = PatternFill(start_color="FFE06666", end_color="FFE06666", fill_type="solid")
FILL_ALERTA = PatternFill(start_color="FFFFD966", end_color="FFFFD966", fill_type="solid")

ABAS = {"1ªSB": ("G", "AK", 3, 30), "2ªSB": ("F", "AJ", 3, 32),
        "3ªSB": ("F", "AJ", 3, 20), "4ªSB": ("F", "AJ", 3, 33)}

wb = load_workbook(SRC)
relatorio = []

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

    f10 = f'AND({c1}{r1}="X",{c1}${AJUDA}<>"",{c1}${AJUDA}<>$B{r1})'
    ws.conditional_formatting.add(f"{c1}{r1}:{c2}{r2}", FormulaRule(formula=[f10], fill=FILL_ERRO))

    viz = get_column_letter(col_ini + 1)
    f11 = f'AND(OR({c1}{r1}="X",{c1}{r1}="PJ"),OR({viz}{r1}="X",{viz}{r1}="PJ"))'
    ws.conditional_formatting.add(f"{c1}{r1}:{c2}{r2}", FormulaRule(formula=[f11], fill=FILL_ALERTA))

    relatorio.append((sh, f"{c1}{r1}:{c2}{r2}", f10, f11, ncols))

res = wb["RESUMO"]
res.conditional_formatting.add("F2:F400",
    FormulaRule(formula=['AND($F2<>"",COUNTIF($F$2:$F$400,$F2)>1)'], fill=FILL_ERRO))
res.conditional_formatting.add("F2:F400",
    FormulaRule(formula=['AND($F2<>"",TRIM($F2)<>$F2)'], fill=FILL_ALERTA))

wb.save(OUT)
print("salvo:", OUT)
for sh, rng, f10, f11, n in relatorio:
    print(f"  {sh}: {rng} | {n} colunas auxiliares")
    print(f"     P10: {f10}")
    print(f"     P11: {f11}")
