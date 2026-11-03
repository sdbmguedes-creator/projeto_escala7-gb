from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wb = load_workbook(V2)

ABAS = {"1ªSB": ("G", "AK", 56, 92), "2ªSB": ("F", "AJ", 39, 53),
        "3ªSB": ("F", "AJ", 27, 41), "4ªSB": ("F", "AJ", 40, 55)}

for sh, (c1, c2, r1, r2) in ABAS.items():
    ws = wb[sh]
    n_form = n_livre = 0
    exemplo = None
    for r in range(r1, r2 + 1):
        if not isinstance(ws.cell(r, 4).value, str):
            continue
        for ci in range(ws[c1 + "1"].column, ws[c2 + "1"].column + 1):
            v = ws.cell(r, ci).value
            if isinstance(v, ArrayFormula):
                n_form += 1
                if exemplo is None:
                    exemplo = (ws.cell(r, ci).coordinate, v.text[:80])
            elif isinstance(v, str) and v.strip():
                n_livre += 1
    print(f"{sh} (linhas {r1}-{r2}): {n_form} celulas com formula | {n_livre} com texto digitado")
    if exemplo:
        print(f"    ex: {exemplo[0]} -> {exemplo[1]}")
    print()
