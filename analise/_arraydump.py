"""Print the real text of array formulas (openpyxl hides them behind objects)."""
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

PATH = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(PATH, data_only=False)

targets = [("RESUMO", ["A2", "J2", "K2", "L2", "M2", "N2", "A112", "B112"]),
           ("1ªSB", ["A3", "AN3", "A39", "AN39", "A55", "AL55"]),
           ("2ªSB", ["A3", "B3", "F3", "AK3"]),
           ("ACO", ["A6", "F6", "A57", "AK57"]),
           ("CONF. INDIV.", ["A2", "C8"]),
           ("Planilha de PJES", ["D2", "K2"]),
           ("Base de Dados", ["AD4"]),
           ("FILTRO_DIAS", ["A1", "A2"])]
for sheet, cells in targets:
    ws = wb[sheet]
    print(f"--- {sheet} ---")
    for c in cells:
        v = ws[c].value
        if isinstance(v, ArrayFormula):
            print(f"  {c}: ref={v.ref} text={v.text[:420]}")
        elif isinstance(v, str) and v.startswith("="):
            print(f"  {c}: {v[:420]}")
        else:
            print(f"  {c}: ({type(v).__name__}) {str(v)[:120]}")
