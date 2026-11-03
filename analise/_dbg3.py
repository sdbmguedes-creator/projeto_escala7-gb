from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v2.xlsx"
wb = load_workbook(F)
ws = wb["2ªSB"]
for r in (39,40,53):
    c = ws.cell(r, 38)  # coluna AL = 38
    v = c.value
    t = v.text if isinstance(v, ArrayFormula) else v
    print(f"AL{r} coord={c.coordinate} tipo={type(v).__name__} repr={t!r}")
    if isinstance(t,str):
        s=t.strip()
        print("   startswith:", s.startswith("=COUNTIF("), "| endswith:", s.endswith('"PJ")*2'), "| len:", len(s))
