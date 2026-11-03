from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wbv = load_workbook(P, data_only=True)
for sh in ["1ªSB","ACO"]:
    ws = wbv[sh]; hits=[]
    for row in ws.iter_rows():
        for c in row:
            if c.value == "#N/A": hits.append(c.coordinate)
    print(sh, len(hits), hits[:16])
ws = wbv["1ªSB"]
rows=[r for r in range(1, ws.max_row+1) if ws.cell(r,4).value=="SD 722022-7 MONTEIRO"]
print("MONTEIRO em 1aSB linhas:", rows)
# linhas vazias com formula na coluna B/C
empty_calc=[r for r in range(3,93) if not ws.cell(r,4).value and ws.cell(r,2).value]
print("linhas sem nome mas com formula em B:", empty_calc)
