from collections import Counter
from openpyxl import load_workbook
V3 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v3.xlsx"
wb = load_workbook(V3); res = wb["RESUMO"]
c = Counter()
for r in range(2,112):
    v = res.cell(r,16).value
    if v is not None: c[str(v)] += 1
print("conteudo da coluna P:", dict(c))
