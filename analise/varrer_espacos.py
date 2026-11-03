"""Varre espacos indevidos em strings de nome em todas as abas."""
import re
from openpyxl import load_workbook

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_conferencias_01-11.xlsx"
wb = load_workbook(F)

PADRAO = re.compile(r"\d{6}-\d")
total = 0
for ws in wb.worksheets:
    achados = []
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if not isinstance(v, str) or not PADRAO.search(v):
                continue
            normal = " ".join(v.split())
            if normal != v:
                achados.append((c.coordinate, repr(v)))
    if achados:
        total += len(achados)
        print(f"=== {ws.title}: {len(achados)} ===")
        for a in achados[:12]:
            print("   ", a)
print()
print("TOTAL de celulas de nome com espaco irregular:", total)
