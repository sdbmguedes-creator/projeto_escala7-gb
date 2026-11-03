"""Monta a base do 2.0: PESSOAS, MES (novembro) e AFASTAMENTOS."""
import re
from collections import Counter
from openpyxl import load_workbook
from openpyxl.styles import Font

ARQS = {
    "AGO": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\8 - Escala de Agosto 2026.xlsx",
    "SET": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\9 - Escala de Setembro 2026.xlsx",
    "OUT": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx",
    "NOV": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\11 - Escala de Novembro 2026.xlsx",
}
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_2.0_base.xlsx"
REAIS = {"ALFA", "BRAVO", "CHARLIE", "DELTA"}

# ---------- 1) PESSOAS: uniao dos 4 meses ----------
pessoas = {}
for k, p in ARQS.items():
    wv = load_workbook(p, data_only=True)
    res = wv["RESUMO"]
    for r in range(2, 160):
        posto, mat, guerra = (res.cell(r, c).value for c in (2, 3, 4))
        if not all(isinstance(x, str) for x in (posto, mat, guerra)):
            continue
        mat = mat.strip()
        pessoas[mat] = (posto.strip(), guerra.strip())

# ---------- 2) MES de novembro ----------
wv = load_workbook(ARQS["NOV"], data_only=True)
res = wv["RESUMO"]
mes = []
for r in range(2, 160):
    posto, mat, guerra = (res.cell(r, c).value for c in (2, 3, 4))
    sec = res.cell(r, 5).value
    mot = res.cell(r, 7).value
    esc = res.cell(r, 8).value
    if not all(isinstance(x, str) for x in (posto, mat, guerra)):
        continue
    admissao = "ADM" if (isinstance(esc, str) and "ADM" in esc) else ""
    mes.append((mat.strip(), sec, esc, admissao, mot))

# ---------- 3) AFASTAMENTOS: blocos de N na grade de novembro ----------
ABAS = {"1ªSB": ("G", "AK"), "2ªSB": ("F", "AJ"), "3ªSB": ("F", "AJ"), "4ªSB": ("F", "AJ")}
sec_de = {m[0]: m[1] for m in mes}
afast = []
for sh, (c1, c2) in ABAS.items():
    ws = wv[sh]
    for r in range(3, 120):
        nome = ws.cell(r, 4).value
        if not isinstance(nome, str) or not re.search(r"\d{6}-\d", nome):
            continue
        mat = re.search(r"\d{6}-\d", nome).group(0)
        dias_n = []
        for ci in range(ws[c1 + "1"].column, ws[c2 + "1"].column + 1):
            d = ws.cell(2, ci).value
            if d is None:
                continue
            if ws.cell(r, ci).value == "N":
                dias_n.append(d.day)
        if dias_n:
            afast.append((mat, nome, min(dias_n), max(dias_n), len(dias_n)))

print("pessoas distintas:", len(pessoas))
print("linhas no MES (nov):", len(mes))
print("blocos de N encontrados:", len(afast))
print()
print("=== afastamentos de novembro (N na grade) ===")
for a in sorted(afast, key=lambda x: -x[4])[:25]:
    print(f"   {a[1][:32]:34} dias {a[2]:2}-{a[3]:2} ({a[4]} dias)")

# ---------- 4) grava ----------
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = "PESSOAS"
ws.append(["ID", "POSTO", "NOME DE GUERRA"])
for c in ws[1]:
    c.font = Font(bold=True)
for mat, (posto, guerra) in sorted(pessoas.items()):
    ws.append([mat, posto, guerra])
ws.freeze_panes = "A2"

ws2 = wb.create_sheet("MES")
ws2.append(["ID", "SEÇÃO", "ESCALA", "ADM", "MOTORISTA"])
for c in ws2[1]:
    c.font = Font(bold=True)
for m in mes:
    ws2.append(list(m))
ws2.freeze_panes = "A2"

ws3 = wb.create_sheet("AFASTAMENTOS")
ws3.append(["ID", "NOME", "DE", "ATÉ", "DIAS"])
for c in ws3[1]:
    c.font = Font(bold=True)
for a in sorted(afast):
    ws3.append([a[0], a[1], a[2], a[3], a[4]])
ws3.freeze_panes = "A2"

for w in (ws, ws2, ws3):
    w.column_dimensions["A"].width = 14
    w.column_dimensions["B"].width = 30
    w.column_dimensions["C"].width = 22

wb.save(OUT)
print()
print("salvo:", OUT)
