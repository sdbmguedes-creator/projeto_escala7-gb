"""Varredura de anomalias no arquivo COM valores em cache (V2)."""
from collections import Counter
from openpyxl import load_workbook

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wv = load_workbook(V2, data_only=True)
res = wv["RESUMO"]

secao = {}
for r in range(2, 112):
    nome, sec = res.cell(r, 6).value, res.cell(r, 5).value
    if isinstance(nome, str) and isinstance(sec, str):
        secao[nome] = sec

VALIDOS = {"X", "PJ", "PJ12", "PJ12H", "S1", "SPJ", "N", "12", "24"}
RANGES = [("1ªSB", 7, 37), ("2ªSB", 6, 36), ("3ªSB", 6, 36), ("4ªSB", 6, 36)]

print("=== S1/SPJ em quem NAO e ADM ===")
achados = []
for sh, c1, c2 in RANGES:
    ws = wv[sh]
    for r in range(3, 120):
        nome = ws.cell(r, 4).value
        if not isinstance(nome, str):
            continue
        cods = {ws.cell(r, c).value for c in range(c1, c2 + 1) if isinstance(ws.cell(r, c).value, str)}
        ruins = cods & {"S1", "SPJ"}
        if ruins and secao.get(nome) != "ADM":
            achados.append((sh, r, nome, secao.get(nome), sorted(ruins)))
print("  total:", len(achados))
for x in achados:
    print("   ", x)

print()
print("=== codigos fora da lista válida (só em linhas com militar) ===")
for sh, c1, c2 in RANGES:
    ws = wv[sh]
    bad = Counter()
    for r in range(3, 120):
        if not isinstance(ws.cell(r, 4).value, str):
            continue
        for c in range(c1, c2 + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() and v.strip() not in VALIDOS:
                bad[v.strip()] += 1
    if bad:
        print(f"  {sh}: {dict(bad)}")
