import re
from collections import Counter
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_corrigida_01-11.xlsx"

a = load_workbook(V2, data_only=True)
b = load_workbook(OUT)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


print("=== diff celula a celula (valores antes -> depois) ===")
n = 0
for sh in b.sheetnames:
    for row in b[sh].iter_rows():
        for c in row:
            antes = a[sh][c.coordinate].value
            depois = c.value
            depois_txt = t(depois)
            if isinstance(depois_txt, str) and depois_txt.startswith("="):
                continue
            if antes != depois:
                n += 1
                if n <= 20:
                    print(f"  {sh}!{c.coordinate}: {antes!r} -> {depois!r}")
print("total de celulas com valor alterado:", n)

print()
print("=== formula da linha 3 da 4aSB (deve estar intacta) ===")
print("  AL3:", str(t(b["4ªSB"]["AL3"].value))[:120])
print("  D3 :", str(t(b["4ªSB"]["D3"].value))[:120])

print()
print("=== varredura: codigos S1/SPJ em militar que NAO e ADM ===")
wv = load_workbook(OUT, data_only=True)
res = wv["RESUMO"]
secao = {}
for r in range(2, 112):
    nome, sec = res.cell(r, 6).value, res.cell(r, 5).value
    if isinstance(nome, str) and isinstance(sec, str):
        secao[nome] = sec
achados = []
for sh, rng in [("1ªSB", (7, 37)), ("2ªSB", (6, 36)), ("3ªSB", (6, 36)), ("4ªSB", (6, 36))]:
    ws = wv[sh]
    for r in range(3, 120):
        nome = ws.cell(r, 4).value
        if not isinstance(nome, str):
            continue
        codigos = {ws.cell(r, c).value for c in range(rng[0], rng[1] + 1)
                   if isinstance(ws.cell(r, c).value, str)}
        ruins = codigos & {"S1", "SPJ"}
        if ruins and secao.get(nome) != "ADM":
            achados.append((sh, r, nome, secao.get(nome), sorted(ruins)))
print("  ocorrencias:", len(achados))
for x in achados[:10]:
    print("   ", x)

print()
print("=== varredura: codigos invalidos ===")
VALIDOS = {"X", "PJ", "PJ12", "PJ12H", "S1", "SPJ", "N", "12", "24"}
for sh, rng in [("1ªSB", (7, 37)), ("2ªSB", (6, 36)), ("3ªSB", (6, 36)), ("4ªSB", (6, 36))]:
    ws = wv[sh]
    bad = Counter()
    for r in range(3, 120):
        for c in range(rng[0], rng[1] + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() and v.strip() not in VALIDOS:
                bad[v.strip()] += 1
    if bad:
        print(f"  {sh}: {dict(bad)}")
