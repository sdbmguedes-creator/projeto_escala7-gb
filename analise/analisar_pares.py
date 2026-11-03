"""Analisa os pares de dias consecutivos, por tipo de codigo."""
from collections import Counter
from openpyxl import load_workbook

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wv = load_workbook(V2, data_only=True)

ABAS = {"1ªSB": ("G", "AK", 3, 30), "2ªSB": ("F", "AJ", 3, 32),
        "3ªSB": ("F", "AJ", 3, 20), "4ªSB": ("F", "AJ", 3, 33)}

pares = Counter()
detalhe = []
for sh, (c1, c2, r1, r2) in ABAS.items():
    ws = wv[sh]
    for r in range(r1, r2 + 1):
        nome = ws.cell(r, 4).value
        if not isinstance(nome, str):
            continue
        for ci in range(ws[f"{c1}1"].column, ws[f"{c2}1"].column):
            a = ws.cell(r, ci).value
            b = ws.cell(r, ci + 1).value
            if a in ("X", "PJ", "PJ12", "PJ12H", "SPJ", "S1") and b in ("X", "PJ", "PJ12", "PJ12H", "SPJ", "S1"):
                pares[(a, b)] += 1
                d = ws.cell(2, ci).value
                detalhe.append((sh, r, nome, d.day if d else "?", a, b))

print("=== pares consecutivos encontrados ===")
for (a, b), n in pares.most_common():
    print(f"  {a:6} -> {b:6} : {n}")

print()
print("=== detalhe: X seguido de PJES (proibido pela sua regra) ===")
proibidos = [d for d in detalhe if d[4] == "X" and d[5].startswith("PJ")]
print("  total:", len(proibidos))
for d in proibidos:
    print(f"   {d[0]} linha {d[1]} {str(d[2])[:26]:28} dias {d[3]}-{d[3]+1}  ({d[4]} -> {d[5]})")

print()
print("=== detalhe: PJES seguido de X (permitido pela sua regra) ===")
permitidos = [d for d in detalhe if d[4].startswith("PJ") and d[5] == "X"]
print("  total:", len(permitidos))
for d in permitidos[:6]:
    print(f"   {d[0]} linha {d[1]} {str(d[2])[:26]:28} dias {d[3]}-{d[3]+1}")

print()
print("=== detalhe: X seguido de X (dois ordinarios em dias seguidos) ===")
xx = [d for d in detalhe if d[4] == "X" and d[5] == "X"]
print("  total:", len(xx))
for d in xx[:6]:
    print(f"   {d[0]} linha {d[1]} {str(d[2])[:26]:28} dias {d[3]}-{d[3]+1}")
