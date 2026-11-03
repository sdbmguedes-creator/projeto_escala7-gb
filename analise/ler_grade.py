"""Le a grade de CELERINO e CLEITON para deduzir a janela de afastamento."""
from openpyxl import load_workbook

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wv = load_workbook(V2, data_only=True)

ALVOS = [("1ªSB", "CELERINO", "G", "AK"), ("4ªSB", "CLEITON", "F", "AJ")]

for sh, trecho, c1, c2 in ALVOS:
    ws = wv[sh]
    for r in range(3, 40):
        nome = ws.cell(r, 4).value
        if isinstance(nome, str) and trecho in nome:
            print(f"=== {sh} linha {r}: {nome} | escala={ws.cell(r,2).value!r} ===")
            linha = []
            for ci in range(ws[f"{c1}1"].column, ws[f"{c2}1"].column + 1):
                d = ws.cell(2, ci).value
                v = ws.cell(r, ci).value
                linha.append(f"{d.day:02d}:{v if v else '.'}")
            print("   " + "  ".join(linha))
            break
