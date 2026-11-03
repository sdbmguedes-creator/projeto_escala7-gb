from openpyxl import load_workbook

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wv = load_workbook(V2, data_only=True)

for sh, trecho, c1, c2 in [("1ªSB", "CELERINO", "G", "AK"), ("4ªSB", "CLEITON", "F", "AJ")]:
    ws = wv[sh]
    for r in range(3, 40):
        nome = ws.cell(r, 4).value
        if isinstance(nome, str) and trecho in nome:
            partes = []
            for ci in range(ws[c1 + "1"].column, ws[c2 + "1"].column + 1):
                d = ws.cell(2, ci).value
                dia = d.day if d is not None else "?"
                v = ws.cell(r, ci).value
                partes.append(str(dia).rjust(2) + ":" + (v if v else "."))
            print(sh, r, nome, "| escala:", ws.cell(r, 2).value)
            print("   " + "  ".join(partes))
            print()
            break
