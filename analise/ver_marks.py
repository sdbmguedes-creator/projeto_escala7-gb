from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11.xlsx"
wb = load_workbook(F)
wv = load_workbook(F, data_only=True)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


def marcas(sh, r, c1, c2):
    wsv = wv[sh]
    out = []
    for c in range(c1, c2 + 1):
        v = wsv.cell(r, c).value
        if isinstance(v, str) and v.strip():
            out.append(f"{wsv.cell(r, c).coordinate}={v}")
    return out


print("=== 1aSB: FERNANDA (linha 3 = vaga | linha 41 = bloco ADM) ===")
print("  linha 3 :", marcas("1ªSB", 3, 7, 37))
print("  linha 41:", marcas("1ªSB", 41, 7, 37))
print("  AM3 =", wv["1ªSB"].cell(3, 39).value, "| AM41 =", wv["1ªSB"].cell(41, 39).value)
print()
print("=== 4aSB: ADERBAL (linha 3 = vaga | linha 13 = antiga) ===")
print("  linha 3 :", marcas("4ªSB", 3, 6, 36))
print("  linha 13:", marcas("4ªSB", 13, 6, 36))
print("  AL3 =", wv["4ªSB"].cell(3, 38).value, "| AL13 =", wv["4ªSB"].cell(13, 38).value)
print()
print("=== RESUMO: valor de P para os dois ===")
res, resv = wb["RESUMO"], wv["RESUMO"]
for r in range(2, 112):
    nome = resv.cell(r, 6).value
    if isinstance(nome, str) and ("FERNANDA" in nome or "ADERBAL" in nome):
        print(f"  linha {r}: {nome} | P={resv.cell(r,16).value!r} | I={resv.cell(r,9).value}")
