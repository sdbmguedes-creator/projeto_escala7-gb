from openpyxl import load_workbook

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_revertida_01-11.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_final_v2_01-11.xlsx"

wb = load_workbook(SRC)
ws = wb["1ªSB"]

alvo = ["J73", "V73", "AD73", "AH73"]
for c in alvo:
    antes = ws[c].value
    ws[c].value = "SB"
    print(f"  1ªSB!{c}: {antes!r} -> 'SB'")

wb.save(OUT)
print("salvo:", OUT)
