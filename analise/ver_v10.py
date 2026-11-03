from openpyxl import load_workbook

wb = load_workbook(r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_2.0_prototipo.xlsx")
g = wb["GRADE 1ªSB"]
f = g["G65"].value
print("G65 (BM):")
print(" ", f)
print()
print("blocos somados:", f.count("COUNTIF(G$"))
print("DIARIO B3:", wb["DIARIO"]["B3"].value)
print("PJES C4:", wb["PJES"]["C4"].value)
print("PJES C8 (ADM):", wb["PJES"]["C8"].value)
