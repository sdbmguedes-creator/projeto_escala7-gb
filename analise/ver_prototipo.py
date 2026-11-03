from openpyxl import load_workbook

wb = load_workbook(r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_2.0_prototipo.xlsx")
print("abas:", wb.sheetnames)
for sh in ["GRADE 1ªSB", "GRADE 2ªSB", "GRADE 3ªSB", "GRADE 4ªSB"]:
    ws = wb[sh]
    print(f"  {sh}")
    print(f"     AN3 = {str(ws['AN3'].value)[:70]}")
    print(f"     D3  = {str(ws['D3'].value)[:80]}")
print()
print("MES H2 (posicao):", wb["MES"]["H2"].value)
print("MES I2 (chave):  ", wb["MES"]["I2"].value)
print("PARAMETROS D:", [wb["PARAMETROS"].cell(r, 4).value for r in range(4, 8)])
