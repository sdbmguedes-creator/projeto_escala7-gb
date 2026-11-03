from openpyxl import load_workbook

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_final_01-11.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_revertida_01-11.xlsx"

wb = load_workbook(SRC)
res = wb["RESUMO"]

print("=== conferindo se as colunas AE/AF existem neste arquivo ===")
print("  AE1:", res["AE1"].value, "| AF1:", res["AF1"].value)
preenchidas = sum(1 for r in range(2, 115) if res.cell(r, 31).value is not None)
print("  celulas preenchidas em AE:", preenchidas)

# corrige apenas o cabecalho do dia 10 (defeito independente do afastamento)
ws3 = wb["3ªSB"]
if ws3["O2"].value is None:
    ws3["O2"] = "='2ªSB'!O2"
    print("  3ªSB!O2 restaurado")

wb.save(OUT)

wb2 = load_workbook(OUT)
res2 = wb2["RESUMO"]
print()
print("=== arquivo revertido ===")
print("  AE1:", res2["AE1"].value, "| AF1:", res2["AF1"].value)
print("  AE preenchidas:", sum(1 for r in range(2, 115) if res2.cell(r, 31).value is not None))
for sh in ["1ªSB", "2ªSB", "3ªSB", "4ªSB", "RESUMO"]:
    n = sum(len(v) for v in wb2[sh].conditional_formatting._cf_rules.values())
    print(f"  {sh}: {n} regras de formatacao condicional")
