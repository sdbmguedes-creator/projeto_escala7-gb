from openpyxl import load_workbook

P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_afastamento_01-11.xlsx"
wb = load_workbook(P)
ws = wb["3ªSB"]
ws["O2"] = "='2ªSB'!O2"
wb.save(P)
print("3ªSB!O2 restaurado:", ws["O2"].value)

wb2 = load_workbook(P)
for sh in ["1ªSB", "2ªSB", "3ªSB", "4ªSB"]:
    w = wb2[sh]
    print(f"  {sh}: N2={w['N2'].value!r} | O2={w['O2'].value!r} | P2={w['P2'].value!r}")

res = wb2["RESUMO"]
print()
print("  RESUMO AE1/AF1:", res["AE1"].value, "/", res["AF1"].value)
print("  janelas preenchidas:", sum(1 for r in range(2, 112) if res.cell(r, 31).value is not None))
