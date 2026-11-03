from openpyxl import load_workbook
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P)
for sh, rng in [("2ªSB","F3:AJ30"),("3ªSB","F3:AJ41"),("4ªSB","F3:AJ31"),("1ªSB","G3:AK92")]:
    ws = wb[sh]; c = {}
    for row in ws[rng]:
        for cell in row:
            v = cell.value
            if isinstance(v,str) and v.strip() and not v.startswith("="):
                c[v.strip()] = c.get(v.strip(),0)+1
    print(sh, dict(sorted(c.items(), key=lambda x:-x[1])[:12]))
