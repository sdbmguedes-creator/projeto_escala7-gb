from openpyxl import load_workbook
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P, data_only=True)
for sh, cells in {"4ªSB":["AM3","AL3","AN3"], "PJES DISTRIBUIÇÃO":["K5","K6","L6","B32","L39","B5","B6"],
                  "RESUMO":["P1","P2"], "ACO":["E5","E6"], "1ªSB":["E2","E3"]}.items():
    ws = wb[sh]
    for c in cells:
        print(f"{sh}!{c} = {ws[c].value!r}")
    print()
