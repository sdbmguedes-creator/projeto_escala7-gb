from openpyxl import load_workbook
V = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v5.xlsx"
wb = load_workbook(V)
for sh in ["1ªSB","2ªSB","3ªSB","4ªSB"]:
    ws = wb[sh]
    print(f"===== {sh} (max_row={ws.max_row}) =====")
    for r in range(1, 60):
        d = ws.cell(r,4).value
        rotulo = ws.cell(r,4).value if isinstance(d,str) and not d.startswith("=") and len(str(d))<40 else None
        tem_nome = isinstance(d,str) and "SGT" in d or (isinstance(d,str) and d[:2] in ("CB","SD","ST","1º","2º","3º"))
        # formula da linha de totais?
        g = ws.cell(r,7).value
        eh_total = isinstance(g,str) and g.startswith("=SUMPRODUCT")
        marca = []
        if tem_nome: marca.append("MILITAR")
        if eh_total: marca.append("TOTAL")
        if rotulo: marca.append(f"'{rotulo}'")
        if marca: print(f"  {r:3}: " + " | ".join(marca))
