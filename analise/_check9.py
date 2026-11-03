from openpyxl import load_workbook
P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P, data_only=True)
res = wb["RESUMO"]; sec = {}
for r in range(2,112):
    n = res.cell(r,6).value
    if isinstance(n,str): sec[n] = res.cell(r,5).value
por_secao = {}
for sh,(a,b,c1,c2) in {"1ªSB":(3,92,7,37),"2ªSB":(3,53,6,36),"3ªSB":(3,41,6,36),"4ªSB":(3,55,6,36)}.items():
    ws = wb[sh]
    for r in range(a,b+1):
        nome = ws.cell(r,4).value
        if not isinstance(nome,str): continue
        codes = {ws.cell(r,c).value for c in range(c1,c2+1) if isinstance(ws.cell(r,c).value,str)}
        for cod in ("S1","SPJ"):
            if cod in codes:
                por_secao.setdefault(cod, set()).add((sec.get(nome), nome))
for cod, vals in por_secao.items():
    secoes = sorted({v[0] for v in vals})
    print(f"{cod}: {len(vals)} militares | secoes envolvidas: {secoes}")
    print(f"   ex: {[v[1] for v in list(vals)[:4]]}")
