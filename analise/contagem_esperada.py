"""Contagem esperada dos destaques, com as regras refinadas."""
from openpyxl import load_workbook

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wv = load_workbook(V2, data_only=True)

escalas = wv["ESCALAS 2026"]
meses = [escalas.cell(r, 1).value for r in range(6, 18)]
tabela = [[escalas.cell(r, c).value for c in range(2, 33)] for r in range(6, 18)]
MES_PT = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
          "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

# correcao ja aplicada no arquivo: marcacoes novas da 4aSB linha 3
CORRIGIDO = {"4ªSB": {3: {"I": "X", "M": "X", "O": "PJ", "Q": "X",
                          "U": "X", "Y": "X", "AC": "X", "AG": "X"}}}

ABAS = {"1ªSB": ("G", "AK", 3, 30), "2ªSB": ("F", "AJ", 3, 32),
        "3ªSB": ("F", "AJ", 3, 20), "4ªSB": ("F", "AJ", 3, 33)}

total = {"P10": 0, "P10b": 0, "P11": 0}
for sh, (c1, c2, r1, r2) in ABAS.items():
    ws = wv[sh]
    p10 = p10b = p11 = 0
    ex10, ex10b, ex11 = [], [], []
    for r in range(r1, r2 + 1):
        nome = ws.cell(r, 4).value
        if not isinstance(nome, str):
            continue
        esc = ws.cell(r, 2).value
        afastado = isinstance(esc, str) and esc.startswith("X -")
        linha_corr = CORRIGIDO.get(sh, {}).get(r, {})
        for ci in range(ws[f"{c1}1"].column, ws[f"{c2}1"].column + 1):
            data = ws.cell(2, ci).value
            if data is None:
                continue
            letra = ws.cell(r, ci).column_letter
            marca = linha_corr.get(letra, ws.cell(r, ci).value)
            esperada = tabela[meses.index(MES_PT[data.month - 1])][data.day - 1] if MES_PT[data.month - 1] in meses else None
            if marca == "X" and esperada and not afastado and esperada != esc:
                p10 += 1
                ex10.append(f"{nome[:20]} dia {data.day}: {esc} x dia {esperada}")
            if afastado and marca not in (None, "", "N"):
                p10b += 1
                ex10b.append(f"{nome[:20]} dia {data.day}: {marca}")
            prox = ws.cell(r, ci + 1).value if ci + 1 <= ws[f"{c2}1"].column else None
            if marca in ("X", "PJ") and prox in ("X", "PJ"):
                p11 += 1
                ex11.append(f"{nome[:20]} dias {data.day}-{data.day + 1}")
    total["P10"] += p10
    total["P10b"] += p10b
    total["P11"] += p11
    print(f"=== {sh} ===")
    print(f"  P10  (X fora da escala)......: {p10}")
    for e in ex10[:4]:
        print("       ", e)
    print(f"  P10b (afastado com marcacao).: {p10b}")
    for e in ex10b[:4]:
        print("       ", e)
    print(f"  P11  (24h em dias seguidos)..: {p11}")
    for e in ex11[:4]:
        print("       ", e)
print()
print("TOTAL esperado:", total)
