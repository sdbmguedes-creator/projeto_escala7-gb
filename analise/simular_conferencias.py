"""Simula as conferencias sobre os dados reais (arquivo com valores em cache)."""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

V2 = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11_V2.xlsx"
wv = load_workbook(V2, data_only=True)

escalas = wv["ESCALAS 2026"]
meses = [escalas.cell(r, 1).value for r in range(6, 18)]
tabela = [[escalas.cell(r, c).value for c in range(2, 33)] for r in range(6, 18)]
MES_PT = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
          "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

ABAS = {"1ªSB": ("G", "AK", 3, 30), "2ªSB": ("F", "AJ", 3, 32),
        "3ªSB": ("F", "AJ", 3, 20), "4ªSB": ("F", "AJ", 3, 33)}

for sh, (c1, c2, r1, r2) in ABAS.items():
    ws = wv[sh]
    i_mes = i_dia = None
    p10 = p11 = 0
    exemplos10, exemplos11 = [], []
    for r in range(r1, r2 + 1):
        nome = ws.cell(r, 4).value
        if not isinstance(nome, str):
            continue
        escala_pessoa = ws.cell(r, 2).value
        for ci in range(ws[f"{c1}1"].column, ws[f"{c2}1"].column + 1):
            data = ws.cell(2, ci).value
            if data is None:
                continue
            mes = MES_PT[data.month - 1]
            dia = data.day
            if mes in meses:
                esperada = tabela[meses.index(mes)][dia - 1]
            else:
                esperada = None
            marca = ws.cell(r, ci).value
            if marca == "X" and esperada and escala_pessoa and esperada != escala_pessoa:
                p10 += 1
                if len(exemplos10) < 5:
                    exemplos10.append(f"{nome[:22]} dia {dia}: escala {escala_pessoa} x dia {esperada}")
            viz = ws.cell(r, ci + 1).value
            if marca in ("X", "PJ") and viz in ("X", "PJ"):
                p11 += 1
                if len(exemplos11) < 5:
                    exemplos11.append(f"{nome[:22]} dia {dia} e {dia + 1}")
    print(f"=== {sh} ===")
    print(f"  P10 (X fora da escala): {p10} celulas")
    for e in exemplos10:
        print("     ", e)
    print(f"  P11 (24h em dias seguidos): {p11} celulas")
    for e in exemplos11:
        print("     ", e)
