import re
from openpyxl import load_workbook

P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(P, data_only=True)
res = [wb["RESUMO"].cell(r, 6).value for r in range(2, 112)]
res = [x for x in res if isinstance(x, str)]
print("RESUMO nomes:", len(res), "| distintos:", len(set(res)))


def names(sh, a, b):
    ws = wb[sh]
    return [ws.cell(r, 4).value for r in range(a, b + 1)
            if isinstance(ws.cell(r, 4).value, str) and re.search(r"\d{6}-\d", ws.cell(r, 4).value)]


for sh, (a, b) in {"1ªSB": (3, 92), "2ªSB": (3, 53), "3ªSB": (3, 41), "4ªSB": (3, 55)}.items():
    ns = names(sh, a, b)
    faltando = [n for n in ns if n not in set(res)]
    print(f"{sh}: {len(ns)} nomes | nao encontrados no RESUMO: {len(faltando)}")
    for n in faltando[:6]:
        print("    ", repr(n))

com_espaco = [n for n in res if n != n.strip()]
print("RESUMO nomes com espaco sobrando:", [repr(x) for x in com_espaco])
