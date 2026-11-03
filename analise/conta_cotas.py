"""Recalcula S9 por casamento exato de nome, a partir do export."""
from openpyxl import load_workbook

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_export_01-11.xlsx"
wv = load_workbook(F, data_only=True)

res = wv["RESUMO"]
pessoas = []
for r in range(2, 112):
    nome = res.cell(r, 6).value
    sec = res.cell(r, 5).value
    if isinstance(nome, str) and isinstance(sec, str):
        pessoas.append((r, nome, sec))


def valor(tab, col_valor, a, b, alvo, col_nome=4):
    ws = wv[tab]
    for r in range(a, b + 1):
        if ws.cell(r, col_nome).value == alvo:
            v = ws.cell(r, col_valor).value
            return v if isinstance(v, (int, float)) else 0
    return 0


por_secao = {}
detalhe = []
for r, nome, sec in pessoas:
    j = valor("ACO", 37, 6, 54, nome)
    k = valor("1ªSB", 39, 3, 120, nome)
    l = valor("2ªSB", 38, 3, 120, nome)
    m = valor("3ªSB", 38, 3, 120, nome)
    n = valor("4ªSB", 38, 3, 120, nome)
    total = j + k + l + m + n
    por_secao[sec] = por_secao.get(sec, 0) + total
    if total:
        detalhe.append((nome, sec, j, k, l, m, n, total))

print("por secao (calculado):", por_secao)
print("TOTAL calculado:", sum(por_secao.values()))
print()
print("valores na planilha: ADM 66 | 1aSB 69 | 2aSB 71 | 3aSB 51 | 4aSB 75 | total 332")
print()
print("=== 4aSB: quem tem cota ===")
for d in detalhe:
    if d[1] == "4ªSB":
        print(f"  {d[0][:30]:32} ACO={d[2]} 1aSB={d[3]} 2aSB={d[4]} 3aSB={d[5]} 4aSB={d[6]} -> {d[7]}")
