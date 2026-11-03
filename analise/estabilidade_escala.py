"""Verifica se a escala acompanha a pessoa entre meses."""
from openpyxl import load_workbook

ARQS = {
    "AGO": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\8 - Escala de Agosto 2026.xlsx",
    "SET": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\9 - Escala de Setembro 2026.xlsx",
    "OUT": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx",
    "NOV": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\11 - Escala de Novembro 2026.xlsx",
}
REAIS = {"ALFA", "BRAVO", "CHARLIE", "DELTA"}

dados = {}
for k, p in ARQS.items():
    wv = load_workbook(p, data_only=True)
    res = wv["RESUMO"]
    d = {}
    for r in range(2, 160):
        posto, mat, guerra = (res.cell(r, c).value for c in (2, 3, 4))
        if not all(isinstance(x, str) for x in (posto, mat, guerra)):
            continue
        d[f"{posto} {mat} {guerra}"] = (res.cell(r, 5).value, res.cell(r, 8).value)
    dados[k] = d

comuns = [n for n in dados["AGO"] if all(n in dados[m] for m in ARQS)]
print("pessoas presentes nos 4 meses:", len(comuns))
print()

estavel = instavel = so_afast = 0
exemplos = []
for n in comuns:
    escs = [dados[m][n][1] for m in ARQS]
    reais = [e for e in escs if e in REAIS]
    if len(set(reais)) <= 1:
        estavel += 1
    else:
        instavel += 1
        exemplos.append((n, escs))
print(f"mantiveram a mesma escala real (quando disponiveis): {estavel}")
print(f"mudaram de escala real entre meses: {instavel}")
print()
for e in exemplos[:10]:
    print("   ", e[0][:32], e[1])

print()
print("=== a escala acompanha a PESSOA quando ela troca de SB? ===")
for n in comuns[:60]:
    secs = [dados[m][n][0] for m in ARQS]
    escs = [dados[m][n][1] for m in ARQS]
    if len(set(secs)) > 1:
        print(f"   {n[:30]:32} secoes={secs} escalas={escs}")
