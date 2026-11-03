"""Compara pessoas e secoes entre agosto, setembro, outubro e novembro."""
from openpyxl import load_workbook

ARQS = {
    "AGO": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\8 - Escala de Agosto 2026.xlsx",
    "SET": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\9 - Escala de Setembro 2026.xlsx",
    "OUT": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx",
    "NOV": r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\11 - Escala de Novembro 2026.xlsx",
}
MES = {}
for k, p in ARQS.items():
    wv = load_workbook(p, data_only=True)
    res = wv["RESUMO"]
    d = {}
    for r in range(2, 160):
        posto, mat, guerra = (res.cell(r, c).value for c in (2, 3, 4))
        sec = res.cell(r, 5).value
        esc = res.cell(r, 8).value
        if all(isinstance(x, str) for x in (posto, mat, guerra)) and isinstance(sec, str):
            d[f"{posto} {mat} {guerra}"] = (sec, esc)
    MES[k] = d
    print(f"{k}: {len(d)} pessoAS")

print()
ordem = ["AGO", "SET", "OUT", "NOV"]
todos = set().union(*(MES[m].keys() for m in ordem))
print("pessoas distintas nos 4 meses:", len(todos))
presentes = {m: sum(1 for n in todos if n in MES[m]) for m in ordem}
print("presentes por mes:", presentes)

saíram = [n for n in todos if n in MES["AGO"] and n not in MES["NOV"]]
entraram = [n for n in todos if n not in MES["AGO"] and n in MES["NOV"]]
print()
print("entraram depois de agosto:", len(entraram))
for n in list(entraram)[:6]:
    print("   ", n, "->", MES["NOV"][n])
print("saíram até novembro:", len(saíram))
for n in list(saíram)[:6]:
    print("   ", n, "->", MES["AGO"][n])

print()
print("=== quem mudou de SECAO entre meses consecutivos ===")
for a, b in [("AGO", "SET"), ("SET", "OUT"), ("OUT", "NOV")]:
    comuns = [n for n in MES[a] if n in MES[b]]
    mud = [(n, MES[a][n][0], MES[b][n][0]) for n in comuns if MES[a][n][0] != MES[b][n][0]]
    print(f"  {a}->{b}: {len(comuns)} em comum, {len(mud)} mudaram de secao")
    for m in mud[:6]:
        print("      ", m)

print()
print("=== quem mudou de ESCALA entre meses consecutivos ===")
for a, b in [("AGO", "SET"), ("SET", "OUT"), ("OUT", "NOV")]:
    comuns = [n for n in MES[a] if n in MES[b]]
    mud = [(n, MES[a][n][1], MES[b][n][1]) for n in comuns if MES[a][n][1] != MES[b][n][1]]
    print(f"  {a}->{b}: {len(mud)} mudaram de escala")
    for m in mud[:5]:
        print("      ", m)
