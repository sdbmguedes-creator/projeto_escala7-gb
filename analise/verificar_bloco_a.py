"""Verificação do Bloco A: diff celula a celula contra o original."""
from openpyxl import load_workbook

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx"

a = load_workbook(SRC)
b = load_workbook(OUT)

print("abas removidas:", [s for s in a.sheetnames if s not in b.sheetnames])
print("abas novas   :", [s for s in b.sheetnames if s not in a.sheetnames])
print()

changed = {}
for name in b.sheetnames:
    wa, wb_ = a[name], b[name]
    diffs = []
    maxr = max(wa.max_row, wb_.max_row)
    maxc = max(wa.max_column, wb_.max_column)
    for r in range(1, maxr + 1):
        for c in range(1, maxc + 1):
            va = wa.cell(r, c).value
            vb = wb_.cell(r, c).value
            ta = va.text if hasattr(va, "text") else va
            tb = vb.text if hasattr(vb, "text") else vb
            if ta != tb:
                diffs.append(wa.cell(r, c).coordinate)
    if diffs:
        changed[name] = diffs
        print(f"{name}: {len(diffs)} celulas alteradas -> {diffs[:8]}{' ...' if len(diffs) > 8 else ''}")
print("\nTOTAL celulas alteradas:", sum(len(v) for v in changed.values()))
print("(esperado: 240 limpezas + 38 da linha 82 + linhas novas em ORIENTAÇÕES)")

# erros remanescentes
bv = load_workbook(OUT, data_only=True)
err = {}
for ws in bv.worksheets:
    n = sum(1 for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("#"))
    if n:
        err[ws.title] = n
print("\n#ERROS remanescentes:", err, "| total:", sum(err.values()))

# referencias residuais a aba excluida
res = []
for ws in b.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            t = v.text if hasattr(v, "text") else v
            if isinstance(t, str) and "VOLUNTÁRIOS PJES" in t:
                res.append(f"{ws.title}!{c.coordinate}")
print("referencias residuais a 'VOLUNTÁRIOS PJES':", res)

# checagem de integridade: totais-chave precisam continuar iguais
print("\n=== totais-chave (original -> novo) ===")
av = load_workbook(SRC, data_only=True)
for sh, cell in [("RESUMO", "S9"), ("RESUMO", "S17"), ("RESUMO", "T19"),
                 ("1ªSB", "AM30"), ("1ªSB", "AM50"), ("1ªSB", "AM51"),
                 ("2ªSB", "AK31"), ("4ªSB", "AL34")]:
    print(f"  {sh}!{cell}: {av[sh][cell].value} -> {bv[sh][cell].value}")
