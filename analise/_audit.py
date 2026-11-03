"""Quantitative audit: broken refs, range drift, duplicate names, hardcoded ranges."""
import re
from collections import Counter, defaultdict
from openpyxl import load_workbook

PATH = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
wb = load_workbook(PATH, data_only=False)
wbv = load_workbook(PATH, data_only=True)
out = []
def p(*a):
    out.append(" ".join(str(x) for x in a))

sheets = set(wb.sheetnames)

# 1) erros em valores cacheados
p("== 1. ERROS EM CELULAS (valores cacheados) ==")
err_counter = Counter()
err_by_sheet = defaultdict(Counter)
for ws in wb.worksheets:
    wsv = wbv[ws.title]
    for row in wsv.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("#"):
                err_counter[v] += 1
                err_by_sheet[ws.title][v] += 1
p("total:", dict(err_counter))
for sh, cnt in err_by_sheet.items():
    p(f"  {sh}: {dict(cnt)}")

# 2) referencias a abas inexistentes
p("\n== 2. ABAS REFERENCIADAS QUE NAO EXISTEM ==")
refs = Counter()
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "!" in c.value:
                for m in re.findall(r"'([^']+)'!", c.value):
                    if m not in sheets:
                        refs[m] += 1
                for m in re.findall(r"(?<![A-Za-zÀ-ÿ'])([A-Za-zÀ-ÿ0-9_]+)!", c.value):
                    if m not in sheets and not m.startswith("_xlfn"):
                        refs[m] += 1
p(dict(refs))

# 3) faixas de VLOOKUP para RESUMO (drift)
p("\n== 3. VLOOKUP EM RESUMO: faixas usadas (fim de intervalo) ==")
rng = Counter()
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.upper().startswith("=VLOOKUP"):
                m = re.search(r"RESUMO!?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", v, re.I)
                if m:
                    rng[(ws.title, f"{m.group(1)}{m.group(2)}:{m.group(3)}{m.group(4)}")] += 1
for (sh, r), n in sorted(rng.items()):
    p(f"  {sh}: {r}  x{n}")

# 4) formulas de matriz
p("\n== 4. FORMULAS DE MATRIZ (ArrayFormula) ==")
for ws in wb.worksheets:
    n = sum(1 for row in ws.iter_rows() for c in row if type(c.value).__name__ == "ArrayFormula")
    if n:
        p(f"  {ws.title}: {n}")

# 5) validacao de dados: listas
p("\n== 5. VALIDACOES (lista) e suas faixas ==")
for ws in wb.worksheets:
    for d in ws.data_validations.dataValidation:
        p(f"  {ws.title}: {d.sqref} -> {str(d.formula1)[:110]}")

# 6) nomes duplicados nas abas das SBs
p("\n== 6. NOMES DUPLICADOS POR ABA ==")
for name in ["1ªSB", "2ªSB", "3ªSB", "4ªSB", "ACO", "RESUMO"]:
    ws = wb[name]
    col = "D" if name != "RESUMO" else "F"
    vals = [c.value for row in ws.iter_rows(min_col=ws[col + "1"].column, max_col=ws[col + "1"].column)
            for c in row if isinstance(c.value, str) and re.search(r"\d{6}-\d", c.value)]
    dup = {k: v for k, v in Counter(vals).items() if v > 1}
    p(f"  {name} ({col}): {len(vals)} nomes, duplicados: {dup if dup else 'nenhum'}")

# 7) hardcodes de meta e de faixas de linha nas formulas de totais
p("\n== 7. CONSTANTES/FAIXAS FIXAS nas formulas de total ==")
for sh in ["1ªSB", "2ªSB", "3ªSB", "4ªSB"]:
    ws = wb[sh]
    hard = Counter()
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("="):
                for m in re.findall(r"SUM\((\d+)-", v):
                    hard[f"meta {m}"] += 1
                for m in re.findall(r"\$(\d+):\$?[A-Z]*(\d+)", v):
                    hard[f"faixa {m[0]}:{m[1]}"] += 1
    p(f"  {sh}: {dict(hard.most_common(14))}")

# 8) contagem de militares por bloco em cada SB
p("\n== 8. BLOCOS DE MILITARES POR SB ==")
for sh, blocks in [("1ªSB", [(3, 29), (39, 49), (55, 92)]), ("2ªSB", [(3, 30), (39, 53)]),
                   ("3ªSB", [(3, 18), (27, 41)]), ("4ªSB", [(3, 31), (40, 55)])]:
    ws = wb[sh]
    info = []
    for a, b in blocks:
        n = sum(1 for r in range(a, b + 1) if isinstance(ws.cell(row=r, column=4).value, str)
                and re.search(r"\d{6}-\d", str(ws.cell(row=r, column=4).value)))
        info.append(f"{a}-{b}:{n}")
    p(f"  {sh}: " + " | ".join(info))

print("\n".join(out))
