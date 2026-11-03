"""P5b (busca linear no RESUMO!I) + normalizacao das 11 faixas da 2aSB."""
import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v4.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v5.xlsx"

wb = load_workbook(F)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


def put(cell, novo):
    v = cell.value
    cell.value = ArrayFormula(ref=v.ref, text=novo) if isinstance(v, ArrayFormula) else novo


# ---- P5b: busca binaria -> linear nas formulas do RESUMO!I ----
res = wb["RESUMO"]
n_bin = 0
for r in range(2, 112):
    c = res[f"I{r}"]
    v = t(c.value)
    if isinstance(v, str) and ",0,0,1)" in v:
        put(c, v.replace(",0,0,1)", ",0,0,0)"))
        n_bin += 1

# ---- faixas deslocadas da 2aSB ----
ws = wb["2ªSB"]
alvos = [("AK", 17), ("AK", 19), ("AK", 24), ("AK", 30),
         ("AL", 5), ("AL", 14), ("AL", 17), ("AL", 19),
         ("AL", 24), ("AL", 26), ("AL", 30)]
pat = re.compile(r'COUNTIF\(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+,"')
n_faixa = 0
for col, r in alvos:
    c = ws[f"{col}{r}"]
    v = t(c.value)
    if not isinstance(v, str):
        continue
    novo = pat.sub(f'COUNTIF(F{r}:AJ{r},"', v)
    if novo != v:
        put(c, novo)
        n_faixa += 1

wb.save(OUT)
print("P5b - formulas convertidas para busca linear:", n_bin)
print("faixas normalizadas na 2aSB:", n_faixa)
print()
for col, r in alvos[:4]:
    print(f"  2ªSB!{col}{r}:", t(wb["2ªSB"][f"{col}{r}"].value)[:150])
