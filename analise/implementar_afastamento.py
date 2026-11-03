"""Implementa o afastamento parcial (AE/AF) + reconferencias, num arquivo so."""
import re
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_corrigida_01-11.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\ESCALA_afastamento_01-11.xlsx"

PADRAO = re.compile(r"\d{6}-\d")
AJUDA = 100
VERM = PatternFill(start_color="FFE06666", end_color="FFE06666", fill_type="solid")
AMAR = PatternFill(start_color="FFFFD966", end_color="FFFFD966", fill_type="solid")
AZUL = PatternFill(start_color="FF9FC5E8", end_color="FF9FC5E8", fill_type="solid")

wb = load_workbook(SRC)
log = []

# ---------- 1) normalizar espacos ----------
res = wb["RESUMO"]
for r in range(2, 115):
    for col in (2, 3, 4):
        c = res.cell(r, col)
        if isinstance(c.value, str):
            novo = " ".join(c.value.split())
            if novo != c.value:
                c.value = novo
                log.append(f"espaco: RESUMO!{c.coordinate}")
for sh in ["1ªSB", "2ªSB", "3ªSB", "4ªSB"]:
    ws = wb[sh]
    for r in range(3, 130):
        c = ws.cell(r, 4)
        if isinstance(c.value, str) and PADRAO.search(c.value):
            novo = " ".join(c.value.split())
            if novo != c.value:
                c.value = novo
                log.append(f"espaco: {sh}!{c.coordinate}")
for tab, col in [("ACO", 4), ("Planilha de PJES", 3)]:
    ws = wb[tab]
    for r in range(2, 260):
        c = ws.cell(r, col)
        if isinstance(c.value, str) and PADRAO.search(c.value):
            novo = " ".join(c.value.split())
            if novo != c.value:
                c.value = novo
                log.append(f"espaco: {tab}!{c.coordinate}")

# ---------- 2) colunas AE/AF ----------
res["AE1"] = "AFASTADO DE (dia)"
res["AF1"] = "AFASTADO ATÉ (dia)"
res["AE1"].font = Font(bold=True)
res["AF1"].font = Font(bold=True)

JANELAS = {"3º SGT 710281-0 CELERINO": (1, 15), "3º SGT 711134-7 CLEITON": (16, 30)}
for r in range(2, 112):
    posto, mat, guerra = (res.cell(r, c).value for c in (2, 3, 4))
    if not all(isinstance(x, str) for x in (posto, mat, guerra)):
        continue
    nome = f"{posto} {mat} {guerra}"
    escala = res.cell(r, 8).value
    if nome in JANELAS:
        res.cell(r, 31).value, res.cell(r, 32).value = JANELAS[nome]
        log.append(f"janela parcial: {nome} = {JANELAS[nome]}")
    elif isinstance(escala, str) and escala.startswith("X - "):
        res.cell(r, 31).value, res.cell(r, 32).value = 1, 31
        log.append(f"janela mes inteiro: {nome}")

# ---------- 3) cabecalho do dia 10 da 4aSB, se estiver vazio ----------
ws4 = wb["4ªSB"]
if ws4["O2"].value is None:
    ws4["O2"].value = "=ACO!O5"
    log.append("4ªSB!O2 restaurado (cabecalho do dia 10 estava vazio)")

# ---------- 4) conferencias ----------
ABAS = {"1ªSB": ("G", "AK", 3, 30), "2ªSB": ("F", "AJ", 3, 32),
        "3ªSB": ("F", "AJ", 3, 20), "4ªSB": ("F", "AJ", 3, 33)}

for sh, (c1, c2, r1, r2) in ABAS.items():
    ws = wb[sh]
    ini, fim = ws[f"{c1}1"].column, ws[f"{c2}1"].column

    ws.cell(AJUDA, 4).value = "CONFERENCIA - escala do dia (linha oculta, nao apagar)"
    for i in range(fim - ini + 1):
        col = get_column_letter(ini + i)
        ws[f"{col}{AJUDA}"].value = (
            "=IFERROR(INDEX('ESCALAS 2026'!$B$6:$AF$17,"
            f'MATCH(UPPER(TEXT({col}$2,"mmmm")),\'ESCALAS 2026\'!$A$6:$A$17,0),DAY({col}$2)),"")'
        )
    ws.row_dimensions[AJUDA].hidden = True

    rng = f"{c1}{r1}:{c2}{r2}"
    # P10 - X fora da escala (ignora dias dentro da janela de afastamento)
    ws.conditional_formatting.add(rng, FormulaRule(formula=[
        f'AND({c1}{r1}="X",{c1}${AJUDA}<>"",{c1}${AJUDA}<>$B{r1},'
        f'LEFT($B{r1},3)<>"X -",LEFT($B{r1},3)<>"ADM",'
        f'NOT(AND($AE{r1}<>"",DAY({c1}$2)>=$AE{r1},DAY({c1}$2)<=$AF{r1})))'], fill=VERM))
    # P10b - servico marcado dentro da janela de afastamento
    ws.conditional_formatting.add(rng, FormulaRule(formula=[
        f'AND($AE{r1}<>"",{c1}{r1}<>"",{c1}{r1}<>"N",'
        f'DAY({c1}$2)>=$AE{r1},DAY({c1}$2)<=$AF{r1})'], fill=AZUL))
    # P11 - PJES seguido de ordinario
    viz = get_column_letter(ini + 1)
    ws.conditional_formatting.add(rng, FormulaRule(formula=[
        f'AND(OR({c1}{r1}="PJ",{c1}{r1}="PJ12",{c1}{r1}="PJ12H"),{viz}{r1}="X")'], fill=AMAR))

# P12 - cadastro
res.conditional_formatting.add("F2:F400", FormulaRule(
    formula=['AND($F2<>"",COUNTIF($F$2:$F$400,$F2)>1)'], fill=VERM))
res.conditional_formatting.add("F2:F400", FormulaRule(
    formula=['AND($F2<>"",TRIM($F2)<>$F2)'], fill=AMAR))

wb.save(OUT)
print("salvo:", OUT)
print("alteracoes registradas:", len(log))
for item in log[:20]:
    print("  ", item)
