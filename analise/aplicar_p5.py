"""P5: padroniza todas as faixas de cadastro para RESUMO!$F$2:$H$400."""
import re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

F = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v3.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_v4.xlsx"

wb = load_workbook(F)


def t(v):
    return v.text if isinstance(v, ArrayFormula) else v


def put(cell, novo):
    v = cell.value
    cell.value = ArrayFormula(ref=v.ref, text=novo) if isinstance(v, ArrayFormula) else novo


# faixas do tipo RESUMO!F2:H189 (com ou sem $, inicio 2 fixo ou nao)
PAT_RESUMO = re.compile(r"RESUMO!?\$?F\$?\d*:\$?H\$?\d*", re.IGNORECASE)
# literais de INDIRECT nas formulas do RESUMO
PAT_IND = [(re.compile(r"'!\$D3:\$D32"), "'!$D$3:$D$120"),
           (re.compile(r"'!\$AL3:\$AL32"), "'!$AL$3:$AL$120"),
           (re.compile(r"'!\$AK3:\$AK32"), "'!$AK$3:$AK$120")]

n_resumo = n_ind = 0
abas_tocadas = {}
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            atual = t(c.value)
            if not (isinstance(atual, str) and atual.startswith("=")):
                continue
            novo = atual
            novo, k1 = PAT_RESUMO.subn("RESUMO!$F$2:$H$400", novo)
            k2 = 0
            for pat, rep in PAT_IND:
                novo, k = pat.subn(rep, novo)
                k2 += k
            if novo != atual:
                put(c, novo)
                n_resumo += k1
                n_ind += k2
                abas_tocadas[ws.title] = abas_tocadas.get(ws.title, 0) + 1

wb.save(OUT)
print("ocorrencias RESUMO!F:H normalizadas:", n_resumo)
print("literais de INDIRECT normalizados :", n_ind)
print("celulas alteradas por aba:", abas_tocadas)
print("total celulas:", sum(abas_tocadas.values()))
