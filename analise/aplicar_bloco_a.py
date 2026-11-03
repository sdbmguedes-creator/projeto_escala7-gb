"""Aplica o BLOCO A em uma copia nova. Nunca altera o arquivo original."""
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO.xlsx"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx"
HIST = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\HISTORICO_2025.xlsx"

ARCHIVE = ["FILTRO_DIAS", "Cópia de RESUMO", "Página30"]

# ---------- 1) arquivo historico (preserva as 3 abas e descarta o resto) ----------
wb = load_workbook(SRC)
for name in list(wb.sheetnames):
    if name not in ARCHIVE:
        del wb[name]
wb.save(HIST)
print("historico salvo:", wb.sheetnames)

# ---------- 2) arquivo de trabalho ----------
wb = load_workbook(SRC)
for name in ARCHIVE:
    del wb[name]

cleared = []
# P1 - limpar formulas da aba excluida
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            t = v.text if hasattr(v, "text") else v
            if isinstance(t, str) and "'VOLUNTÁRIOS PJES'!" in t:
                cleared.append(f"{ws.title}!{c.coordinate}")
                c.value = None

# P2 - remover a duplicidade do MONTEIRO (linha 82 da 1aSB, sem marcacao)
ws = wb["1ªSB"]
dup = []
for col in range(1, 41):
    c = ws.cell(82, col)
    if c.value is not None:
        dup.append(c.coordinate)
    c.value = None

# P4 - desocultar e documentar ORIENTAÇÕES
ori = wb["ORIENTAÇÕES"]
ori.sheet_state = "visible"
rows = [
    ("C", "CÓDIGOS E REGRAS — revisão de 01/11/2026", True),
    (None, None, None),
    ("C", "X", True),
    ("D", "Serviço ordinário de 24h, na escala fixa do militar (rodízio ALFA/BRAVO/CHARLIE/DELTA). Valor 1 na meta diária.", False),
    ("C", "PJ", True),
    ("D", "PJES de 24h. Valor 1 na meta diária.", False),
    ("C", "PJ12", True),
    ("D", "PJES de 12h. Valor 0,5 na meta diária.", False),
    ("C", "PJ12H", True),
    ("D", "Código de controle usado nas abas 2ªSB/3ªSB/4ªSB para o PJES de 12h de militar de OUTRA SB (visitante), "
          "para não entrar na contagem de serviço e de cotas da SB que o recebe. Não é erro de digitação.", False),
    ("C", "S1", True),
    ("D", "Apenas saldo de 1º turno. Valor 0,5 na meta diária.", False),
    ("C", "SPJ", True),
    ("D", "Saldo de 1º turno + PJES de 24h. Valor 1 na meta diária.", False),
    ("C", "N", True),
    ("D", "Não disponível / sem lançamento naquele dia. Não conta na meta.", False),
    ("C", "12 e 24 (só no ACO)", True),
    ("D", "Turno de 12h ou 24h do oficial de ACO. É hora de turno, não é PJES.", False),
    (None, None, None),
    ("C", "AFASTAMENTO — X - A, X - B, X - C, X - D, X - ADM", True),
    ("D", "Militar afastado no mês da escala. A letra indica a escala dele. Ele continua aparecendo na aba da sua "
          "seção, com N nos dias. Não o remova da aba.", False),
    (None, None, None),
    ("C", "MOT. (coluna de motorista)", True),
    ("D", "M = motorista de viatura operacional (entra na conta de motoristas do dia). "
          "B = motorista administrativo (não entra).", False),
    ("C", "FEM (coluna F)", True),
    ("D", "F = efetivo feminino, usado na conferência de limite de feminino por dia.", False),
    (None, None, None),
    ("C", "ORDEM DE LANÇAMENTO", True),
    ("D", "No RESUMO, lançar as praças por ANTIGUIDADE. Na aba da SB, manter a mesma ordem, agrupado pela escala "
          "(ALFA, BRAVO, CHARLIE, DELTA). Não inserir linha: usar a próxima linha vazia do bloco.", False),
    (None, None, None),
    ("C", "ROTINA DE ABERTURA DO MÊS", True),
    ("D", "1) duplicar as abas; 2) lançar o 1º dia do mês em ACO!A3; 3) conferir a tabela ESCALAS; "
          "4) limpar as marcações antigas; 5) reordenar o efetivo por antiguidade; 6) revisar as cotas PJES; "
          "7) rodar as conferências.", False),
]
r = 32
for col, text, bold in rows:
    if col is None:
        r += 1
        continue
    cell = ori[f"{col}{r}"]
    cell.value = text
    if bold:
        cell.font = Font(bold=True)
    r += 1

wb.save(OUT)
print("arquivo limpo salvo")
print("celulas limpas (VOLUNTARIOS PJES):", len(cleared))
print("linha 82 duplicada, celulas limpas:", len(dup), dup[:6], "...")
print("abas finais:", wb.sheetnames)
