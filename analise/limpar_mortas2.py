"""Remove funcoes mortas respeitando strings, templates e comentarios."""
import io, re

P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\analise\fase1_app.py"
ALVOS = ["desenharGrade", "desenharOperacional", "desenharParametros", "desenharPjes"]

s = io.open(P, encoding="utf-8").read()
orig = len(s)


def fim_do_corpo(txt, i):
    """i aponta para '{'; devolve indice do '}' correspondente."""
    prof = 0
    n = len(txt)
    while i < n:
        c = txt[i]
        if c in "'\"`":
            q = c
            i += 1
            while i < n:
                if txt[i] == "\\":
                    i += 2
                    continue
                if txt[i] == q:
                    i += 1
                    break
                i += 1
            continue
        if c == "/" and i + 1 < n and txt[i + 1] == "/":
            j = txt.find("\n", i)
            i = n if j < 0 else j
            continue
        if c == "/" and i + 1 < n and txt[i + 1] == "*":
            j = txt.find("*/", i + 2)
            i = n if j < 0 else j + 2
            continue
        if c == "{":
            prof += 1
        elif c == "}":
            prof -= 1
            if prof == 0:
                return i
        i += 1
    return -1


removidos = []
for nome in ALVOS:
    m = re.search(r"\nfunction " + nome + r"\s*\(", s)
    if not m:
        print("  nao encontrada:", nome)
        continue
    ini = m.start() + 1
    j = s.index("{", m.end())
    k = fim_do_corpo(s, j)
    if k < 0:
        print("  corpo nao fechado:", nome)
        continue
    fim = k + 1
    while fim < len(s) and s[fim] in "\r\n":
        fim += 1
    removidos.append((ini, fim, nome))

# remover de tras para frente
for ini, fim, nome in sorted(removidos, key=lambda t: -t[0]):
    print(f"  removida {nome}: {fim-ini} caracteres")
    s = s[:ini] + s[fim:]

io.open(P, "w", encoding="utf-8", newline="\n").write(s)
print(f"template: {orig} -> {len(s)} (removidos {orig-len(s)})")
