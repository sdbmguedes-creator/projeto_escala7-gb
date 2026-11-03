"""Remove funcoes mortas do template JS, casando chaves."""
import io, re, sys

P = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\analise\fase1_app.py"
ALVOS = ["desenharGrade", "desenharOperacional", "desenharParametros", "desenharPjes"]

s = io.open(P, encoding="utf-8").read()
antes = len(s)

for nome in ALVOS:
    pad = re.compile(r"\nfunction " + nome + r"\(")
    m = pad.search(s)
    if not m:
        print("  nao encontrada:", nome)
        continue
    i = m.start() + 1              # inicio da funcao
    j = s.index("{", s.index("(", m.end()))  # primeira chave do corpo
    prof = 0
    k = j
    while k < len(s):
        if s[k] == "{":
            prof += 1
        elif s[k] == "}":
            prof -= 1
            if prof == 0:
                break
        k += 1
    fim = k + 1
    while fim < len(s) and s[fim] in "\r\n":
        fim += 1
    removido = fim - i
    s = s[:i] + s[fim:]
    print(f"  removida {nome}: {removido} caracteres")

io.open(P, "w", encoding="utf-8", newline="\n").write(s)
print(f"template: {antes} -> {len(s)} caracteres (removidos {antes-len(s)})")
