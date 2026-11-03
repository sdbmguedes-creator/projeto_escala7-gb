import math

SBS = ["1ªSB", "2ªSB", "3ªSB", "4ªSB"]
nec = {"1ªSB": 105, "2ªSB": 30, "3ªSB": 0, "4ªSB": 0}
necMot = {"1ªSB": 0, "2ªSB": 0, "3ªSB": 15, "4ªSB": 15}


def distribuir(rec):
    alvo = {sb: max(nec[sb], necMot[sb]) for sb in SBS}
    soma = sum(alvo.values())
    base = {}
    resto = []
    if rec >= soma:
        for sb in SBS:
            base[sb] = alvo[sb]
        sobra_total = rec - soma
        for sb in SBS:
            bruto = sobra_total * alvo[sb] / soma
            base[sb] += math.floor(bruto)
            resto.append((sb, bruto - math.floor(bruto)))
        sobra = rec - sum(base.values())
    else:
        for sb in SBS:
            bruto = rec * alvo[sb] / soma
            base[sb] = math.floor(bruto)
            resto.append((sb, bruto - math.floor(bruto)))
        sobra = rec - sum(base.values())
    resto.sort(key=lambda x: -x[1])
    i = 0
    while sobra > 0 and i < len(resto):
        base[resto[i][0]] += 1
        sobra -= 1
        i += 1
    return alvo, soma, base


for rec in (238, 165, 100):
    alvo, soma, base = distribuir(rec)
    print(f"SEI={rec} | alvos={alvo} somaAlvos={soma} | distribuido={base} "
          f"total={sum(base.values())} | recebido-alvos={rec-soma}")
