import re, io, json, math

h = io.open(r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\app\index.html", encoding="utf-8").read()
m = re.search(r"const D = (\{.*?\});\n", h, re.S)
D = json.loads(m.group(1))

EQUIPES = ["ALFA", "BRAVO", "CHARLIE", "DELTA"]
metaMot = D["params"]["metaMot"]
ndias = 30

mes = D["mes"]
print("Nec. mot. por SB (meta, equipes, deficit, cotas):")
for sb in ["1ªSB", "2ªSB", "3ªSB", "4ªSB"]:
    pessoas = [x for x in mes if x["secao"] == sb and x.get("adm") != "S"]
    Mm = metaMot.get(sb, 0) or 0
    det = {}
    deficit = 0
    for e in EQUIPES:
        n = len([x for x in pessoas if str(x["escala"]).strip().upper() == e
                 and str(x.get("motorista")).strip().upper() == "M"])
        det[e] = n
        deficit += max(0, Mm - n) * (ndias / 4)
    cotas = math.ceil(deficit * 2)
    print(f"  {sb}: meta={Mm} | " + ", ".join(f"{e} {det[e]}" for e in EQUIPES)
          + f" | deficit={deficit} | NEC MOT = {cotas}")
