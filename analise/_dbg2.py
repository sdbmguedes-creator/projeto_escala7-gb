import re
s = '=COUNTIF(F39:AJ39,"PJ")*2'
p = re.compile(r'^(=COUNTIF\((\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+),"PJ"\)\*2)$')
print("match:", bool(p.match(s)))
if p.match(s): print(p.match(s).group(1), "|", p.match(s).group(2))
