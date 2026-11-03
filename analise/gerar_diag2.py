"""Diagnostico do botao Distribuir proporcional."""
import io

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\app\index.html"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\app\diag2.html"

TESTES = """
<pre id="RES"></pre>
<script>
(function(){
  var R = [];
  window.alert = function(m){ R.push("ALERT: " + m); };
  window.confirm = function(m){ R.push("CONFIRM: " + m); return true; };
  try {
    var bs = document.querySelectorAll("#nav button");
    for (var i=0;i<bs.length;i++){ if (bs[i].textContent === "OPERACIONAL"){ bs[i].click(); break; } }
    R.push("cotas iniciais: " + JSON.stringify(D.params.cotas));

    var chamouSalvar = 0, chamouDesenhar = 0;
    var origSalvar = window.salvarLocal;
    var origDesenhar = window.desenhar;
    window.salvarLocal = function(){ chamouSalvar++; return origSalvar.apply(null, arguments); };
    window.desenhar = function(){ chamouDesenhar++; return origDesenhar.apply(null, arguments); };

    var si = document.getElementById("seitotal");
    R.push("seitotal existe: " + !!si);
    si.value = "238";
    R.push("valor no input: " + si.value);

    var bp = document.getElementById("btnprop");
    R.push("btnprop existe: " + !!bp);
    R.push("tipo do onclick: " + (typeof bp.onclick));

    var ev = new Event("change");
    si.dispatchEvent(ev);
    R.push("apos change do SEI, D.params.sei = " + D.params.sei);

    D.params.sei = 238;
    bp.click();
    R.push("chamouSalvar=" + chamouSalvar + " chamouDesenhar=" + chamouDesenhar);
    R.push("cotas apos clique: " + JSON.stringify(D.params.cotas));
  } catch(e) { R.push("ERRO | " + e.message); }
  document.getElementById("RES").textContent = R.join(" ||| ");
})();
</script>
"""

html = io.open(SRC, encoding="utf-8").read()
html = html.replace("</body>", TESTES + "</body>")
io.open(OUT, "w", encoding="utf-8").write(html)
print("diag2.html gerado")
