"""Diagnostico: mostra cabecalhos e contagens reais das tabelas."""
import io

SRC = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\app\index.html"
OUT = r"C:\Users\Guedes\Downloads\TESTES_DEEPSEEK\app\diag.html"

TESTES = """
<pre id="RES"></pre>
<script>
(function(){
  var R = [];
  try {
    R.push("main tables: " + document.querySelectorAll("main table").length);
    var t0 = document.querySelector("main table");
    R.push("table0 rows total: " + t0.querySelectorAll("tr").length);
    R.push("table0 tr (qualquer): " + t0.getElementsByTagName("tr").length);
    R.push("tbody tr: " + t0.querySelectorAll("tbody tr").length);
    var primeira = t0.getElementsByTagName("tr")[2];
    R.push("linha2 celulas: " + (primeira ? primeira.children.length : -1));

    var bs = document.querySelectorAll("#nav button");
    for (var i=0;i<bs.length;i++){
      if (bs[i].textContent === "OPERACIONAL") { bs[i].click(); break; }
    }
    var t1 = document.querySelector("main table");
    var ths = t1.querySelectorAll("thead th");
    var cab = [];
    for (var j=0;j<ths.length;j++) cab.push(ths[j].textContent.trim());
    R.push("CABECALHO: " + cab.join(" | "));

    var trs = t1.querySelectorAll("tbody tr");
    R.push("linhas oper: " + trs.length);
    if (trs.length){
      var cells = [];
      for (var k=0;k<trs[0].children.length;k++) cells.push(k + "=" + trs[0].children[k].textContent.trim());
      R.push("LINHA 1: " + cells.join(" ; "));
    }
  } catch(e) { R.push("ERRO | " + e.message); }
  document.getElementById("RES").textContent = R.join(" ||| ");
})();
</script>
"""

html = io.open(SRC, encoding="utf-8").read()
html = html.replace("</body>", TESTES + "</body>")
io.open(OUT, "w", encoding="utf-8").write(html)
print("diag.html gerado")
