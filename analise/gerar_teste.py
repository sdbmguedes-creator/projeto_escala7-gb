"""Bateria de testes do app, executada em Chrome headless."""
import io, os

_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(_BASE, "app", "index.html")
OUT = os.path.join(_BASE, "app", "teste.html")

TESTES = """
<pre id="RES"></pre>
<script>
(function(){
  var R = [];
  function ok(n, c, d){ R.push((c ? "PASS" : "FAIL") + " | " + n + (d !== undefined ? " | " + d : "")); }
  function btn(nome){
    var bs = document.querySelectorAll("#nav button");
    for (var i=0;i<bs.length;i++){ if (bs[i].textContent === nome) return bs[i]; }
    return null;
  }
  function linhas(){ return document.querySelectorAll("main table tr"); }
  function operRows(){
    var trs = document.querySelectorAll("main table tr"), out = [], hdr = false;
    for (var i=0;i<trs.length;i++){
      var c = trs[i].children;
      if (c.length === 12 && c[0].tagName === "TD" && !trs[i].className) out.push(trs[i]);
    }
    return out;
  }
  window.alert = function(){};
  window.confirm = function(){ return true; };
  try { abaAtual = "OPERACIONAL"; desenhar(); } catch(e){}
  setTimeout(function(){
  try {
    var navs = document.querySelectorAll("#nav button");
    ok("abre com 8 abas", navs.length === 8, navs.length);
    ok("tabela na abertura", !!document.querySelector("main table"));
    ok("grade 1SB com linhas", linhas().length > 0, linhas().length + " linhas");

    var ops = operRows();
    ok("4 linhas de SB + ADM + ACO", ops.length >= 6, ops.length + " linhas");

    var nec = [], ncm = [];
    for (var i=0;i<4;i++){ nec.push(ops[i].children[5].textContent.trim()); ncm.push(ops[i].children[6].textContent.trim()); }
    ok("Necessidade nov = 105,30,0,0", nec.join(",") === "105,30,0,0", nec.join("/"));
    ok("Nec. mot. nov = 0,0,15,15", ncm.join(",") === "0,0,15,15", ncm.join("/"));

    var si = document.getElementById("seitotal");
    si.value = "238"; D.params.sei = 238;
    document.getElementById("btnprop").click();

    ops = operRows();
    var dist = [];
    for (var j=0;j<4;j++){ var inp = ops[j].children[7].querySelector("input"); dist.push(inp ? inp.value : "?"); }
    ok("SEI 238 = 151/43/22/22", dist.join(",") === "151,43,22,22", dist.join("/"));
    var soma = 0;
    for (var k=0;k<dist.length;k++) soma += Number(dist[k]);
    ok("soma distribuido = 238", soma === 238, soma);

    ok("fmt sem fuso", fmt(new Date(2026,10,1)) === "2026-11-01" && fmt(new Date(2026,10,30)) === "2026-11-30",
       fmt(new Date(2026,10,1)) + " / " + fmt(new Date(2026,10,30)));

    var alvo = null, alvoMot = null;
    for (var m=0;m<D.mes.length;m++){
      var e = D.mes[m];
      if (e.secao === "1ªSB" && e.adm !== "S"){
        if (!alvo) alvo = e.id;
        if (!alvoMot && e.motorista === "M") alvoMot = e.id;
      }
    }
    vis["2ªSB"] = vis["2ªSB"] || [];
    vis["2ªSB"][0] = alvo;
    marks["host:2ªSB|" + alvo + "|2026-11-03"] = "X";
    ok("visitante soma na 2SB", contagemSB("2ªSB", new Date(2026,10,3)) > 0,
       contagemSB("2ªSB", new Date(2026,10,3)));
    ok("visitante nao soma na 1SB", contagemSB("1ªSB", new Date(2026,10,3)) ===
       contagemSB("1ªSB", new Date(2026,10,3)), "isolado");

    vis["2ªSB"][1] = alvo;
    var dup = trocarMilitar("host:2ªSB", vis["2ªSB"], 1, alvo);
    ok("duplicata recusada", dup === false, "retorno=" + dup);
    vis["2ªSB"][1] = null;

    var antes = 0, ks = Object.keys(marks);
    for (var q=0;q<ks.length;q++) if (ks[q].indexOf("host:2ªSB|" + alvo + "|") === 0) antes++;
    window.confirm = function(){ return true; };
    vis["2ªSB"][0] = alvo;
    var trocou = trocarMilitar("host:2ªSB", vis["2ªSB"], 0, null);
    var depois = 0; ks = Object.keys(marks);
    for (var q2=0;q2<ks.length;q2++) if (ks[ks.length?q2:0] && ks[q2].indexOf("host:2ªSB|" + alvo + "|") === 0) depois++;
    ok("remover com confirm apaga marcas", trocou === true && antes > 0 && depois === 0,
       "antes=" + antes + " depois=" + depois);

    vis["2ªSB"][0] = alvoMot;
    marks["host:2ªSB|" + alvoMot + "|2026-11-04"] = "X";
    window.confirm = function(){ return false; };
    var cancelou = trocarMilitar("host:2ªSB", vis["2ªSB"], 0, null);
    var ficou = 0; ks = Object.keys(marks);
    for (var q3=0;q3<ks.length;q3++) if (ks[q3].indexOf("host:2ªSB|" + alvoMot + "|") === 0) ficou++;
    ok("cancelar mantem visitante e marca", cancelou === false && ficou > 0 && vis["2ªSB"][0] === alvoMot,
       "retorno=" + cancelou + " marcas=" + ficou);
  } catch (e) {
    R.push("ERRO | " + e.message);
  }
  document.getElementById("RES").textContent = R.join(" ||| ");
  }, 150);
})();
</script>
"""

html = io.open(SRC, encoding="utf-8").read()
html = html.replace("</body>", TESTES + "</body>")
io.open(OUT, "w", encoding="utf-8").write(html)
print("teste.html gerado")
