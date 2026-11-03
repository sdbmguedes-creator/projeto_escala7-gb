// Testes simulaveis do app (rodar com node <arquivo>)
process.env.TZ = process.env.TZ_TEST || "America/Sao_Paulo";

function fmtLocal(d){
  const p = n => String(n).padStart(2,"0");
  return d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate());
}
function fmtISO(d){ return d.toISOString().slice(0,10); }

const nov1 = new Date(2026, 10, 1, 12, 0, 0);
const nov30 = new Date(2026, 10, 30, 12, 0, 0);
console.log("TZ =", process.env.TZ);
console.log("  dia 1  -> local:", fmtLocal(nov1), "| ISO:", fmtISO(nov1));
console.log("  dia 30 -> local:", fmtLocal(nov30), "| ISO:", fmtISO(nov30));

// distribuicao com dezembro (31 dias)
function distribuir(nec, necMot, rec, SBS){
  const alvo = {}, base = {};
  let soma = 0;
  SBS.forEach(sb => { alvo[sb] = Math.max(nec[sb], necMot[sb]); soma += alvo[sb]; });
  const resto = [];
  if (rec >= soma){
    SBS.forEach(sb => { base[sb] = alvo[sb]; });
    const sobraTotal = rec - soma;
    SBS.forEach(sb => {
      const bruto = sobraTotal * alvo[sb] / soma;
      base[sb] += Math.floor(bruto);
      resto.push({ sb, r: bruto - Math.floor(bruto) });
    });
  } else {
    SBS.forEach(sb => {
      const bruto = rec * alvo[sb] / soma;
      base[sb] = Math.floor(bruto);
      resto.push({ sb, r: bruto - Math.floor(bruto) });
    });
  }
  let sobra = rec - SBS.reduce((a, sb) => a + base[sb], 0);
  resto.sort((a, b) => b.r - a.r);
  let i = 0;
  while (sobra > 0 && i < resto.length){ base[resto[i].sb]++; sobra--; i++; }
  return { alvo, soma, base, total: SBS.reduce((a, sb) => a + base[sb], 0) };
}

const SBS = ["1ªSB", "2ªSB", "3ªSB", "4ªSB"];

// novembro (30 dias): efetivo 25/26/17/28, metas 8/7/4/7
const D = 30;
const efetN = { "1ªSB": 25, "2ªSB": 26, "3ªSB": 17, "4ªSB": 28 };
const metaN = { "1ªSB": 8, "2ªSB": 7, "3ªSB": 4, "4ªSB": 7 };
const necN = {};
SBS.forEach(sb => { necN[sb] = Math.ceil(2 * Math.max(0, metaN[sb] * D - efetN[sb] * D / 4)); });
const necMotN = { "1ªSB": 0, "2ªSB": 0, "3ªSB": 15, "4ªSB": 15 };
console.log("\nNOVEMBRO (30 dias)");
console.log("  Necessidade:", necN);
const r238 = distribuir(necN, necMotN, 238, SBS);
console.log("  SEI 238 -> alvos", r238.alvo, "soma", r238.soma, "| distribuido", r238.base, "total", r238.total);

// dezembro (31 dias), metas e efetivo iguais
const D2 = 31;
const necD = {};
SBS.forEach(sb => { necD[sb] = Math.ceil(2 * Math.max(0, metaN[sb] * D2 - efetN[sb] * D2 / 4)); });
console.log("\nDEZEMBRO (31 dias)");
console.log("  Necessidade:", necD);
const r300 = distribuir(necD, necMotN, 300, SBS);
console.log("  SEI 300 -> alvos", r300.alvo, "soma", r300.soma);
console.log("  distribuido", r300.base, "| total", r300.total);
const inteiros = SBS.every(sb => Number.isInteger(r300.base[sb]));
console.log("  todos inteiros:", inteiros, "| soma == 300:", r300.total === 300);
