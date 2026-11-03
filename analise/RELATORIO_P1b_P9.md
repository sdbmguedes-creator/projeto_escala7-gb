# RELATÓRIO — P1b + P9 (sobre o Bloco A)

**Arquivo:** `BASE_Escala_de_SERVIÇO_v2.xlsx` · **Base:** `BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx` (o original segue intacto)

---

## 1. O incidente (transparência)

Minha primeira tentativa de completar a coluna de PJES casou 27 células; a segunda, 121; a terceira — que deveria pegar as 45 restantes — **retornou 0, duas vezes seguidas**. Foi bug meu no casamento da fórmula (regex/string sobre célula obtida por endereço), não problema nos dados.

**Nada foi perdido nem corrompido:** cada passada só gravou o que efetivamente alterou, e a terceira versão (acessando por índice de coluna) fechou as 45. O arquivo foi reverificado do zero ao final.

## 2. O que foi executado

| Item | Ação | Células |
|---|---|---|
| **P1b** | Numeração do bloco de visitantes da 1ªSB corrigida para intervalos de mesmo tamanho e busca linear | **37** |
| **P1b** | Células órfãs removidas: `1ªSB!B29`, `C29`, `ACO!A54`, `B54` | **4** |
| **P9a** | Meta diária (BM e motoristas) das 4 SBs passou a contar `PJ12H` no grupo de 0,5 | **247** |
| **P9b** | Coluna de PJES passou a somar os dois códigos de 12h | **166** |
| **Total** | | **481** |

### Antes → depois

```
1ªSB!A55   =_xlfn.XLOOKUP($D55,RESUMO!$F$2:$F$111,RESUMO!$A$2:$A$112,"",0,1)   →  #N/A
           =_xlfn.XLOOKUP($D55,RESUMO!$F$2:$F$111,RESUMO!$A$2:$A$111,"",0,0)   →  número do cadastro

1ªSB!G30   {"PJ12";"S1"}                        →  {"PJ12";"PJ12H";"S1"}

2ªSB!AL3   COUNTIF(...,"PJ")*2+COUNTIF(...,"SPJ")+COUNTIF(...,"PJ12H")
       →   ...+"PJ12"

1ªSB!AM3   COUNTIF(...,"PJ")*2+COUNTIF(...,"PJ12")
       →   ...+"PJ12H"

2ªSB!AL39  COUNTIF(F39:AJ39,"PJ")*2
       →   ...+"PJ12"+"PJ12H"
```

## 3. Verificação

- `DV`, formatações condicionais e `merges`: **idênticos** ao Bloco A em todas as abas.
- Referências à aba extinta: **0**.
- Metas diárias sem `PJ12H`: **0**.
- Células de PJES com os dois códigos: **193 de 193**.
- `#REF!`: **0**.

> Ressalva já conhecida: o `.xlsx` sai sem valores em cache (openpyxl não calcula). Os números só aparecem ao recalcular no Sheets. Por isso a aplicação real deve ser feita na planilha viva.

## 4. Aplicação no Google Sheets — receita por bloco

Como são centenas de células espalhadas, o caminho é **editar a primeira célula de cada bloco e arrastar**.

### 4.1 Meta diária (2 linhas por aba)

Na linha de **BM/dia** e na de **motoristas/dia**, substitua o trecho `{"PJ12";"S1"}` por `{"PJ12";"PJ12H";"S1"}`, depois arraste até a última coluna de dia.

| Aba | Linhas |
|---|---|
| 1ªSB | 30 e 31 |
| 2ªSB | 31 e 32 |
| 3ªSB | 19 e 20 |
| 4ªSB | 32 e 33 |

### 4.2 Coluna de PJES

Fórmula-modelo (troque a faixa de cada bloco):

```
=COUNTIF(FAIXA,"PJ")*2+COUNTIF(FAIXA,"SPJ")+COUNTIF(FAIXA,"PJ12")+COUNTIF(FAIXA,"PJ12H")
```

| Aba | Coluna | Blocos (arrastar em cada um) |
|---|---|---|
| 1ªSB | `AM` | 3–29 · 39–49 · 55–92 |
| 2ªSB | `AL` | 3–30 · 39–53 |
| 3ªSB | `AL` | 3–18 · 27–41 |
| 4ªSB | `AL` | 3–31 · 40–55 |

> ⚠️ Ao arrastar até o fim do bloco, as **11 células desalinhadas da 2ªSB** (`AL5`, `AL14`, `AL26`, `AL30`, `AK17`, `AK19`, `AK24`, `AK30`…) são normalizadas junto — o que **corrige** a subcontagem, mas **altera números** que hoje estão errados. Isso é desejável, mas você precisa saber que é uma mudança de valor, não só de fórmula.

### 4.3 Numeração dos visitantes (1ªSB)

- `A55:A92`: substituir por `=XLOOKUP($D55; RESUMO!$F$2:$F$111; RESUMO!$A$2:$A$111; ""; 0; 0)` e arrastar.
- Apagar `B29`, `C29` (1ªSB) e `A54`, `B54` (ACO).

---

## 5. O caso ERALDIR — desenho para ADM por SB

### O diagnóstico

`RESUMO!E` carrega **dois significados ao mesmo tempo**:

1. **onde o militar é lançado** (a aba);
2. **qual regime ele segue** (ADM = saldo; demais = ordinário 1x3).

Enquanto havia um só ADM (o da sede), isso funcionava: `E="ADM"` dava as duas informações. Quando cada SB ganhou seu ADM, não passou a existir valor que dissesse *"ADM da 4ªSB"* sem quebrar as contagens que usam `"ADM"` como chave. Foi por isso que o ERALDIR ficou numa linha de ADM com fórmula de saldo — só que **sem o ×2**, contando metade.

### A solução: separar os papéis, sem tocar em `E`

| Onde | O que passa a ser |
|---|---|
| `RESUMO!E` | **não muda em nada** → todas as contagens (`SUMIF "ADM"`, `COUNTIFS`, `FILTER`, listas derivadas, cotas PJES) seguem idênticas |
| `RESUMO!P` (coluna recém-liberada) | **"ADM DA SB"** — preenchida só para ADM: `ADM 7ºGB`, `ADM 1ªSB`, `ADM 2ªSB`, `ADM 3ªSB`, `ADM 4ªSB` |
| Coluna `ORDINÁRIO` | passa a decidir o regime **por essa coluna**, não pela posição da linha |

Fórmula única para a coluna `ORDINÁRIO` (exemplo na 1ªSB, arrastar para toda a coluna):

```
=IF(XLOOKUP($D3; RESUMO!$F$2:$F$400; RESUMO!$P$2:$P$400; "")<>"";
    COUNTIF(G3:AK3;"X")*2+COUNTIF(G3:AK3;"SPJ")+COUNTIF(G3:AK3;"S1");
    COUNTIF(G3:AK3;"X"))
```

### Por que isso resolve o seu problema

- **ERALDIR:** basta `P="ADM 4ªSB"`. Ele continua na 4ªSB, no bloco da própria SB, e passa a contar **4 saldos** corretamente — em qualquer linha que esteja.
- **ADM novo em qualquer SB:** preencher `P` e lançar na aba da SB. A fórmula se ajusta sozinha; ninguém precisa lembrar de trocar a fórmula da linha.
- **Contagens por "ADM":** intactas, porque `E` não muda.
- **Fim da armadilha:** a escolha deixa de ser manual por linha.

### Alternativa que eu não recomendo

Trocar `E` por `"ADM 4ªSB"` e usar curinga `"ADM*"` nas contagens. Funciona, mas obriga a mexer em ~15 fórmulas e **quebra o `INDIRECT` do RESUMO**, que monta o nome da aba a partir de `E` (não existe aba "ADM 4ªSB"). Mesmo ganho, risco bem maior.

### O que muda nas contas com esse desenho

**Nada** nas cotas e nos totais por seção. Muda apenas a coluna `ORDINÁRIO` dos ADMs de SB, que hoje está pela metade (caso ERALDIR) — e essa mudança é a correção.

---

## 6. Arquivos

| Arquivo | Situação |
|---|---|
| `BASE_Escala_de_SERVIÇO.xlsx` | original, intacto |
| `BASE_Escala_de_SERVIÇO_BLOCO_A.xlsx` | Bloco A |
| `BASE_Escala_de_SERVIÇO_v2.xlsx` | Bloco A + P1b + P9 |
| `HISTORICO_2025.xlsx` | 3 abas removidas |
