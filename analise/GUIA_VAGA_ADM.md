# GUIA DE MIGRAÇÃO — Vaga do ADM da SB

> Executar no **Google Sheets** (não no .xlsx). Tempo estimado: 10–15 min.
> Faça backup antes: Arquivo → Fazer download → XLSX.

> **Idioma (pt-BR):** separador de argumentos `;` e decimal `,`. O Sheets em português
> **mostra os nomes de função traduzidos** (`SE`, `SEERRO`, `ÍNDICE`, `CORRESP`, `CONT.SE`,
> `SOMARPRODUTO`) e aceita tanto o nome em português quanto o em inglês. Dentro de chaves,
> o `;` separa linhas da matriz: `{"x";"PJ";"SPJ"}`.

> **Atenção ao `PROCX` / `XLOOKUP`:** o último argumento (modo de busca) aceita `1` (binária,
> padrão), `-1`, `2` (linear) e `-2`. **`0` não existe no Sheets e gera `#VALUE`.** Como a
> ordem dos nomes não é alfabética, busca binária pode devolver outra pessoa — por isso as
> fórmulas deste guia usam `ÍNDICE`+`CORRESP`, que é sempre exata.

---

## Por que no Sheets e não no arquivo de referência

Ao inserir uma linha, o Sheets **realinha automaticamente** todas as referências — dentro da aba, entre abas, no `DIÁRIO` e no `CONF. INDIV.`. A biblioteca que eu uso move células mas **não reescreve fórmulas**: eu teria de corrigir centenas de referências manualmente, com risco de erro silencioso.

---

## Resumo do que muda

| Aba | Inserir linha? | Onde | Efeito |
|---|---|---|---|
| 2ªSB | **não** | — | ADM já está na linha 3 |
| 3ªSB | **não** | — | ADM já está na linha 3 |
| 1ªSB | **sim, 1** | na linha 3 | cria a vaga; TELES desce para a 4 |
| 4ªSB | **sim, 1** | na linha 3 | cria a vaga; ERALDIR desce para a 4 |

---

## PASSO 1 — 2ªSB e 3ªSB (sem inserir nada)

Só trocar o nome por fórmula, na célula `D3`:

**2ªSB!D3**
```
=SEERRO(ÍNDICE(RESUMO!$F$2:$F$400;CORRESP("ADM 2ªSB";RESUMO!$P$2:$P$400;0));"")
```

**3ªSB!D3**
```
=SEERRO(ÍNDICE(RESUMO!$F$2:$F$400;CORRESP("ADM 3ªSB";RESUMO!$P$2:$P$400;0));"")
```

Pronto. Se não houver ADM definido no mês, a célula fica vazia e a linha vira uma vaga em branco.

---

## PASSO 2 — 4ªSB (inserir 1 linha)

### 2.1 Inserir
Clique com o botão direito no número da **linha 3** → *Inserir 1 linha acima*.

Resultado: a vaga passa a ser a linha 3, e ERALDIR desce para a linha 4.

### 2.2 Copiar a estrutura da linha
Selecione a **linha 4 inteira** → copiar → colar na **linha 3**.
Depois, na linha 3:
- `D3` → apague e cole a fórmula da vaga:
```
=SEERRO(ÍNDICE(RESUMO!$F$2:$F$400;CORRESP("ADM 4ªSB";RESUMO!$P$2:$P$400;0));"")
```
- `G3:AJ3` → apague (eram as marcações do ERALDIR copiadas junto)

### 2.3 Refazer as linhas de total
As linhas de total desceram para **33 (BM)**, **34 (motoristas)** e **35 (PJES)**.
Selecione `F33:AJ33`, apague e cole **uma vez** (a fórmula se ajusta às colunas):

**BM (linha 33)**
```
=SUMPRODUCT(COUNTIF(F$3:F$32;{"x";"PJ";"SPJ"}))+(SUMPRODUCT(COUNTIF(F$3:F$32;{"PJ12";"PJ12H";"S1"}))*0,5)+SUMPRODUCT(COUNTIF(F$41:F$56;{"x";"PJ";"SPJ"}))+(SUMPRODUCT(COUNTIF(F$41:F$56;{"PJ12";"PJ12H";"S1"}))*0,5)
```

**Motoristas (linha 34)**
```
=SUMPRODUCT(COUNTIFS($C$3:$C$32;"M";F$3:F$32;{"x";"PJ";"SPJ"}))+(SUMPRODUCT(COUNTIFS($C$3:$C$32;"M";F$3:F$32;{"PJ12";"PJ12H";"S1"}))*0,5)+SUMPRODUCT(COUNTIFS($C$41:$C$56;"M";F$41:F$56;{"x";"PJ";"SPJ"}))+(SUMPRODUCT(COUNTIFS($C$41:$C$56;"M";F$41:F$56;{"PJ12";"PJ12H";"S1"}))*0,5)
```

> O intervalo `F$3:F$32` inclui a vaga (3) + os 29 militares (4–32). O bloco de visitantes passou a 41–56.

---

## PASSO 3 — 1ªSB (inserir 1 linha)

### 3.1 Inserir
Botão direito na **linha 3** → *Inserir 1 linha acima*.
A vaga é a linha 3; TELES desce para a 4; os totais vão para **31 (BM)** e **32 (motoristas)**; o bloco ADM para 40–50; os visitantes para 56–93.

### 3.2 Copiar a estrutura
Selecione a **linha 4** → copiar → colar na **linha 3**. Depois:
- `D3` → apague e cole:
```
=SEERRO(ÍNDICE(RESUMO!$F$2:$F$400;CORRESP("ADM 1ªSB";RESUMO!$P$2:$P$400;0));"")
```
- `G3:AK3` → apague as marcações copiadas

### 3.3 Refazer as linhas de total

**BM (linha 31)** — selecione `G31:AK31`, apague e cole:
```
=SUMPRODUCT(COUNTIF(G$3:G$30;{"x";"PJ";"SPJ"}))+(SUMPRODUCT(COUNTIF(G$3:G$30;{"PJ12";"PJ12H";"S1"}))*0,5)+SUMPRODUCT(COUNTIF(G$40:G$50;{"x";"PJ";"SPJ"}))+(SUMPRODUCT(COUNTIF(G$40:G$50;{"PJ12";"PJ12H";"S1"}))*0,5)+SUMPRODUCT(COUNTIF(G$56:G$93;"PJ"))+(SUMPRODUCT(COUNTIF(G$56:G$93;"PJ12"))*0,5)
```

**Motoristas (linha 32)** — selecione `G32:AK32`, apague e cole:
```
=SUMPRODUCT(COUNTIFS($C$3:$C$30;"M";G$3:G$30;{"x";"PJ";"SPJ"}))+(SUMPRODUCT(COUNTIFS($C$3:$C$30;"M";G$3:G$30;{"PJ12";"PJ12H";"S1"}))*0,5)+SUMPRODUCT(COUNTIFS($C$40:$C$50;"M";G$40:G$50;{"x";"PJ";"SPJ"}))+(SUMPRODUCT(COUNTIFS($C$40:$C$50;"M";G$40:G$50;{"PJ12";"PJ12H";"S1"}))*0,5)+SUMPRODUCT(COUNTIFS($C$56:$C$93;"M";G$56:G$93;"PJ"))+(SUMPRODUCT(COUNTIFS($C$56:$C$93;"M";G$56:G$93;"PJ12"))*0,5)
```

---

## PASSO 4 — Marcar o ADM no RESUMO

Coluna `P` do `RESUMO` (lista suspensa com `ADM 7ºGB`, `ADM 1ªSB`, `ADM 2ªSB`, `ADM 3ªSB`, `ADM 4ªSB`).
- Hoje: JOÃO LIANDA = `ADM 2ªSB`, ANDRADE = `ADM 3ªSB`, os 11 da sede = `ADM 7ºGB`.
- 1ªSB e 4ªSB: **vazio** (não havia ADM em outubro).

**Para definir:** ache a linha do militar e escolha o valor na lista. Nada de digitar nome.

---

## PASSO 5 — Conferência de marcação órfã

Formatação condicional em `G3:AJ3` (2ª/3ª/4ªSB) e `G3:AK3` (1ªSB):

```
=AND(COUNTA(G3:AJ3)>0; $D3="")
```

Se alguém marcar serviço na linha do ADM sem que haja ADM definido (ou depois de trocar o ADM sem limpar), a linha acende — é o sinal para limpar as marcações antigas.

---

## Verificação final

1. Cada aba tem **uma** vaga de ADM na linha 3 (preenchida ou vazia).
2. `MILITARES POR DIA` continua somando o número de antes (nenhum militar saiu).
3. `DIÁRIO` mostra os mesmos totais do mês anterior.
4. Nenhuma célula com `#REF!`.
5. `CONF. INDIV.` continua encontrando qualquer militar (para incluir o ADM, ajuste os intervalos que começam em `$D$4` para `$D$3`).

---

## Efeito no dia a dia

| Situação | O que acontece |
|---|---|
| Tem ADM no mês | nome aparece sozinho na linha 3 da SB |
| Não tem ADM | linha 3 fica vazia e não conta |
| ADM muda | troca o `P` no RESUMO; o nome troca sozinho |
| Marcação esquecida | a conferência acende a linha |
