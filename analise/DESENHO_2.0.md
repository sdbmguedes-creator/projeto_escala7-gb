# ESCALA 2.0 — DESENHO

> Construção nova, em Google Sheets, incorporando todas as regras e correções aprendidas no modelo atual.
> Nada é importado do modelo antigo além dos **dados** (cadastro e marcações).

---

## Decisões fechadas com o responsável (01/11/2026)

| Tema | Decisão |
|---|---|
| Antiguidade | é a **ordem da lista**, editável; muda quando o GB recebe ou transfere militar |
| Lista de militares | **editável**: incluir, retirar e reordenar |
| Saída de militar | **apagar a linha** (sem inativação) |
| Escala e ADM | **não são fixos** — definidos mês a mês pela **DOp** |
| Virada do mês | **continua copiando a planilha** |
| Cotas PJES | vêm do processo **SEI**, transcritas pela DOp |
| Ambiente | Google Sheets **compartilhado** com os escalantes de cada SB |
| Excel | **não é ambiente de trabalho**; o `.xlsx` é só veículo de importação |

### Consequência crítica da cópia mensal

Copiar é o que corrompeu o modelo atual (faixas que envelhecem, fórmulas por célula). Para o 2.0 ser **à prova de cópia**:

1. nenhuma fórmula referencia faixa fixa de linha — sempre **coluna inteira** ou **intervalo nomeado**;
2. o mês inteiro deriva de **uma célula de competência**;
3. a única entrada manual na grade são as **marcações**;
4. listas vêm de `FILTER`/`SORTBY`, nunca de cópia de nomes;
5. a limpeza do mês é **apagar a área de marcação**, não reescrever estrutura.

---

## Princípios

1. **Um dado, um lugar.** Nome, escala, serviço e afastamento existem uma vez só.
2. **O que pode ser derivado, é derivado.** Listas, efetivo, efetivo por dia, cotas, somas.
3. **Regra vira tabela, não fórmula espalhada.** Meta, peso e limite ficam visíveis.
4. **O mês é parâmetro, não cópia.** Uma célula de competência manda em tudo.
5. **Conferência embutida.** A planilha avisa; ninguém precisa procurar.

---

## Abas

| Aba | Papel | Como é preenchida |
|---|---|---|
| `PARAMETROS` | metas por SB, limite masculino/feminino, peso dos códigos, cotas | manual (raro) |
| `CADASTRO` | uma linha por militar: ID, posto, nome de guerra, seção, escala, motorista, ADM | manual |
| `AFASTAMENTOS` | ID, tipo, dia inicial, dia final | manual |
| `ESCALAS` | rodízio do ano (qual escala cai em cada dia) | **fórmula** (âncora + MOD) |
| `EFETIVO` | quem está em cada SB, por antiguidade | **derivado** (FILTER do cadastro) |
| `GRADE 1ªSB` … `GRADE 4ªSB` | grade visual do mês | nomes derivados; **marcação manual** |
| `VISITANTES` | militares de outras SBs e o serviço deles | **derivado** (espelho) |
| `DIARIO` | BM/dia, motoristas/dia, déficit, médias | **derivado** |
| `PJES` | cotas por SB, ledger, saldo | cotas derivadas; ledger manual |
| `CONFERENCIAS` | painel de pendências | **derivado** |

---

## As cinco decisões que resolvem os problemas do modelo atual

### 1. Identidade por matrícula, não por texto

O `CADASTRO` tem uma coluna **ID** (a matrícula). Todo vínculo usa o ID.
Nome de guerra, posto e seção podem mudar sem quebrar nada — e dois "Silvas" deixam de ser um risco.

### 2. Afastamento é dado, não marcação

Uma linha em `AFASTAMENTOS`:

| ID | Tipo | De | Até |
|---|---|---|---|
| 710281-0 | Férias | 1 | 15 |
| 711134-7 | Férias | 16 | 30 |
| 707276-7 | Afastamento geral | 1 | 31 |

**O `N` deixa de ser necessário.** A grade pinta o período de cinza, e **todos os contadores ignoram** os dias dentro da janela. Se alguém digitar `N`, é tratado igual — nada quebra.

Isso resolve de uma vez: férias fracionadas, afastamento total, e o "N automático" (que por fórmula seria impossível).

### 3. Visitante é espelho, não digitação

`VISITANTES` mostra, para cada militar de fora, o serviço que ele tem na **SB de origem**, em tempo real, com `SB` no lugar de `PJ` (para não contar cota na SB que recebe).

E — o ponto que você pediu — **vale nas quatro SBs**, não só na 1ªSB.

### 4. Escala do dia é derivada

`ESCALAS` é gerada por fórmula a partir de uma data-âncora do ciclo:
`=CHOOSE(MOD(data - âncora; 4) + 1; "BRAVO"; "CHARLIE"; "DELTA"; "ALFA")`

Com isso a conferência "serviço fora da escala" passa a ser estrutural: a própria grade sabe qual escala deveria estar de serviço naquele dia.

### 5. Regras na tabela

`PARAMETROS` guarda meta de BM, meta de motoristas, limite masculino/feminino e o peso de cada código:

| Código | Peso no dia | Cota PJES | Saldo ADM |
|---|---|---|---|
| X | 1 | — | 2 |
| PJ | 1 | 2 | — |
| PJ12 / PJ12H | 0,5 | 1 / 0 | — |
| S1 | 0,5 | — | 1 |
| SPJ | 1 | 1 | 1 |
| SB | 0 | 0 | 0 |
| N | 0 | 0 | 0 |

Mudar um peso = uma célula.

---

## Regras de negócio que o 2.0 precisa respeitar

- **Rodízio** ALFA → BRAVO → CHARLIE → DELTA, 4 dias, 1x3 para operacionais.
- **ADM** não segue 1x3: cumpre 4 saldos/mês, serviços avulsos.
- **Ordem dos serviços:** `X` seguido de `PJ`/`PJ12` **pode**; `PJ` seguido de `X` **não** (12h de descanso).
- **PJES:** 24h = 2 cotas, 12h = 1 cota; cota pertence à SB de origem.
- **Motorista:** `M` operacional entra na conta do dia; `B` administrativo não.
- **FEM:** limite por dia na 1ªSB + ACO.
- **Afastamento:** férias 1ª metade, 2ª metade, mês inteiro; `X` = geral.

---

## Conferências (painel único)

| Verificação | Critério |
|---|---|
| Serviço fora da escala do dia | marcação ≠ escala vigente |
| Serviço em dia de afastamento | marcação dentro da janela |
| PJES antes do ordinário | `PJ` seguido de `X` |
| ADM com saldo ≠ 4 | saldo do mês |
| Operacional com ORD ≠ 7–8 | contagem de `X` |
| Duas SBs no mesmo dia | mesmo ID marcado em duas grades |
| Cadastro incompleto | ID, seção ou escala em branco |
| Cadastro duplicado | mesmo ID duas vezes |

O painel mostra **contagem por verificação** e um link para a linha. Sem pintar a planilha inteira.

---

## Migração (como sair do atual sem risco)

1. **Importar o cadastro** dos 110 militares do `RESUMO` atual.
2. **Importar os afastamentos** de outubro (as janelas que deduzimos do `N`).
3. **Rodar outubro no 2.0** e comparar com o atual: 328 cotas, BM/dia, motoristas/dia, saldo por militar.
4. **Validar contra agosto, setembro e novembro** — é o teste que nenhum modelo novo pode pular.
5. Só então virar a chave, **na virada de um mês**, com o antigo congelado como contingência.

**Critério de aceite:** o 2.0 reproduz os totais do mês anterior com diferença zero. Se não reproduzir, o modelo está errado — não o dado.

---

## O que depende dos arquivos que você vai mandar

| Arquivo | Para quê |
|---|---|
| agosto, setembro, novembro | ver variações reais: mês de 30/31 dias, entrada e saída de militar, cota diferente, feriados |
| novembro | reproduzir o mês **em curso** no 2.0 e comparar com o que está valendo |

Sem eles eu construo no escuro — e a única coisa que o 2.0 não pode ser é uma suposição minha sobre como os meses variam.
