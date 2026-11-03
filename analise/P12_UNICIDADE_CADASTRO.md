# P12 — Unicidade e formato de cadastro (RESUMO)

> Objetivo: impedir cadastro duplicado e formato de matrícula inválido, pegando o erro na origem (RESUMO) antes de se espalhar pelas SBs.

---

## 1. Duplicidade de matrícula

A coluna `C` (matrícula) é a chave do cadastro. Hoje não há validação de unicidade.

### Formatação condicional (aplicar em C2:C112)

```
=CONT.SE($C$2:$C$112; C2)>1
```

Destaca em vermelho qualquer matrícula repetida. O caso MONTEIRO (duplicado na 1ªSB) seria pego aqui se estivesse duplicado no RESUMO — no RESUMO não há duplicidade de matrícula, o problema está no bloco de visitantes da 1ªSB (já resolvido no Bloco A/P2, removendo a linha 82).

---

## 2. Formato de matrícula

O padrão é `NNNNNNN-N` (7 dígitos, hífen, 1 dígito), ex.: `707155-8`.

### Validação de dados (aplicar em C2:C112)

No Google Sheets: Dados → Validação de dados → Critério: "Fórmula personalizada":

```
=OU(C2=""; CORRESP(C2; REGEXEXTRACT(C2; "^[0-9]{7}-[0-9]{1}$"); 0)=1)
```

Ou mais simples, aceitar vazio ou padrão:

```
=SE(C2=""; VERDADEIRO; CORRESP(TEXTO(C2; "0000000-0"); ... ))
```

**Simplificação:** usar expressão regular direta na validação se o Sheets suportar `REGEXMATCH`:

```
=OU(C2=""; REGEXMATCH(C2; "^[0-9]{7}-[0-9]$"))
```

Se `REGEXMATCH` não estiver disponível, usar validação por comprimento + posição do hífen:

```
=OU(C2=""; E(COMPRIMENTO(C2)=9; ESQUERDA(C2;7)<>" "; PARCIAL(C2;8;1)="-"; ÉNÚMERO(VALUE(ESQUERDA(C2;7))); ÉNÚMERO(VALUE(DIREITA(C2;1)))))
```

---

## 3. Nomes com espaço sobrando (3 casos)

Já identificados:

| Linha | Nome de guerra | Correção |
|-------|---------------|----------|
| D85 | `JEFERSON ` | `JEFERSON` |
| D88 | `AMANDA BRITO ` | `AMANDA BRITO` |
| D109 | `MENEZES ` | `MENEZES` |

**Atenção:** corrigir apenas o RESUMO **não basta**. As SBs usam `VLOOKUP`/`XLOOKUP` sobre a coluna `F` (nome completo, que é fórmula `TEXTJOIN` de B+C+D). Se D muda, F muda sozinho (é fórmula). Mas se alguma SB tem o nome gravado **em texto** (não via fórmula), ela ficará desatualizada → `#N/A` silencioso.

**Procedimento seguro:**
1. Corrigir os 3 nomes no RESUMO (coluna D).
2. Verificar que a coluna `F` (fórmula `TEXTJOIN`) recalcula automaticamente.
3. Conferir que nenhum `VLOOKUP` nas SBs retorna `#N/A` para esses militares.
4. Só então ativar o `TRIM` geral (se for adotado).

---

## 4. TRIM geral — armadilha

Não aplicar `TRIM` em massa na coluna `D` ou `F` sem fazer o mesmo nas SBs. A chave de busca entre abas é o nome exato; `TRIM` só no RESUMO quebraria os `VLOOKUP` das SBs que ainda têm o nome com espaço.

Recomendação: corrigir os 3 casos à mão e documentar "não digitar espaço final" na `COMO USAR` (P4). Adiar `TRIM` automático para quando houver uma rotina de limpeza sincronizada entre abas.

---

## 5. Folga nas listas de validação

Hoje o RESUMO valida escala/seção só até a linha 112 — 1 linha de folga para 110 militares. Ampliar para linha 200 (ou 400, alinhado com P5).

---

## Arquivos afetados

| Item | Onde |
|------|------|
| CF duplicidade | `RESUMO!C2:C112` |
| Validação matrícula | `RESUMO!C2:C112` |
| Correção 3 nomes | `RESUMO!D85`, `RESUMO!D88`, `RESUMO!D109` |
| Folga validação | listas de validação do RESUMO (ampliar alcance) |

---

**Status:** especificação pronta. Correção dos 3 nomes à mão + CF de duplicidade + validação de matrícula aplicáveis imediatamente no Google Sheets.
