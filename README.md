# Escala 7º GB — CBMPE

Aplicativo web simples para gestão de escala diária do 7º Grupamento de Bombeiros.

## O que faz
- Login com e-mail/senha (Supabase Auth)
- Tabela de militares por seção (1ªSB, 2ªSB, 3ªSB, 4ªSB, ADM)
- Grade diária de escala com códigos (X, PJ, PJ12, PJ12H, S1, SPJ, SB, N)
- Atualização em tempo real: quando um usuário salva, os outros veem instantâneo
- Funciona em celular e computador

## Requisitos
- Projeto grátis no [Supabase](https://supabase.com)
- Navegador (não precisa instalar nada nos computadores dos usuários)

## Como configurar (3 passos)

### 1. Banco de dados
No painel do Supabase, vá em **SQL Editor**, crie uma nova query e cole o conteúdo de `schema.sql`. Execute.

### 2. Usuários
Vá em **Authentication → Users** e cadastre os 5 usuários (e-mail + senha) que vão usar o app.

### 3. Publicar o front-end
1. Pegue o arquivo `index.html` desta pasta.
2. Suba em um serviço de hospedagem estática grátis, como:
   - [Netlify Drop](https://app.netlify.com/drop) (arraste a pasta)
   - [GitHub Pages](https://pages.github.com/)
   - [Vercel](https://vercel.com/)
3. Abra o link gerado e faça login com um dos usuários cadastrados.

## Arquivos
- `index.html` — aplicação completa (login + grade diária + realtime)
- `schema.sql` — estrutura do banco (SQL para copiar no Supabase)
- `README.md` — este arquivo

## Códigos de escala
| Código | Significado |
|--------|-------------|
| X | Serviço ordinário |
| PJ | Plantão 24h |
| PJ12 | Plantão 12h |
| PJ12H | Plantão 12h (horário especial) |
| S1 | Sobreaviso 1 |
| SPJ | Sobreaviso PJ |
| SB | Sem serviço |
| N | Não escalado / folga |