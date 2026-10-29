# SFG Financeiro

Sistema de contas a pagar e a receber, extrato, viabilidade por lote, fluxo de caixa e checklist de documentos da SFG Incorporadora.

- Funciona no Mac e no iPhone (Safari → Compartilhar → **Adicionar à Tela de Início**).
- Funciona sem internet: os lançamentos ficam guardados no aparelho e sincronizam sozinhos quando a conexão volta.
- Dados no Firebase (Firestore) do projeto `sfg-financeiro`, acessíveis só pelos e-mails listados nas regras.

## Estrutura

- `fonte/src/` — código-fonte (`main.js`, `index.html`, `sw.js`, `config.js`, ícones, fontes)
- raiz do repositório — versão publicada (gerada por `node build.mjs` dentro de `fonte/`), servida pelo GitHub Pages
- `fonte/firestore.rules` — regras de segurança com perfis (colar no console do Firebase → Firestore → Regras)

## Atualizar o sistema

```
npm install
node build.mjs
```
Rode dentro de `fonte/`. Os arquivos do app são gerados na raiz do repositório; faça commit e push. Os aparelhos recebem a nova versão na próxima abertura com internet.

## Trocar ou incluir um usuário

1. Firebase → Authentication → Usuários → **Adicionar usuário** com o novo e-mail.
2. Firestore → Regras: inclua o e-mail na lista `['luismatagalvao@gmail.com', 'novo@sfg...']` e publique.
3. Teste o login com o novo e-mail e, se quiser, remova o antigo da lista.

## Backup

Menu **Dados e backup** → *Exportar Excel* (pode ser reimportado) ou *Backup completo (.json)*.
