# SFG Financeiro — instruções para o Claude

Sistema financeiro de obras da SFG Incorporadora (PWA offline + Firebase).

- Código-fonte em `fonte/` (`src/main.js` é o app inteiro; `src/index.html` traz o CSS).
- A raiz do repositório é o app publicado pelo GitHub Pages: https://sfgincorporadora.github.io/financeiro/
- Para publicar uma mudança: editar `fonte/src/…`, rodar `cd fonte && npm install && node build.mjs` (gera app.js, index.html, sw.js, icons, fonts e manifest na raiz) e fazer commit/push na `main`.
- Testes de interface usam um Firebase simulado: `node test/build-test.mjs`, servir `test/www` em 127.0.0.1:5173 e rodar os scripts Python de `test/` (precisam de uma planilha em `test/SFG.xlsx`, que nunca deve ir para o repositório).
- Nunca commitar dados reais (planilhas, extratos, PDFs, backups). O `.gitignore` bloqueia esses formatos.
- Regras de segurança do Firestore: `fonte/firestore.rules` (são coladas manualmente no console do Firebase).
- Dono do sistema: luismatagalvao@gmail.com (constante `DONOS` em main.js e lista `dono()` nas regras).
