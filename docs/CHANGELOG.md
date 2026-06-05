# Changelog

## Inicial

- Documentada a estrutura atual do projeto.
- Registradas regras para preservar safe area Android.
- Registrada regra de sanitizacao antes de salvar no Firebase.
- Registrado checklist de validacao.
- Criado `.gitignore` para evitar versionamento de arquivos locais e gerados.

## Analise de lucro por marketplace

- Criada aba `Analise` com dashboard, importacao, pedidos e precificacao.
- Adicionado parser em lote para XLSX/CSV da Shopee e TikTok Shop.
- Adicionada biblioteca SheetJS local para leitura de XLSX.
- Criados caminhos Firebase separados para pedidos importados, configuracoes de marketplace e regras de vinculo.
- Adicionada sanitizacao antes de gravar produtos, configuracoes e pedidos importados.
- Adicionados `manifest.json` e `sw.js` para suporte PWA.
