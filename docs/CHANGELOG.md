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

## Dashboard financeiro profissional

- Reorganizada a aba `Analise` para priorizar recebimentos, custos, resultado e alertas.
- Adicionados cards para Shopee `_kaline98`, Shopee `mateusoliver98`, TikTok Shop e consolidado.
- Adicionados graficos de recebimento por dia, lucro liquido por dia, Shopee vs TikTok Shop e composicao de custos.
- Adicionados filtros globais para hoje, 7 dias, 30 dias e periodo personalizado.
- Adicionado espelho `users/default/productCatalog` com SKU, foto, aliases, links, custos e configuracoes por marketplace.
- Atualizada a gravacao da analise para `users/default/marketplaceOrders`, mantendo leitura dos caminhos legados como fallback.
- Alterada a chave unica dos pedidos para `marketplace + loja + orderId`, consolidando linhas do mesmo pedido.
- Ajustada a consolidacao para evitar dupla contagem quando o mesmo pedido aparece em arquivos de pedidos e em arquivos financeiros.
- Melhorado o parser do TikTok Income para cruzar `Detalhes do pedido` com `Extratos`.
- Melhorada a vinculacao automatica por SKU, itemId/link, alias e nome parecido.
- Adicionados campos profissionais no cadastro de produto: SKU, categoria, aliases, links marketplace e observacoes internas.
- Melhorada a decomposicao de custo para filamento, energia, embalagem, MEI e outros custos.

## Mini ERP financeiro

- Substituida a navegacao de calculadora/analise por `Financeiro`, `Produtos`, `Importacoes`, `Precificacao` e `Configuracoes`.
- A tela inicial agora prioriza `Caiu na sua conta`, `Lucro livre`, `Margem real`, `Ticket medio`, `ROI` e reservas para filamento, energia, embalagens, MEI e manutencao.
- Criada central de produtos com quantidade vendida, recebido, custo, lucro, margem, marketplace principal e badges de negocio.
- Criada tela de importacoes com fluxo profissional, revisao executiva e historico de lotes.
- Criado registro `users/default/marketplaceImports` para resumo dos lotes importados.
- Refeita a precificacao como simulador por Shopee, TikTok Shop e venda direta, com preco minimo, saudavel, recomendado e premium.
- Adicionada tela de configuracoes financeiras dentro da navegacao principal.
- Adicionado reforco de reconhecimento por URL do anuncio e foto semelhante quando os dados estiverem disponiveis.
