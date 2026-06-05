# Contexto do Projeto

## Objetivo

O JM 3D Precificador calcula precos de produtos impressos em 3D e gerencia um catalogo sincronizado online. O app roda como Android WebView e tambem como HTML local/exportavel.

## Funcionalidades existentes

- Calculo de custo e preco sugerido.
- Cadastro de produto.
- Listagem de produtos.
- Visualizacao de produto.
- Exclusao de produto.
- Sincronizacao online.
- Fallback local.
- Configuracoes persistentes.
- Compartilhamento de produto em HTML.
- Compartilhamento de catalogo em HTML.
- Upload e compressao local de foto.
- PIN de acesso.
- Importacao em lote de relatorios Shopee/TikTok.
- Mini ERP financeiro com recebimentos, reservas, resultado, alertas, produtos, importacoes e simulador.
- Precificacao profissional por Shopee, TikTok Shop e venda direta.
- Catalogo profissional espelhado com SKU, aliases e links por marketplace.

## Arquitetura atual

O projeto tem uma camada Android pequena e a maior parte da regra de negocio dentro de `app/src/main/assets/index.html`.

### Android

`MainActivity.kt`:

- Configura WebView.
- Habilita JavaScript, DOM storage e acesso a arquivos.
- Abre seletor de imagem.
- Expoe `Bridge.shareHtml` para compartilhamento nativo.
- Carrega `file:///android_asset/index.html`.

### Web

`index.html`:

- Renderiza toda a interface.
- Calcula custos e precos.
- Inicializa Firebase.
- Mantem cache local.
- Renderiza catalogo.
- Exporta HTML.
- Gerencia configuracoes.

`marketplace.js`:

- Processa XLSX e CSV usando SheetJS local.
- Detecta Shopee Income, Shopee Order.completed, TikTok income e TikTok Orders CSV.
- Normaliza dinheiro, datas e nomes de colunas.
- Cruza TikTok Income quando os detalhes do pedido e extratos financeiros estao em abas separadas.
- Evita duplicidade por marketplace, loja e pedido.
- Calcula custo real, lucro, margem e status de saude do preco.
- Permite vinculo manual entre produto importado e produto cadastrado.
- Faz vinculo automatico por SKU, itemId/link, alias e nome parecido.
- Renderiza as telas `Financeiro`, `Produtos`, `Importacoes`, `Precificacao` e `Configuracoes`.

## Firebase

O app usa Firebase Realtime Database nos caminhos:

- `products`
- `appConfig`
- `users/default/productCatalog`
- `users/default/marketplaceOrders`
- `users/default/marketplaceImports`
- `users/default/marketplaceSettings`
- `users/default/productMatchRules`

Os caminhos legados `marketplaceOrders`, `marketplaceSettings` e `productMatchRules` ainda sao lidos como fallback, mas novas gravacoes da analise devem ir para `users/default/...`.

Os dados existentes devem continuar compativeis. Campos opcionais devem ser enviados como `null` ou omitidos, nunca como `undefined`.

`products` continua sendo a estrutura principal do cadastro antigo. O novo `productCatalog` e um espelho/migracao aditiva usado pela analise profissional e pela vinculacao dos relatorios.

`marketplaceImports` armazena apenas resumo de lotes de importacao. A chave de deduplicacao dos pedidos continua em `marketplaceOrders`.

## Problemas ja resolvidos que nao devem voltar

### Sobreposicao da status bar

O conteudo ja foi ajustado para nao ficar escondido pela barra superior do Android. Preservar os ajustes de janela Android e safe area CSS.

### Erro de undefined no Firebase

O Firebase falha com:

```text
set failed: value argument contains undefined
```

Qualquer novo fluxo de escrita precisa sanitizar objetos antes de enviar ao banco.

## Observacoes importantes

- A pasta atual nao estava inicializada como repositorio Git.
- O repositorio remoto informado estava vazio no momento da configuracao inicial.
- O projeto possui arquivos de build gerados que nao devem ser versionados.
- `local.properties` contem configuracao local de SDK e nao deve ser commitado.
