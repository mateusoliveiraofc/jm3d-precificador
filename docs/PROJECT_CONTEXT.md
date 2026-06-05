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
- Dashboard de lucro real por marketplace/loja.
- Precificacao profissional por Shopee, TikTok Shop e venda direta.

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
- Evita duplicidade por marketplace, loja, pedido e item/SKU.
- Calcula custo real, lucro, margem e status de saude do preco.
- Permite vinculo manual entre produto importado e produto cadastrado.

## Firebase

O app usa Firebase Realtime Database nos caminhos:

- `products`
- `appConfig`
- `marketplaceOrders`
- `marketplaceSettings`
- `productMatchRules`

Os dados existentes devem continuar compativeis. Campos opcionais devem ser enviados como `null` ou omitidos, nunca como `undefined`.

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
