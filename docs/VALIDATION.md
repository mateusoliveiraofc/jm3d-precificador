# Validacao

Use este checklist antes de considerar uma mudanca concluida.

## Codigo

- Conferir arquivos modificados.
- Evitar alteracoes em `build/`, `.gradle/`, `.idea/` e `local.properties`.
- Preservar comportamento existente quando a mudanca nao exigir refatoracao.

## Android

- Confirmar que `MainActivity.kt` ainda carrega `file:///android_asset/index.html`.
- Confirmar que `WindowCompat.setDecorFitsSystemWindows(window, false)` nao foi removido sem justificativa.
- Confirmar que o tema continua com barras transparentes se o layout depender disso.
- Build recomendado quando o wrapper estiver disponivel:

```powershell
.\gradlew.bat assembleDebug
```

## HTML / WebView

- Abrir `app/src/main/assets/index.html` no navegador para checar renderizacao basica.
- Testar calculo com:
  - peso vazio;
  - tempo vazio;
  - plastico bolha padrao;
  - preco manual;
  - lucro padrao.
- Testar foto opcional.
- Testar produto sem descricao.

## Firebase

- Salvar produto com todos os campos.
- Salvar produto sem foto.
- Salvar produto com campos opcionais vazios.
- Confirmar que nenhum objeto enviado contem `undefined`.
- Confirmar fallback local em falha de conexao.
- Importar o mesmo lote de relatorios duas vezes e confirmar que duplicados sao ignorados.
- Confirmar que pedidos importados sao salvos em `users/default/marketplaceOrders`, sem alterar `products`.
- Confirmar que produtos antigos sao espelhados em `users/default/productCatalog` sem remover o cadastro original.

## Catalogo e compartilhamento

- Exportar produto individual.
- Exportar catalogo completo.
- Confirmar que o HTML gerado abre fora do app.

## PWA / iPhone

Se forem adicionados ou alterados manifest e service worker:

- Testar Safari.
- Testar "Adicionar a Tela de Inicio".
- Preservar safe area com `env(safe-area-inset-top)` e `env(safe-area-inset-bottom)`.

## Relatorios Shopee/TikTok

- Selecionar varios arquivos ao mesmo tempo.
- Misturar XLSX e CSV no mesmo lote.
- Testar Shopee Income.lancado.
- Testar Shopee Order.completed.
- Testar TikTok income.
- Testar TikTok Enviado pedido CSV.
- Confirmar cruzamento do TikTok Income quando `Detalhes do pedido` e `Extratos` estiverem em abas separadas.
- Confirmar tela de revisao antes de salvar.
- Confirmar dashboard consolidado.
- Confirmar filtros por loja e marketplace.
- Confirmar produto nao vinculado.
- Confirmar vinculo manual de produto.
- Exportar CSV da analise.

## Validacao executada nesta revisao

- `node --check app/src/main/assets/marketplace.js`
- `node --check app/src/main/assets/sw.js`
- Parser em lote com 6 arquivos reais:
  - Shopee `Income.lancado`: 27 + 1 linhas.
  - Shopee `Order.completed`: 29 + 10 linhas.
  - TikTok `income_20260604221509`: 3 linhas cruzando detalhes e extratos.
  - TikTok `Enviado pedido`: 54 linhas.
  - Consolidado: 92 pedidos unicos, 0 chaves duplicadas.
- Teste de vinculo automatico com produto simulado:
  - SKU.
  - itemId/link.
  - alias.
- Smoke test no navegador local em `http://127.0.0.1:8765/index.html`:
  - Cadastro/calculo abre com os novos campos profissionais.
  - Aba `Analise` renderiza dashboard financeiro.
  - Aba `Importar` renderiza upload multiplo.
  - Aba `Precificacao` renderiza custos globais e tabelas por produto.
  - Confirmados 4 graficos no dashboard.
