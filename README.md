# JM 3D Precificador

Aplicativo Android com WebView e versao HTML para calculo de precos de produtos impressos em 3D, gerenciamento de catalogo e sincronizacao online via Firebase Realtime Database.

## Estado atual

O projeto ja esta funcional e deve ser tratado como uma base estavel. As funcionalidades existentes incluem:

- Calculo de precificacao para produtos impressos em 3D.
- Cadastro, listagem, visualizacao e exclusao de produtos.
- Sincronizacao online via Firebase Realtime Database.
- Fallback local usando `localStorage` quando houver falha de conexao.
- App Android em WebView.
- Exportacao/compartilhamento de orcamento e catalogo em HTML.
- Persistencia de configuracoes.
- Tela de PIN para acesso.

## Estrutura principal

- `app/src/main/java/com/jm3d/precificador/MainActivity.kt`: activity Android, WebView, seletor de imagem e ponte nativa de compartilhamento.
- `app/src/main/res/layout/activity_main.xml`: layout Android com WebView em tela cheia.
- `app/src/main/res/values/themes.xml`: tema Android com barras transparentes.
- `app/src/main/assets/index.html`: aplicacao web principal, incluindo UI, estilos, calculos, Firebase, catalogo e exportacao HTML.
- `app/src/main/assets/logo.png` e `app/src/main/assets/logo_small.png`: assets de marca.

## Regras de manutencao

Antes de alterar qualquer comportamento, leia o codigo existente e preserve a implementacao atual sempre que possivel. Mudancas devem ser pequenas, documentadas e validadas.

Nao remover nem alterar sem necessidade:

- Ajustes de safe area/status bar no Android.
- Configuracao de WebView em `MainActivity.kt`.
- Logica de Firebase e fallback local.
- Estrutura de dados existente no Realtime Database.
- Recursos de compartilhamento/exportacao HTML.
- Persistencia de configuracoes.

## Pontos criticos

### Safe area Android

O Android usa `WindowCompat.setDecorFitsSystemWindows(window, false)` e barras transparentes. O conteudo web precisa respeitar safe area no CSS, especialmente no topo da tela.

No HTML atual, os pontos principais sao:

- `#hdr` com `env(safe-area-inset-top)`.
- `.sht` com `env(safe-area-inset-bottom)`.

Ao modificar layout principal, validar em telas com notch/status bar, especialmente Samsung S25 Ultra ou dispositivos equivalentes.

### Firebase

O Firebase Realtime Database nao aceita propriedades com valor `undefined`. Antes de enviar objetos para o Firebase:

- Remover propriedades `undefined`.
- Usar `null` quando o campo opcional deve existir sem valor.
- Preservar compatibilidade com dados antigos.
- Nao alterar a estrutura do banco sem migracao.

### PWA / HTML

A versao HTML precisa continuar funcionando em navegador movel. Ao adicionar manifest ou service worker, preservar compatibilidade com Safari/iPhone e safe area.

## Validacao recomendada

Quando houver alteracoes:

1. Revisar o diff dos arquivos modificados.
2. Verificar se `app/src/main/assets/index.html` continua abrindo no navegador.
3. Testar calculo com campos vazios e preenchidos.
4. Testar salvamento online e fallback local.
5. Testar exportacao/compartilhamento HTML.
6. Testar Android com build debug quando o Gradle wrapper estiver disponivel.

Observacao: neste workspace foi encontrado `gradle/wrapper/gradle-wrapper.properties`, mas nao foram encontrados `gradlew` ou `gradlew.bat` na raiz.

## Repositorio

Repositorio remoto oficial:

https://github.com/mateusoliveiraofc/jm3d-precificador
