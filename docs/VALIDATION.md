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

## Catalogo e compartilhamento

- Exportar produto individual.
- Exportar catalogo completo.
- Confirmar que o HTML gerado abre fora do app.

## PWA / iPhone

Se forem adicionados ou alterados manifest e service worker:

- Testar Safari.
- Testar "Adicionar a Tela de Inicio".
- Preservar safe area com `env(safe-area-inset-top)` e `env(safe-area-inset-bottom)`.
