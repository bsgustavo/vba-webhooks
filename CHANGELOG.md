# Changelog

Todas as mudanças relevantes do projeto ficam registradas aqui.

Formato: [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/) · Versionamento: [SemVer](https://semver.org/lang/pt-BR/).

## [1.0.0] - 2026-10-02

### Adicionado

- `src/AlertaTeams.bas`: módulo que manda um cartão adaptável para um webhook do Power Automate quando células monitoradas mudam.
  - Dois modos: `Ao salvar`, um cartão por salvamento via `Workbook_AfterSave`, e `A cada alteração`.
  - Intervalos por aba ou em qualquer aba. A mesma célula editada duas vezes aparece uma vez, com o último valor.
  - Rótulos com o cabeçalho da coluna, valores como aparecem na tela e limite de células listadas.
  - Falhas vão para a barra de status e a fila fica para o próximo envio.
  - Cópia somente leitura não envia.
- Aba `Config_Alerta` lida por nomes definidos, com as macros `CriarConfiguracaoAlerta`, `TestarAlertaTeams`, `MostrarConfiguracaoAlerta`, `OcultarConfiguracaoAlerta` e `EnviarPendentes`.
- `src/EstaPastaDeTrabalho.txt`: os eventos `Workbook_SheetChange` e `Workbook_AfterSave` para colar no módulo da pasta.
- `scripts/Gerar-Exemplo.ps1`: gera o `.xlsm` de exemplo a partir de `src/`, sem autor, "salvo por" ou empresa nas propriedades.
- `tests/Executar-Testes.ps1`: 32 verificações num Excel oculto contra um webhook falso local, com envio real opcional.
- `tests/Verificar-Arquivos.ps1`: confere a codificação do `.bas` (Windows-1252, CRLF, sem BOM) e recusa URL de webhook assinada no repositório. Roda no GitHub Actions.

[1.0.0]: https://github.com/bsgustavo/vba-webhooks/releases/tag/v1.0.0
