# Como contribuir

Obrigado por ajudar. Relatos de bug, correções e melhorias pequenas são bem-vindos.

## Antes de começar

- Para qualquer coisa maior que uma correção pequena, abra uma issue antes, para combinarmos o caminho.
- Nunca cole a URL do seu webhook numa issue, num pull request ou num log. Ela é a credencial do fluxo.

## Ambiente

Você precisa de Windows, Excel desktop e PowerShell 7.

```powershell
git clone https://github.com/bsgustavo/vba-webhooks.git
cd vba-webhooks
pwsh -File .\tests\Verificar-Arquivos.ps1
pwsh -File .\tests\Executar-Testes.ps1
```

O `Executar-Testes.ps1` abre um Excel oculto e manda os POSTs para um servidor falso em `127.0.0.1`, então nenhum cartão sai para o Teams. Os dois scripts precisam passar na sua máquina antes do merge de um pull request. A CI roda só o `Verificar-Arquivos.ps1`, porque os runners do GitHub não têm Excel.

## Regras do projeto

- **Nada de `Application.OnTime`.** Um agendamento pendente faz o Excel reabrir a pasta depois de fechada, e cancelá-lo no `BeforeClose` não é confiável (ver README).
- **O alerta nunca interrompe quem preenche.** Erro vai para a barra de status, não para `MsgBox`. Caixa de diálogo só nas macros que a pessoa roda de propósito (`CriarConfiguracaoAlerta`, `TestarAlertaTeams`).
- **Todo texto que vai para o JSON passa por `JStr`.** Ele escapa aspas, barras, quebras de linha e acentos (`\uXXXX`).
- **O contrato do webhook é público.** Mudar o formato de `destinatarios` ou do envelope quebra os fluxos já criados: é `feat!` ou `BREAKING CHANGE`.
- **Arquivos:**
  - `src/AlertaTeams.bas` fica em **Windows-1252 com CRLF**, porque é o que o editor do VBA importa. O `Verificar-Arquivos.ps1` confere.
  - Os `.ps1` ficam em UTF-8 com BOM e CRLF (ver `.editorconfig`).
  - Código, comentários e documentação em português.
- **Nenhuma pasta de trabalho no repositório.** O exemplo é gerado pelo `scripts/Gerar-Exemplo.ps1` e publicado como anexo da release.

## Commits

[Conventional Commits](https://www.conventionalcommits.org), em português:

```
<tipo>(<escopo opcional>): <descrição no imperativo, minúsculo, sem ponto final>

<corpo opcional: o porquê da mudança; o diff já mostra o quê>
```

- Tipos: `feat`, `fix`, `docs`, `style`, `refactor`, `perf`, `test`, `build`, `ci`, `chore`, `revert`.
- Assunto com até 50 caracteres, no imperativo ("adiciona", não "adicionado"), sem emoji.
- Um commit por mudança lógica.
- Mudança que quebra compatibilidade: `feat!:` ou rodapé `BREAKING CHANGE:`.

Exemplos:

```
fix: escapa tabulacao nos valores do cartao
feat(cartao): mostra o valor anterior da celula
```

## Branches e pull requests

- Crie a branch a partir da `main` como `<tipo>/<descricao-curta-em-kebab>`, por exemplo `fix/escapa-tabulacao`.
- Título do pull request no mesmo formato de Conventional Commit.
- Preencha o template: contexto, o que muda, como testar e a saída dos testes.

## Versões

[SemVer](https://semver.org/lang/pt-BR/): `fix` sobe o PATCH, `feat` sobe o MINOR e mudança que quebra compatibilidade sobe o MAJOR. Toda versão ganha uma tag `vX.Y.Z`, uma release no GitHub (com o `.xlsm` de exemplo gerado pelo `Gerar-Exemplo.ps1`) e uma entrada no [CHANGELOG.md](CHANGELOG.md) (formato [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/)).
