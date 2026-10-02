<#
.SYNOPSIS
    Confere o que quebra sem aviso: a codificação do .bas e segredos no repositório.

.DESCRIPTION
    Roda na CI (não precisa de Excel) e localmente:
      - src\AlertaTeams.bas em Windows-1252, CRLF, sem BOM e começando por Attribute VB_Name.
        Salvo em UTF-8 ou com LF, o editor do VBA importa o módulo quebrado.
      - nenhum arquivo versionado com URL de webhook assinada (sig=...).

.EXAMPLE
    pwsh -File .\tests\Verificar-Arquivos.ps1
#>
$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$falhas = 0
function Confere([string]$nome, [bool]$ok, $detalhe = '') {
    if ($ok) { Write-Output "  ok    $nome" }
    else { $script:falhas++; Write-Output "  FALHA $nome  -> $detalhe" }
}

'src\AlertaTeams.bas'
$bytes = [IO.File]::ReadAllBytes((Join-Path $raiz 'src\AlertaTeams.bas'))
$ascii = [Text.Encoding]::ASCII.GetString($bytes)
Confere 'sem BOM' (-not ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF))
Confere 'primeira linha é Attribute VB_Name' ($ascii.StartsWith('Attribute VB_Name = "AlertaTeams"'))
$lfSozinho = [regex]::Matches($ascii, '(?<!\r)\n').Count
Confere 'fim de linha CRLF' ($lfSozinho -eq 0) "$lfSozinho linha(s) só com LF"
$utf8 = $true
try { [void][Text.UTF8Encoding]::new($false, $true).GetString($bytes) } catch { $utf8 = $false }
$temAcento = [bool]($bytes | Where-Object { $_ -gt 127 } | Select-Object -First 1)
Confere 'Windows-1252, não UTF-8' (-not ($utf8 -and $temAcento)) 'o arquivo é UTF-8 válido: salve em Windows-1252'
$indefinidos = @($bytes | Where-Object { $_ -in 0x81, 0x8D, 0x8F, 0x90, 0x9D }).Count
Confere 'só bytes definidos no Windows-1252' ($indefinidos -eq 0) "$indefinidos byte(s) indefinido(s)"

'segredos nos arquivos versionados'
$arquivos = git -C $raiz ls-files
$achados = foreach ($a in $arquivos) {
    $caminho = Join-Path $raiz $a
    if (-not (Test-Path $caminho -PathType Leaf)) { continue }
    $texto = [Text.Encoding]::Latin1.GetString([IO.File]::ReadAllBytes($caminho))
    if ($texto -match 'sig=[A-Za-z0-9_\-]{20,}') { $a }
}
Confere 'nenhuma URL de webhook assinada' (@($achados).Count -eq 0) ($achados -join ', ')

Write-Output ''
if ($falhas -gt 0) { Write-Output "$falhas falha(s)"; exit 1 }
Write-Output 'tudo certo'
