<#
.SYNOPSIS
    Testa o módulo AlertaTeams num Excel oculto, contra um webhook falso local.

.DESCRIPTION
    Gera a pasta de trabalho com scripts\Gerar-Exemplo.ps1 (em tests\tmp), sobe um servidor
    HTTP em 127.0.0.1 que grava cada POST recebido e responde 202, e exercita a planilha por
    automação: alterações dentro e fora do intervalo, modo "Ao salvar" e "A cada alteração",
    colar bloco grande, apagar, Ativo = Não, webhook fora do ar, fechar sem salvar e cópia
    somente leitura. Cada POST capturado é conferido campo a campo.

    Nenhum cartão sai para o Teams, a não ser que você passe -WebhookUrl e -Destinatario:
    aí o último teste manda um cartão de verdade.

    Requer Windows com Excel desktop e PowerShell 7.

.PARAMETER WebhookUrl
    Opcional. URL real do fluxo, para o teste final que envia um cartão ao Teams.

.PARAMETER Destinatario
    Opcional. E-mail que recebe o cartão do teste final.

.EXAMPLE
    pwsh -File .\tests\Executar-Testes.ps1
#>
[CmdletBinding()]
param(
    [string]$WebhookUrl,
    [string]$Destinatario
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$tmp = Join-Path $PSScriptRoot 'tmp'
$xlsm = Join-Path $tmp 'teste.xlsm'
$capDir = Join-Path $tmp 'captura'
if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
New-Item -ItemType Directory $capDir -Force | Out-Null

& (Join-Path $raiz 'scripts\Gerar-Exemplo.ps1') -Saida $xlsm | Out-Null

# ---- webhook falso: grava cada POST em captura\reqNN.json e responde 202
$sonda = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0); $sonda.Start()
$porta = $sonda.LocalEndpoint.Port; $sonda.Stop()
$urlFalsa = "http://127.0.0.1:$porta/webhook"
$servidor = Start-ThreadJob -ArgumentList $porta, $capDir -ScriptBlock {
    param($port, $dir)
    $l = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $port); $l.Start()
    $n = 0; $fim = (Get-Date).AddMinutes(10)
    while ((Get-Date) -lt $fim -and -not (Test-Path (Join-Path $dir 'parar'))) {
        if (-not $l.Pending()) { Start-Sleep -Milliseconds 30; continue }
        $c = $l.AcceptTcpClient(); $s = $c.GetStream(); $s.ReadTimeout = 5000
        $buf = New-Object byte[] 65536; $ms = [IO.MemoryStream]::new(); $fimCab = -1
        while ($fimCab -lt 0) {
            $r = $s.Read($buf, 0, $buf.Length); if ($r -le 0) { break }
            $ms.Write($buf, 0, $r); $fimCab = [Text.Encoding]::ASCII.GetString($ms.ToArray()).IndexOf("`r`n`r`n")
        }
        $tudo = $ms.ToArray(); $cab = [Text.Encoding]::ASCII.GetString($tudo, 0, $fimCab)
        $len = 0; if ($cab -match 'Content-Length:\s*(\d+)') { $len = [int]$Matches[1] }
        while ($tudo.Length - ($fimCab + 4) -lt $len) {
            $r = $s.Read($buf, 0, $buf.Length); if ($r -le 0) { break }
            $ms.Write($buf, 0, $r); $tudo = $ms.ToArray()
        }
        $n++
        [IO.File]::WriteAllBytes((Join-Path $dir ('req{0:00}.json' -f $n)), [byte[]]$tudo[($fimCab + 4)..($fimCab + 3 + $len)])
        [IO.File]::WriteAllText((Join-Path $dir ('req{0:00}.head' -f $n)), $cab)
        $resp = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 202 Accepted`r`nContent-Length: 0`r`nConnection: close`r`n`r`n")
        $s.Write($resp, 0, $resp.Length); $s.Flush(); $c.Close()
    }
    $l.Stop()
}

$script:total = 0
$script:falhas = 0
function Confere([string]$nome, [bool]$ok, $detalhe = '') {
    $script:total++
    if ($ok) { Write-Output "  ok    $nome" }
    else { $script:falhas++; Write-Output "  FALHA $nome  -> $detalhe" }
}
function Qtd { @(Get-ChildItem $capDir -Filter 'req*.json').Count }
function Req([int]$n) {
    $bytes = [IO.File]::ReadAllBytes((Join-Path $capDir ('req{0:00}.json' -f $n)))
    [pscustomobject]@{
        Ascii = -not ($bytes | Where-Object { $_ -gt 127 })
        Json  = [Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json
        Cab   = Get-Content (Join-Path $capDir ('req{0:00}.head' -f $n)) -Raw
    }
}
function Celulas($json) { $json.attachments[0].content.body[4].facts }
function Resumo($json) { $json.attachments[0].content.body[1].text }
function BarraStatus { $s = $xl.StatusBar; if ($s -is [bool]) { '(padrão)' } else { $s } }
function Cfg([string]$nome) { , $wb.Names.Item($nome).RefersToRange }   # vírgula: o PowerShell não desmonta o Range em células

Add-Type -Namespace VbaWebhooksTeste -Name Win32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'
$xl = $null
$excelPid = 0
try {
    $xl = New-Object -ComObject Excel.Application
    [void][VbaWebhooksTeste.Win32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$excelPid)
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($xlsm)
    $ws = $wb.Worksheets.Item('Registros')

    'T0 configuração gerada'
    Confere 'aba Config_Alerta existe' ($null -ne $wb.Worksheets.Item('Config_Alerta'))
    Confere 'modo padrão = Ao salvar' ((Cfg 'AlertaModo').Value2 -eq 'Ao salvar') (Cfg 'AlertaModo').Value2
    Confere 'sai sem URL e sem destinatário' (-not (Cfg 'AlertaURL').Value2 -and -not (Cfg 'AlertaDestinatarios').Cells.Item(1, 1).Value2)
    (Cfg 'AlertaURL').Value2 = $urlFalsa
    (Cfg 'AlertaTitulo').Value2 = 'Planilha de exemplo alterada'
    (Cfg 'AlertaDestinatarios').Cells.Item(1, 1).Value2 = 'pessoa@exemplo.com'
    Confere 'mexer na configuração não envia' ((Qtd) -eq 0) (Qtd)

    'T1 TestarAlertaTeams'
    $r = $xl.Run('AlertaTeams.ResultadoTeste')
    Confere 'retorno Enviado (HTTP 202)' ($r -like 'Enviado (HTTP 202)*') $r
    $q = Req 1
    Confere 'destinatarios vai como lista' ($q.Json.destinatarios -is [array] -and $q.Json.destinatarios[0] -eq 'pessoa@exemplo.com') ($q.Json.destinatarios | ConvertTo-Json -Compress)
    Confere 'cartão em attachments[0].content' ($q.Json.attachments[0].contentType -eq 'application/vnd.microsoft.card.adaptive' -and $q.Json.attachments[0].content.type -eq 'AdaptiveCard')
    Confere 'título do teste' ($q.Json.attachments[0].content.body[0].text -eq 'Teste: Planilha de exemplo alterada') $q.Json.attachments[0].content.body[0].text
    Confere 'corpo só ASCII (acentos em \u)' $q.Ascii
    Confere 'Content-Type json utf-8' ($q.Cab -match 'Content-Type: application/json; charset=utf-8') $q.Cab

    'T2 modo Ao salvar'
    $ws.Range('F5').Value2 = 'fora do intervalo'
    $ws.Range('A2').Value2 = 'NC-0001'
    $ws.Range('B2').Value2 = 46296   # 01/10/2026
    $ws.Range('C2').Value2 = 'Usinagem'
    $desc = "Rebarba na peça ""X"" — ação imediata`nsegunda linha \ fim"
    $ws.Range('D2').Value2 = $desc
    $ws.Range('E2').Value2 = 'Aberta'
    $ws.Range('C2').Value2 = 'Usinagem CNC'
    Confere 'nada enviado antes de salvar' ((Qtd) -eq 1) (Qtd)
    $wb.Save()
    Confere 'salvar envia um cartão' ((Qtd) -eq 2) (Qtd)
    $q = Req 2; $cel = Celulas $q.Json
    Confere '5 células (F5 fora, C2 deduplicada)' (@($cel).Count -eq 5) (($cel | ForEach-Object title) -join ' | ')
    Confere 'rótulo usa o cabeçalho' ($cel[0].title -eq 'Nº (Registros!A2)') $cel[0].title
    Confere 'data como aparece na tela' (($cel | Where-Object title -like 'Data*').value -eq '01/10/2026') ($cel | Where-Object title -like 'Data*').value
    Confere 'reedição fica com o último valor' (($cel | Where-Object title -like 'Setor*').value -eq 'Usinagem CNC') ($cel | Where-Object title -like 'Setor*').value
    Confere 'aspas, acentos, quebra e barra preservados' (($cel | Where-Object title -like 'Descri*').value -eq $desc) ($cel | Where-Object title -like 'Descri*').value
    Confere 'resumo "alterou 5 célula(s)"' ((Resumo $q.Json) -like '* alterou 5 célula(s) em teste.xlsm.') (Resumo $q.Json)
    Confere 'barra de status confirma' ((BarraStatus) -like 'Alerta Teams enviado às *(5 célula(s)).') (BarraStatus)
    $wb.Save()
    Confere 'salvar sem alteração não envia' ((Qtd) -eq 2) (Qtd)

    'T3 modo A cada alteração'
    (Cfg 'AlertaModo').Value2 = 'A cada alteração'
    $ws.Range('E2').Value2 = 'Em análise'
    Confere 'uma edição = um cartão na hora' ((Qtd) -eq 3) (Qtd)
    $cel = Celulas (Req 3).Json
    Confere 'cartão traz a célula editada' (@($cel).Count -eq 1 -and $cel[0].title -eq 'Status (Registros!E2)' -and $cel[0].value -eq 'Em análise') ($cel | ConvertTo-Json -Compress)

    'T4 colar bloco grande (28 x 5)'
    $ws.Range('A3:E30').Value2 = 'bloco colado'
    Confere 'um bloco = um cartão' ((Qtd) -eq 4) (Qtd)
    $q = Req 4; $cel = Celulas $q.Json
    Confere '15 listadas + "e mais 125"' (@($cel).Count -eq 16 -and $cel[15].value -eq 'e mais 125 célula(s)') ('{0} fatos, último: {1}' -f @($cel).Count, $cel[-1].value)
    Confere 'resumo conta 140' ((Resumo $q.Json) -like '* alterou 140 célula(s) *') (Resumo $q.Json)

    'T5 apagar'
    $ws.Range('A3:E30').ClearContents()
    $cel = Celulas (Req 5).Json
    Confere 'conteúdo apagado aparece como (apagado)' ((Qtd) -eq 5 -and $cel[0].value -eq '(apagado)') ($cel[0] | ConvertTo-Json -Compress)

    'T6 Ativo = Não'
    (Cfg 'AlertaAtivo').Value2 = 'Não'
    $ws.Range('E2').Value2 = 'x'
    Confere 'nada enviado' ((Qtd) -eq 5) (Qtd)
    (Cfg 'AlertaAtivo').Value2 = 'Sim'

    'T7 webhook fora do ar'
    (Cfg 'AlertaURL').Value2 = 'http://127.0.0.1:9/ninguem'
    $ms = (Measure-Command { $ws.Range('E2').Value2 = 'Fechada' }).TotalMilliseconds
    Confere 'barra de status avisa a falha' ((BarraStatus) -like 'Alerta Teams NÃO enviado*') (BarraStatus)
    Confere ('Excel não trava (edição levou {0:N0} ms)' -f $ms) ($ms -lt 4000)
    (Cfg 'AlertaURL').Value2 = $urlFalsa
    (Cfg 'AlertaModo').Value2 = 'Ao salvar'
    $wb.Save()
    $cel = Celulas (Req 6).Json
    Confere 'fila da falha vai no próximo envio' ((Qtd) -eq 6 -and $cel[0].value -eq 'Fechada') ($cel | ConvertTo-Json -Compress)

    'T8 fechar sem salvar'
    $ws.Range('E3').Value2 = 'descartado'
    $wb.Close($false)
    Start-Sleep -Seconds 3
    $outra = $xl.Workbooks.Add(); Start-Sleep -Seconds 2
    Confere 'nada enviado e a pasta não reabre' ((Qtd) -eq 6 -and $xl.Workbooks.Count -eq 1) ('req={0} pastas={1}' -f (Qtd), (($xl.Workbooks | ForEach-Object Name) -join ','))
    $outra.Close($false)

    'T9 cópia somente leitura'
    $ro = $xl.Workbooks.Open($xlsm, 0, $true)
    $ro.Worksheets.Item('Registros').Range('A5').Value2 = 'cópia'
    $xl.Run('AlertaTeams.EnviarPendentes')
    Confere 'somente leitura não enfileira' ($ro.ReadOnly -and (Qtd) -eq 6) ('ro={0} req={1}' -f $ro.ReadOnly, (Qtd))
    $ro.Close($false)

    if ($WebhookUrl -and $Destinatario) {
        'T10 cartão real no Teams'
        $wb = $xl.Workbooks.Open($xlsm)
        $ws = $wb.Worksheets.Item('Registros')
        $xl.EnableEvents = $false
        $ws.Range('A2:F40').ClearContents()
        (Cfg 'AlertaURL').Value2 = $WebhookUrl
        (Cfg 'AlertaDestinatarios').Cells.Item(1, 1).Value2 = $Destinatario
        $xl.EnableEvents = $true
        $ws.Range('A2').Value2 = 'NC-0123'
        $ws.Range('B2').Value2 = 46296
        $ws.Range('C2').Value2 = 'Usinagem'
        $ws.Range('D2').Value2 = 'Exemplo enviado pelos testes do vba-webhooks'
        $ws.Range('E2').Value2 = 'Aberta'
        $wb.Save()
        Confere 'webhook real aceitou (confira o Teams)' ((BarraStatus) -like 'Alerta Teams enviado às *(5 célula(s)).') (BarraStatus)
        $wb.Close($false)
    }
}
catch {
    $script:falhas++
    Write-Output ('ERRO: {0} (linha {1})' -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber)
}
finally {
    if ($xl) {
        try { $xl.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    $ws = $wb = $ro = $outra = $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    New-Item (Join-Path $capDir 'parar') -ItemType File -Force | Out-Null
    $servidor | Wait-Job -Timeout 10 | Out-Null
    $servidor | Remove-Job -Force
    if ($excelPid -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) {
        Start-Sleep -Seconds 3
        Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue   # só o Excel aberto por este script
    }
}

Write-Output ''
Write-Output ('{0} verificações, {1} falha(s)' -f $script:total, $script:falhas)
if ($script:falhas -gt 0) { exit 1 }
