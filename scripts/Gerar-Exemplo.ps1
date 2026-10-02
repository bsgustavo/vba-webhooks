<#
.SYNOPSIS
    Gera uma pasta de trabalho .xlsm de exemplo com o módulo AlertaTeams instalado.

.DESCRIPTION
    Abre um Excel oculto, importa src\AlertaTeams.bas, cola src\EstaPastaDeTrabalho.txt no
    módulo da pasta, cria a aba Config_Alerta e uma aba Registros de exemplo monitorada em
    A2:E500. O arquivo sai sem URL e sem destinatário, a menos que você passe -WebhookUrl
    e -Destinatario.

    Importar código por automação exige a opção "Confiar no acesso ao modelo de objeto do
    projeto do VBA". O script liga essa opção (HKCU, só para o seu usuário) enquanto roda e
    devolve o valor anterior no fim, mesmo se der erro.

.PARAMETER Saida
    Caminho do .xlsm gerado. Padrão: dist\Alerta-Teams-Exemplo.xlsm.

.PARAMETER WebhookUrl
    URL do gatilho "Quando uma solicitação de webhook do Teams é recebida". Opcional.

.PARAMETER Destinatario
    Um ou mais e-mails que recebem o cartão. Opcional.

.EXAMPLE
    pwsh -File .\scripts\Gerar-Exemplo.ps1

.EXAMPLE
    pwsh -File .\scripts\Gerar-Exemplo.ps1 -WebhookUrl $url -Destinatario ana@empresa.com
#>
[CmdletBinding()]
param(
    [string]$Saida = (Join-Path $PSScriptRoot '..\dist\Alerta-Teams-Exemplo.xlsm'),
    [string]$WebhookUrl,
    [string[]]$Destinatario
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
$bas = Join-Path $raiz 'src\AlertaTeams.bas'
$evt = Join-Path $raiz 'src\EstaPastaDeTrabalho.txt'
$Saida = [IO.Path]::GetFullPath($Saida)
New-Item -ItemType Directory -Force (Split-Path $Saida) | Out-Null
if (Test-Path $Saida) { Remove-Item $Saida -Force }

Add-Type -Namespace VbaWebhooks -Name Win32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);'

$chave = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $chave)) { New-Item $chave -Force | Out-Null }
$antes = (Get-ItemProperty $chave -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty $chave -Name AccessVBOM -Value 1 -Type DWord

$xl = $null
$excelPid = 0
try {
    $xl = New-Object -ComObject Excel.Application
    [void][VbaWebhooks.Win32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$excelPid)
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    $wb = $xl.Workbooks.Add()
    [void]$wb.VBProject.VBComponents.Import($bas)
    # CodeName é "EstaPastaDeTrabalho" no Excel em português e "ThisWorkbook" em inglês
    $wb.VBProject.VBComponents.Item($wb.CodeName).CodeModule.AddFromString([IO.File]::ReadAllText($evt))

    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Registros'
    $cabecalho = 'Nº', 'Data', 'Setor', 'Descrição', 'Status'
    for ($i = 0; $i -lt $cabecalho.Count; $i++) { $ws.Cells.Item(1, $i + 1).Value2 = $cabecalho[$i] }
    $ws.Range('A1:E1').Font.Bold = $true
    $ws.Range('B:B').NumberFormat = 'dd/mm/yyyy'
    $ws.Columns.Item(4).ColumnWidth = 40

    [void]$xl.Run('AlertaTeams.PrepararConfiguracao')
    $intervalos = $wb.Names.Item('AlertaIntervalos').RefersToRange
    $intervalos.Cells.Item(1, 1).Value2 = 'Registros'
    $intervalos.Cells.Item(1, 2).Value2 = 'A2:E500'
    if ($WebhookUrl) { $wb.Names.Item('AlertaURL').RefersToRange.Value2 = $WebhookUrl }
    $destinos = $wb.Names.Item('AlertaDestinatarios').RefersToRange
    $linha = 1
    foreach ($email in @($Destinatario | Where-Object { $_ })) { $destinos.Cells.Item($linha, 1).Value2 = $email; $linha++ }

    $ws.Activate()
    $wb.SaveAs($Saida, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
}
finally {
    if ($xl) {
        try { $xl.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    $ws = $wb = $intervalos = $destinos = $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($null -eq $antes) { Remove-ItemProperty $chave -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { Set-ItemProperty $chave -Name AccessVBOM -Value $antes -Type DWord }
    if ($excelPid -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) {
        Start-Sleep -Seconds 3
        Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue   # só o Excel aberto por este script
    }
}

# Autor, "salvo por" e empresa saem do XML do pacote. RemovePersonalInformation não serve:
# em pasta com macro ele abre um aviso modal que o DisplayAlerts não suprime e trava o script.
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$pacote = [IO.Compression.ZipFile]::Open($Saida, 'Update')
try {
    foreach ($parte in 'docProps/core.xml', 'docProps/app.xml') {
        $entrada = $pacote.GetEntry($parte)
        if (-not $entrada) { continue }
        $leitor = [IO.StreamReader]::new($entrada.Open()); $xml = $leitor.ReadToEnd(); $leitor.Close()
        foreach ($tag in 'dc:creator', 'cp:lastModifiedBy', 'Company', 'Manager') {
            $xml = $xml -replace "<$tag>[^<]*</$tag>", "<$tag></$tag>"
        }
        $entrada.Delete()
        $escritor = [IO.StreamWriter]::new($pacote.CreateEntry($parte).Open(), [Text.UTF8Encoding]::new($false))
        $escritor.Write($xml); $escritor.Close()
    }
}
finally { $pacote.Dispose() }
Write-Output "Gerado: $Saida"
