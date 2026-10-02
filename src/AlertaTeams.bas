Attribute VB_Name = "AlertaTeams"
'==============================================================================
' AlertaTeams - avisa no Teams quando células monitoradas da pasta mudam.
'
' Como funciona
'   Cada alteração nas células monitoradas entra numa fila. A fila vira UM
'   cartão, enviado por POST ao webhook do fluxo do Power Automate ("Quando
'   uma solicitação de webhook do Teams é recebida"), conforme a opção
'   "Quando enviar" da aba Config_Alerta:
'     Ao salvar         - um cartão por salvamento, com tudo o que mudou.
'     A cada alteração  - um cartão por edição (colar várias células = 1).
'   O fluxo faz "Aplicar a cada" em destinatarios e posta
'   attachments[0].content como Flow bot.
'
'   Não usa Application.OnTime de propósito: um agendamento pendente faz o
'   Excel reabrir a pasta depois de fechada, e cancelá-lo no BeforeClose
'   não é confiável.
'
' Instalação (uma vez por pasta de trabalho)
'   1. Alt+F11 > Arquivo > Importar arquivo > AlertaTeams.bas
'   2. Cole o conteúdo de EstaPastaDeTrabalho.txt no módulo
'      EstaPastaDeTrabalho (ThisWorkbook).
'   3. Alt+F8 > CriarConfiguracaoAlerta e preencha a aba Config_Alerta.
'   4. Alt+F8 > TestarAlertaTeams.
'   5. Salve como .xlsm.
'==============================================================================
Option Explicit

Private Const ABA_CONFIG As String = "Config_Alerta"
Private Const MAX_FILA As Long = 500        ' teto da fila se o webhook ficar fora do ar
Private Const MAX_TAM_VALOR As Long = 200   ' corta valores longos no cartão

Private mFila As Object          ' Scripting.Dictionary "Aba!A1" -> Array(aba, endereço, cabeçalho, valor)
Private mExcedente As Double     ' células alteradas que não couberam na fila

'------------------------------------------------------------------ eventos

' Chamado por Workbook_SheetChange.
Public Sub RegistrarAlteracao(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Falha                       ' o alerta nunca pode travar quem está preenchendo
    If ThisWorkbook.ReadOnly Then Exit Sub    ' cópia somente leitura: a alteração não vai ser salva
    If Sh.Name = ABA_CONFIG Then Exit Sub
    If Not ConfigExiste() Then Exit Sub
    If Not AlertaAtivo() Then Exit Sub

    Dim alvo As Range
    Set alvo = CelulasMonitoradas(Sh, Target)
    If alvo Is Nothing Then Exit Sub

    If mFila Is Nothing Then Set mFila = CreateObject("Scripting.Dictionary")
    Dim linhaCab As Long, c As Range, chave As String, lidas As Double
    linhaCab = ConfigNumero("AlertaLinhaCabecalho", 1)
    For Each c In alvo.Cells
        chave = Sh.Name & "!" & c.Address(False, False)
        If Not mFila.Exists(chave) And mFila.Count >= MAX_FILA Then Exit For
        mFila(chave) = Array(Sh.Name, c.Address(False, False), Cabecalho(Sh, c, linhaCab), ValorTexto(c))
        lidas = lidas + 1
    Next c
    mExcedente = mExcedente + alvo.CountLarge - lidas

    If EnviaACadaAlteracao() Then EnviarPendentes
    Exit Sub
Falha:
    Application.StatusBar = "Alerta Teams: erro ao registrar a alteração (" & Err.Description & ")"
End Sub

' Envia a fila agora. Chamado por Workbook_AfterSave e, no modo
' "A cada alteração", por RegistrarAlteracao.
Public Sub EnviarPendentes()
    On Error GoTo Falha
    If mFila Is Nothing Then Exit Sub
    If mFila.Count = 0 Then Exit Sub
    If Not ConfigExiste() Then Exit Sub
    If Not AlertaAtivo() Then
        mFila.RemoveAll
        mExcedente = 0
        Exit Sub
    End If
    If DestinatariosJson() = "[]" Then
        Application.StatusBar = "Alerta Teams NÃO enviado: nenhum destinatário na aba " & ABA_CONFIG & "."
        Exit Sub
    End If

    Dim detalhe As String, total As Double
    total = mFila.Count + mExcedente
    If Postar(PayloadAlteracoes(), detalhe) Then
        mFila.RemoveAll
        mExcedente = 0
        Application.StatusBar = "Alerta Teams enviado às " & Format$(Now, "hh:nn") & " (" & Format$(total, "0") & " célula(s))."
    Else
        Application.StatusBar = "Alerta Teams NÃO enviado (" & detalhe & "). Tenta de novo no próximo envio."
    End If
    Exit Sub
Falha:
    Application.StatusBar = "Alerta Teams: erro ao enviar (" & Err.Description & ")"
End Sub

'------------------------------------------------------------------ macros (Alt+F8)

Public Sub CriarConfiguracaoAlerta()
    MsgBox PrepararConfiguracao(), vbInformation, "Alerta Teams"
End Sub

Public Sub TestarAlertaTeams()
    Dim msg As String
    msg = ResultadoTeste()
    MsgBox msg, IIf(Left$(msg, 7) = "Enviado", vbInformation, vbExclamation), "Alerta Teams"
End Sub

Public Sub MostrarConfiguracaoAlerta()
    If AbaConfig() Is Nothing Then
        MsgBox "Rode CriarConfiguracaoAlerta primeiro.", vbExclamation, "Alerta Teams"
        Exit Sub
    End If
    AbaConfig().Visible = xlSheetVisible
    AbaConfig().Activate
End Sub

' Esconde a aba de quem preenche a planilha (só volta por MostrarConfiguracaoAlerta).
Public Sub OcultarConfiguracaoAlerta()
    If AbaConfig() Is Nothing Then Exit Sub
    AbaConfig().Visible = xlSheetVeryHidden
End Sub

Private Function PrepararConfiguracao() As String
    If Not AbaConfig() Is Nothing Then
        AbaConfig().Visible = xlSheetVisible
        AbaConfig().Activate
        PrepararConfiguracao = "A aba " & ABA_CONFIG & " já existe; só deixei visível."
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = ABA_CONFIG

    ws.Range("A1").Value = "Alerta no Teams quando a planilha for alterada"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    LinhaConfig ws, 3, "Ativo", "Sim", "Sim ou Não. Com Não, nada é enviado.", "AlertaAtivo"
    LinhaConfig ws, 4, "URL do webhook", "", "Do gatilho ""Quando uma solicitação de webhook do Teams é recebida"".", "AlertaURL"
    LinhaConfig ws, 5, "Título do alerta", "Planilha alterada", "Primeira linha do cartão.", "AlertaTitulo"
    LinhaConfig ws, 6, "Quando enviar", "Ao salvar", "Ao salvar = um cartão por salvamento, com tudo o que mudou. A cada alteração = um cartão por edição.", "AlertaModo"
    LinhaConfig ws, 7, "Máx. células listadas no cartão", 15, "As demais aparecem como ""e mais N célula(s)"".", "AlertaMaxCelulas"
    LinhaConfig ws, 8, "Linha do cabeçalho", 1, "Linha com o nome das colunas, usado no cartão. 0 = não usar.", "AlertaLinhaCabecalho"

    ws.Range("A10").Value = "Destinatários (um e-mail por linha)"
    ws.Range("C10").Value = "Aba"
    ws.Range("D10").Value = "Intervalo (ex.: B2:F500)"
    ws.Range("A10:D10").Font.Bold = True
    ws.Range("F10").Value = "Sem nenhum intervalo = monitora todas as abas. Aba vazia = vale para qualquer aba. Um intervalo por linha."
    ws.Range("F10").Font.Color = RGB(118, 118, 118)
    ws.Range("A11:A40,C11:D40").NumberFormat = "@"
    ws.Range("A11:A40,C11:D40").Interior.Color = RGB(255, 249, 219)
    ThisWorkbook.Names.Add Name:="AlertaDestinatarios", RefersTo:="='" & ABA_CONFIG & "'!$A$11:$A$40"
    ThisWorkbook.Names.Add Name:="AlertaIntervalos", RefersTo:="='" & ABA_CONFIG & "'!$C$11:$D$40"

    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 60
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 26
    ws.Activate

    PrepararConfiguracao = "Aba " & ABA_CONFIG & " criada. Preencha as células amarelas (URL do webhook, destinatários e intervalos) e rode TestarAlertaTeams." & _
        vbCrLf & vbCrLf & "Depois, OcultarConfiguracaoAlerta esconde a aba de quem preenche a planilha."
End Function

Private Sub LinhaConfig(ByVal ws As Worksheet, ByVal linha As Long, ByVal rotulo As String, _
                        ByVal valor As Variant, ByVal ajuda As String, ByVal nome As String)
    ws.Cells(linha, 1).Value = rotulo
    If VarType(valor) = vbString Then ws.Cells(linha, 2).NumberFormat = "@"
    ws.Cells(linha, 2).Value = valor
    ws.Cells(linha, 2).Interior.Color = RGB(255, 249, 219)
    ws.Cells(linha, 3).Value = ajuda
    ws.Cells(linha, 3).Font.Color = RGB(118, 118, 118)
    ThisWorkbook.Names.Add Name:=nome, RefersTo:="='" & ABA_CONFIG & "'!" & ws.Cells(linha, 2).Address
End Sub

Private Function ResultadoTeste() As String
    If Not ConfigExiste() Then
        ResultadoTeste = "Rode CriarConfiguracaoAlerta primeiro."
        Exit Function
    End If
    Dim dest As String
    dest = DestinatariosJson()
    If dest = "[]" Then
        ResultadoTeste = "Nenhum destinatário na aba " & ABA_CONFIG & "."
        Exit Function
    End If

    Dim fatos As String, detalhe As String
    fatos = Fato("Arquivo", ThisWorkbook.FullName) & "," & _
            Fato("Disparado por", QuemAlterou()) & "," & _
            Fato("Ativo", ConfigTexto("AlertaAtivo")) & "," & _
            Fato("Quando enviar", ConfigTexto("AlertaModo")) & "," & _
            Fato("Monitorando", DescricaoIntervalos())
    If Postar(Envelope(Cartao("Teste: " & ConfigTexto("AlertaTitulo"), _
                              "Se este cartão chegou, a configuração está certa.", fatos, ""), dest), detalhe) Then
        ResultadoTeste = "Enviado (" & detalhe & ")." & vbCrLf & vbCrLf & _
            "O fluxo aceitou o pedido. Confira se o cartão chegou no Teams; se não chegou, veja o histórico de execuções do fluxo."
    Else
        ResultadoTeste = "Falhou: " & detalhe
    End If
End Function

'------------------------------------------------------------------ configuração

Private Function AbaConfig() As Worksheet
    On Error Resume Next
    Set AbaConfig = ThisWorkbook.Worksheets(ABA_CONFIG)
End Function

Private Function ConfigExiste() As Boolean
    Dim r As Range
    On Error Resume Next
    Set r = ThisWorkbook.Names("AlertaURL").RefersToRange
    ConfigExiste = Not r Is Nothing
End Function

Private Function ConfigTexto(ByVal nome As String) As String
    ConfigTexto = TextoDe(ThisWorkbook.Names(nome).RefersToRange.Cells(1, 1))
End Function

Private Function ConfigNumero(ByVal nome As String, ByVal padrao As Long) As Long
    Dim t As String
    t = ConfigTexto(nome)
    If Len(t) > 0 And IsNumeric(t) Then ConfigNumero = CLng(t) Else ConfigNumero = padrao
End Function

Private Function AlertaAtivo() As Boolean
    AlertaAtivo = (UCase$(Left$(ConfigTexto("AlertaAtivo"), 1)) = "S")
End Function

Private Function EnviaACadaAlteracao() As Boolean
    EnviaACadaAlteracao = (InStr(1, ConfigTexto("AlertaModo"), "cada", vbTextCompare) > 0)
End Function

Private Function DestinatariosJson() As String
    Dim c As Range, parte As Variant, email As String, lista As String
    For Each c In ThisWorkbook.Names("AlertaDestinatarios").RefersToRange.Cells
        For Each parte In Split(Replace(TextoDe(c), ",", ";"), ";")
            email = Trim$(parte)
            If InStr(email, "@") > 1 Then
                If Len(lista) > 0 Then lista = lista & ","
                lista = lista & JStr(email)
            End If
        Next parte
    Next c
    DestinatariosJson = "[" & lista & "]"   ' sempre lista, mesmo com 1 e-mail
End Function

' Interseção da alteração com os intervalos configurados (Nothing = ignorar).
Private Function CelulasMonitoradas(ByVal Sh As Object, ByVal Target As Range) As Range
    Dim usadas As Range, lista As Range, trecho As Range, resultado As Range
    Dim i As Long, aba As String, intervalo As String, temRegra As Boolean

    Set usadas = Intersect(Target, Sh.UsedRange)   ' corta linha/coluna inteira
    If usadas Is Nothing Then Exit Function

    Set lista = ThisWorkbook.Names("AlertaIntervalos").RefersToRange
    For i = 1 To lista.Rows.Count
        aba = TextoDe(lista.Cells(i, 1))
        intervalo = Replace(TextoDe(lista.Cells(i, 2)), ";", ",")
        If Len(aba) > 0 Or Len(intervalo) > 0 Then
            temRegra = True
            If Len(aba) = 0 Or StrComp(aba, Sh.Name, vbTextCompare) = 0 Then
                Set trecho = Nothing
                If Len(intervalo) = 0 Then
                    Set trecho = usadas
                Else
                    On Error Resume Next            ' intervalo digitado errado: ignora a linha
                    Set trecho = Intersect(usadas, Sh.Range(intervalo))
                    On Error GoTo 0
                End If
                If Not trecho Is Nothing Then
                    If resultado Is Nothing Then Set resultado = trecho Else Set resultado = Union(resultado, trecho)
                End If
            End If
        End If
    Next i

    If Not temRegra Then Set resultado = usadas   ' nenhum intervalo = monitora tudo
    Set CelulasMonitoradas = resultado
End Function

Private Function DescricaoIntervalos() As String
    Dim lista As Range, i As Long, aba As String, intervalo As String, s As String
    Set lista = ThisWorkbook.Names("AlertaIntervalos").RefersToRange
    For i = 1 To lista.Rows.Count
        aba = TextoDe(lista.Cells(i, 1))
        intervalo = TextoDe(lista.Cells(i, 2))
        If Len(aba) > 0 Or Len(intervalo) > 0 Then
            If Len(aba) = 0 Then aba = "(qualquer aba)"
            If Len(intervalo) = 0 Then intervalo = "(aba toda)"
            If Len(s) > 0 Then s = s & "; "
            s = s & aba & "!" & intervalo
        End If
    Next i
    If Len(s) = 0 Then s = "todas as abas"
    DescricaoIntervalos = s
End Function

'------------------------------------------------------------------ conteúdo do cartão

Private Function PayloadAlteracoes() As String
    Dim maxCel As Long, listadas As Long, total As Double
    Dim k As Variant, it As Variant, rotulo As String, celulas As String
    Dim abas As Object
    Set abas = CreateObject("Scripting.Dictionary")
    maxCel = ConfigNumero("AlertaMaxCelulas", 15)
    If maxCel < 1 Then maxCel = 1
    total = mFila.Count + mExcedente

    For Each k In mFila.Keys
        it = mFila(k)
        abas(it(0)) = True
        If listadas < maxCel Then
            rotulo = it(0) & "!" & it(1)
            If Len(it(2)) > 0 Then rotulo = it(2) & " (" & rotulo & ")"
            If listadas > 0 Then celulas = celulas & ","
            celulas = celulas & Fato(rotulo, CStr(it(3)))
            listadas = listadas + 1
        End If
    Next k
    If total > listadas Then
        celulas = celulas & "," & Fato("...", "e mais " & Format$(total - listadas, "0") & " célula(s)")
    End If

    Dim resumo As String, fatos As String
    resumo = QuemAlterou() & " alterou " & Format$(total, "0") & " célula(s) em " & ThisWorkbook.Name & "."
    fatos = Fato("Arquivo", ThisWorkbook.FullName) & "," & _
            Fato("Aba(s)", Join(abas.Keys, ", ")) & "," & _
            Fato("Quando", Format$(Now, "dd/mm/yyyy hh:nn"))
    PayloadAlteracoes = Envelope(Cartao(ConfigTexto("AlertaTitulo"), resumo, fatos, celulas), DestinatariosJson())
End Function

Private Function Cabecalho(ByVal Sh As Object, ByVal c As Range, ByVal linha As Long) As String
    If linha < 1 Or c.Row <= linha Then Exit Function
    Cabecalho = TextoDe(Sh.Cells(linha, c.Column))
End Function

Private Function ValorTexto(ByVal c As Range) As String
    Dim v As Variant
    v = c.Value
    If IsError(v) Then
        ValorTexto = c.Text                       ' #N/D, #DIV/0!...
    ElseIf IsEmpty(v) Then
        ValorTexto = "(apagado)"
    Else
        ValorTexto = c.Text                       ' como aparece na tela (datas, moeda...)
        If Len(Replace(ValorTexto, "#", "")) = 0 Then ValorTexto = CStr(v)   ' coluna estreita mostra ####
    End If
    If Len(ValorTexto) > MAX_TAM_VALOR Then ValorTexto = Left$(ValorTexto, MAX_TAM_VALOR) & "..."
End Function

Private Function QuemAlterou() As String
    Dim login As String
    login = Environ$("USERNAME")
    If Len(Application.UserName) > 0 And StrComp(Application.UserName, login, vbTextCompare) <> 0 Then
        QuemAlterou = Application.UserName & " (" & login & ")"
    Else
        QuemAlterou = login
    End If
End Function

Private Function TextoDe(ByVal c As Range) As String
    On Error Resume Next   ' célula com erro (#N/D) vira texto vazio
    TextoDe = Trim$(CStr(c.Value))
End Function

'------------------------------------------------------------------ JSON

' Envelope que o fluxo lê: attachments[0].content = cartão; destinatarios = lista.
Private Function Envelope(ByVal cartaoJson As String, ByVal destinatarios As String) As String
    Envelope = "{""type"":""message"",""attachments"":[{""contentType"":""application/vnd.microsoft.card.adaptive"",""content"":" & _
               cartaoJson & "}],""destinatarios"":" & destinatarios & "}"
End Function

Private Function Cartao(ByVal titulo As String, ByVal resumo As String, ByVal fatos As String, ByVal celulas As String) As String
    Dim corpo As String
    corpo = "{""type"":""TextBlock"",""text"":" & JStr(titulo) & ",""weight"":""Bolder"",""size"":""Medium"",""color"":""accent"",""wrap"":true}" & _
            ",{""type"":""TextBlock"",""text"":" & JStr(resumo) & ",""wrap"":true}" & _
            ",{""type"":""FactSet"",""facts"":[" & fatos & "]}"
    If Len(celulas) > 0 Then
        corpo = corpo & ",{""type"":""TextBlock"",""text"":" & JStr("Células alteradas") & ",""weight"":""Bolder"",""spacing"":""Medium"",""wrap"":true}" & _
                        ",{""type"":""FactSet"",""facts"":[" & celulas & "]}"
    End If
    Cartao = "{""$schema"":""http://adaptivecards.io/schemas/adaptive-card.json"",""type"":""AdaptiveCard"",""version"":""1.4"",""body"":[" & corpo & "]}"
End Function

Private Function Fato(ByVal titulo As String, ByVal valor As String) As String
    Fato = "{""title"":" & JStr(titulo) & ",""value"":" & JStr(valor) & "}"
End Function

Private Function JStr(ByVal s As String) As String
    JStr = """" & JsonEscape(s) & """"
End Function

' Escapa para JSON. Acentos viram \uXXXX: o envio não depende da codificação.
Private Function JsonEscape(ByVal s As String) As String
    Dim i As Long, ch As String, cod As Long, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        cod = AscW(ch) And &HFFFF&
        Select Case cod
            Case 34: r = r & "\"""
            Case 92: r = r & "\\"
            Case 10: r = r & "\n"
            Case 13: r = r & "\r"
            Case 9: r = r & "\t"
            Case Is < 32, Is > 126: r = r & "\u" & Right$("000" & Hex$(cod), 4)
            Case Else: r = r & ch
        End Select
    Next i
    JsonEscape = r
End Function

'------------------------------------------------------------------ envio

' POST síncrono (~0,5 s). Timeouts curtos: sem rede, o Excel trava no máximo
' alguns segundos e a fila fica para o próximo envio.
Private Function Postar(ByVal corpo As String, ByRef detalhe As String) As Boolean
    Dim url As String, http As Object
    url = ConfigTexto("AlertaURL")
    If Len(url) = 0 Then
        detalhe = "URL do webhook vazia"
        Exit Function
    End If
    On Error GoTo Falha
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 3000, 3000, 5000, 10000
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send corpo
    detalhe = "HTTP " & http.Status
    Postar = (http.Status >= 200 And http.Status < 300)
    If Not Postar Then detalhe = detalhe & ": " & Left$(http.responseText, 200)
    Exit Function
Falha:
    detalhe = Err.Description
End Function
