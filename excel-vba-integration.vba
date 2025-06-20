' Código VBA para Excel - Integração com Sistema de Monitoramento
' Este código deve ser adicionado a um módulo VBA no Excel

' Configurações da API
Const API_ENDPOINT As String = "https://script.google.com/macros/s/AKfycbwSwlJATYl9L0GHOwNrGzRnhRsrNbaZedUd0lLGujwiF4noP8xHP8dUH9SrfVh7fAi0Sw/exec"
Const SECURITY_TOKEN As String = "SEU_TOKEN_SEGURO_AQUI" ' Substitua pelo token real

' Função principal para atualizar equipamento via Excel
Sub AtualizarEquipamentoViaExcel()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tag As String
    Dim status As String
    Dim motivo As String
    Dim pts As String
    Dim os As String
    Dim retorno As String
    Dim cadeado As String
    Dim observacoes As String
    Dim modificadoPor As String
    Dim resultado As String
    
    ' Definir a planilha ativa
    Set ws = ActiveSheet
    
    ' Verificar se a planilha tem o formato correto
    If ws.Cells(1, 1).Value <> "TAG" Then
        MsgBox "A planilha deve ter o cabeçalho 'TAG' na célula A1", vbCritical
        Exit Sub
    End If
    
    ' Encontrar a última linha com dados
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Nenhum dado encontrado para atualizar", vbInformation
        Exit Sub
    End If
    
    ' Confirmar a operação
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja atualizar " & (lastRow - 1) & " equipamento(s) no sistema?", vbYesNo + vbQuestion)
    
    If resposta = vbNo Then
        Exit Sub
    End If
    
    ' Processar cada linha de dados
    For i = 2 To lastRow
        tag = Trim(ws.Cells(i, 1).Value) ' Coluna A - TAG
        status = Trim(ws.Cells(i, 2).Value) ' Coluna B - STATUS
        motivo = Trim(ws.Cells(i, 3).Value) ' Coluna C - MOTIVO
        pts = Trim(ws.Cells(i, 4).Value) ' Coluna D - PTS
        os = Trim(ws.Cells(i, 5).Value) ' Coluna E - OS
        retorno = Trim(ws.Cells(i, 6).Value) ' Coluna F - RETORNO
        cadeado = Trim(ws.Cells(i, 7).Value) ' Coluna G - CADEADO
        observacoes = Trim(ws.Cells(i, 8).Value) ' Coluna H - OBSERVACOES
        modificadoPor = Trim(ws.Cells(i, 9).Value) ' Coluna I - MODIFICADO_POR
        
        ' Validar dados obrigatórios
        If tag = "" Then
            ws.Cells(i, 10).Value = "ERRO: TAG obrigatória"
            GoTo NextRow
        End If
        
        If status = "" Then
            ws.Cells(i, 10).Value = "ERRO: STATUS obrigatório"
            GoTo NextRow
        End If
        
        ' Enviar atualização para a API
        resultado = EnviarAtualizacaoParaAPI(tag, status, motivo, pts, os, retorno, cadeado, observacoes, modificadoPor)
        
        ' Registrar resultado na coluna J
        ws.Cells(i, 10).Value = resultado
        
        ' Atualizar a tela
        Application.StatusBar = "Processando linha " & i & " de " & lastRow
        DoEvents
        
NextRow:
    Next i
    
    Application.StatusBar = False
    MsgBox "Processamento concluído! Verifique a coluna J para os resultados.", vbInformation
End Sub

' Função para enviar atualização para a API
Function EnviarAtualizacaoParaAPI(tag As String, status As String, motivo As String, pts As String, os As String, retorno As String, cadeado As String, observacoes As String, modificadoPor As String) As String
    Dim http As Object
    Dim jsonData As String
    Dim response As String
    
    ' Criar objeto HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Preparar dados JSON
    jsonData = "{"
    jsonData = jsonData & """type"": ""atualizacaoExcel"","
    jsonData = jsonData & """token"": """ & SECURITY_TOKEN & ""","
    jsonData = jsonData & """TAG"": """ & EscapeJSON(tag) & ""","
    jsonData = jsonData & """STATUS"": """ & EscapeJSON(status) & ""","
    jsonData = jsonData & """MOTIVO"": """ & EscapeJSON(motivo) & ""","
    jsonData = jsonData & """PTS"": """ & EscapeJSON(pts) & ""","
    jsonData = jsonData & """OS"": """ & EscapeJSON(os) & ""","
    jsonData = jsonData & """RETORNO"": """ & EscapeJSON(retorno) & ""","
    jsonData = jsonData & """CADEADO"": """ & EscapeJSON(cadeado) & ""","
    jsonData = jsonData & """OBSERVACOES"": """ & EscapeJSON(observacoes) & ""","
    jsonData = jsonData & """MODIFICADO_POR"": """ & EscapeJSON(modificadoPor) & ""","
    jsonData = jsonData & """DATA"": """ & Format(Now, "yyyy-mm-ddThh:mm:ss") & """"
    jsonData = jsonData & "}"
    
    ' Configurar requisição HTTP
    http.Open "POST", API_ENDPOINT, False
    http.setRequestHeader "Content-Type", "application/json"
    
    ' Enviar requisição
    On Error GoTo ErrorHandler
    http.send jsonData
    
    ' Processar resposta
    If http.Status = 200 Then
        response = http.responseText
        
        ' Verificar se a resposta indica sucesso
        If InStr(response, """success"":true") > 0 Then
            EnviarAtualizacaoParaAPI = "SUCESSO"
        Else
            ' Extrair mensagem de erro se possível
            Dim errorMsg As String
            errorMsg = ExtrairMensagemErro(response)
            EnviarAtualizacaoParaAPI = "ERRO: " & errorMsg
        End If
    Else
        EnviarAtualizacaoParaAPI = "ERRO HTTP: " & http.Status
    End If
    
    Exit Function
    
ErrorHandler:
    EnviarAtualizacaoParaAPI = "ERRO: " & Err.Description
End Function

' Função para escapar caracteres especiais no JSON
Function EscapeJSON(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, Chr(10), "\n")
    result = Replace(result, Chr(13), "\r")
    result = Replace(result, Chr(9), "\t")
    EscapeJSON = result
End Function

' Função para extrair mensagem de erro da resposta JSON
Function ExtrairMensagemErro(jsonResponse As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim errorMsg As String
    
    ' Procurar por "error":"mensagem"
    startPos = InStr(jsonResponse, """error"":""")
    If startPos > 0 Then
        startPos = startPos + 9 ' Pular 'error":"'
        endPos = InStr(startPos, jsonResponse, """")
        If endPos > startPos Then
            errorMsg = Mid(jsonResponse, startPos, endPos - startPos)
            ExtrairMensagemErro = errorMsg
            Exit Function
        End If
    End If
    
    ExtrairMensagemErro = "Erro desconhecido"
End Function

' Função para criar template de planilha
Sub CriarTemplatePlanilha()
    Dim ws As Worksheet
    Dim headers As Variant
    Dim i As Integer
    
    ' Criar nova planilha
    Set ws = Worksheets.Add
    ws.Name = "Template_Equipamentos"
    
    ' Definir cabeçalhos
    headers = Array("TAG", "STATUS", "MOTIVO", "PTS", "OS", "RETORNO", "CADEADO", "OBSERVACOES", "MODIFICADO_POR", "RESULTADO")
    
    ' Inserir cabeçalhos
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
        ws.Cells(1, i + 1).Font.Bold = True
        ws.Cells(1, i + 1).Interior.Color = RGB(200, 200, 200)
    Next i
    
    ' Adicionar exemplo de dados
    ws.Cells(2, 1).Value = "EQP001"
    ws.Cells(2, 2).Value = "OPE"
    ws.Cells(2, 3).Value = ""
    ws.Cells(2, 4).Value = ""
    ws.Cells(2, 5).Value = ""
    ws.Cells(2, 6).Value = ""
    ws.Cells(2, 7).Value = ""
    ws.Cells(2, 8).Value = "Exemplo de observação"
    ws.Cells(2, 9).Value = "Admin"
    
    ' Ajustar largura das colunas
    ws.Columns.AutoFit
    
    ' Adicionar validação de dados para STATUS
    With ws.Range("B:B").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="OPE,ST-BY,MANU"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    MsgBox "Template criado com sucesso na planilha '" & ws.Name & "'", vbInformation
End Sub

' Função para testar conexão com a API
Sub TestarConexaoAPI()
    Dim http As Object
    Dim response As String
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    On Error GoTo ErrorHandler
    
    ' Testar endpoint GET
    http.Open "GET", API_ENDPOINT, False
    http.send
    
    If http.Status = 200 Then
        MsgBox "Conexão com a API estabelecida com sucesso!", vbInformation
    Else
        MsgBox "Erro na conexão: HTTP " & http.Status, vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao conectar com a API: " & Err.Description, vbCritical
End Sub

