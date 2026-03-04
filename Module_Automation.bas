' Module de Automação Excel VBA
' Funções principais para processamento de dados

' Importar dados de API
Sub ImportarDados()
    Dim ws As Worksheet
    Dim url As String
    Dim httpRequest As Object
    
    Set ws = ThisWorkbook.Sheets("Dados")
    ws.Range("A:D").Clear
    
    url = "https://api.example.com/vendas"
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    With httpRequest
        .Open "GET", url, False
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        
        ' Processar resposta JSON
        Dim jsonResponse As String
        jsonResponse = .ResponseText
        
        ' Parse e insere dados
        ws.Range("A1").Value = "Status: Importação realizada"
    End With
End Sub

' Validar dados
Function ValidarDados() As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Dados")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value < 0 Then
            MsgBox "Erro: Valor negativo na linha " & i
            ValidarDados = False
            Exit Function
        End If
    Next i
    
    ValidarDados = True
    MsgBox "Validação concluída com sucesso!"
End Function

' Gerar relatório
Sub GerarRelatorio()
    Dim ws As Worksheet
    Dim sumValues As Double
    Dim countRows As Long
    
    Set ws = ThisWorkbook.Sheets("Dados")
    
    countRows = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1
    sumValues = Application.WorksheetFunction.Sum(ws.Range("B2:B" & countRows))
    
    ' Criar relatório
    ws.Range("F1").Value = "RELATÓRIO"
    ws.Range("F2").Value = "Total de linhas:"
    ws.Range("G2").Value = countRows
    ws.Range("F3").Value = "Soma total:"
    ws.Range("G3").Value = sumValues
    ws.Range("F4").Value = "Média:"
    ws.Range("G4").Value = sumValues / countRows
End Sub

' Filtrar dados
Sub FiltrarDados()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dados")
    
    With ws.Range("A1").CurrentRegion
        .AutoFilter
        .AutoFilter.ApplyCriterion Column:=2, Criteria1:=">100"
    End With
    
    MsgBox "Filtro aplicado!"
End Sub

' Limpar dados
Sub LimparDados()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dados")
    ws.Range("A:G").Clear
    MsgBox "Dados limpos!"
End Sub

' Função para copiar dados
Function CopiarDadosPara(nomeSheet As String) As Boolean
    On Error GoTo Erro
    
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    
    Set sourceSheet = ThisWorkbook.Sheets("Dados")
    Set targetSheet = ThisWorkbook.Sheets(nomeSheet)
    
    sourceSheet.UsedRange.Copy
    targetSheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    
    Application.CutCopyMode = False
    CopiarDadosPara = True
    Exit Function
    
Erro:
    CopiarDadosPara = False
    MsgBox "Erro ao copiar dados: " & Err.Description
End Function
