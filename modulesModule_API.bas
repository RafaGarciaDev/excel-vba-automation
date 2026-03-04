Attribute VB_Name = "Module_API"
Option Explicit

' ============================================================
' MODULE_API — Integração com APIs externas via HTTP
' ============================================================

Private Const BASE_URL As String = "https://api.exemplo.com/v1"
Private Const API_KEY  As String = "SUA_API_KEY_AQUI"

' ── Requisição GET genérica ─────────────────────────────────
Public Function HttpGet(endpoint As String) As String
    Dim http As Object
    Dim url  As String

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    url = BASE_URL & endpoint

    On Error GoTo ErrHandler
    http.Open "GET", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & API_KEY
    http.Send

    If http.Status = 200 Then
        HttpGet = http.responseText
    Else
        HttpGet = ""
        Module_Utils.LogError "HttpGet", http.Status, http.statusText
    End If

    Set http = Nothing
    Exit Function

ErrHandler:
    Module_Utils.LogError "HttpGet", Err.Number, Err.Description
    HttpGet = ""
End Function

' ── Requisição POST genérica ────────────────────────────────
Public Function HttpPost(endpoint As String, payload As String) As String
    Dim http As Object
    Dim url  As String

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    url = BASE_URL & endpoint

    On Error GoTo ErrHandler
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & API_KEY
    http.Send payload

    If http.Status = 200 Or http.Status = 201 Then
        HttpPost = http.responseText
    Else
        HttpPost = ""
        Module_Utils.LogError "HttpPost", http.Status, http.statusText
    End If

    Set http = Nothing
    Exit Function

ErrHandler:
    Module_Utils.LogError "HttpPost", Err.Number, Err.Description
    HttpPost = ""
End Function

' ── Extrai valor de um campo JSON simples (chave: "valor") ──
Public Function ParseJsonValue(json As String, key As String) As String
    Dim pattern As String
    Dim startPos As Long
    Dim endPos   As Long

    pattern = """" & key & """:"
    startPos = InStr(json, pattern)
    If startPos = 0 Then
        ParseJsonValue = ""
        Exit Function
    End If

    startPos = startPos + Len(pattern)

    ' Verifica se é string (começa com ")
    If Mid(json, startPos, 1) = """" Then
        startPos = startPos + 1
        endPos = InStr(startPos, json, """")
    Else
        endPos = InStr(startPos, json, ",")
        If endPos = 0 Then endPos = InStr(startPos, json, "}")
    End If

    ParseJsonValue = Mid(json, startPos, endPos - startPos)
End Function

' ── Busca cotação de moeda (exemplo: USD/BRL) ───────────────
Public Sub FetchExchangeRate(fromCurrency As String, toCurrency As String)
    Dim endpoint As String
    Dim response As String
    Dim rate     As String
    Dim ws       As Worksheet

    endpoint = "/exchange?from=" & fromCurrency & "&to=" & toCurrency
    response = HttpGet(endpoint)

    If response <> "" Then
        rate = ParseJsonValue(response, "rate")
        Set ws = Module_Utils.GetOrCreateSheet("API_Data")
        Dim nextRow As Long
        nextRow = Module_Utils.LastRow(ws) + 1
        ws.Cells(nextRow, 1).Value = Now()
        ws.Cells(nextRow, 2).Value = fromCurrency & "/" & toCurrency
        ws.Cells(nextRow, 3).Value = CDbl(rate)
        MsgBox "Cotação " & fromCurrency & "/" & toCurrency & ": " & rate, vbInformation, "API"
    Else
        MsgBox "Falha ao buscar cotação.", vbExclamation, "Erro API"
    End If
End Sub

' ── Envia dados de uma planilha para a API ──────────────────
Public Sub PostSheetData(wsName As String, endpoint As String)
    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim i       As Long
    Dim payload As String
    Dim response As String

    If Not Module_Utils.SheetExists(wsName) Then
        MsgBox "Planilha '" & wsName & "' não encontrada.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets(wsName)
    lastRow = Module_Utils.LastRow(ws)

    Module_Utils.SpeedOn
    Module_Utils.ShowStatus "Enviando dados para API..."

    Dim successCount As Integer
    successCount = 0

    For i = 2 To lastRow
        payload = "{""id"":""" & ws.Cells(i, 1).Value & """," & _
                  """value"":""" & ws.Cells(i, 2).Value & """," & _
                  """date"":""" & Format(ws.Cells(i, 3).Value, "YYYY-MM-DD") & """}"

        response = HttpPost(endpoint, payload)
        If response <> "" Then successCount = successCount + 1

        Module_Utils.ShowStatus "Enviando linha " & i & " de " & lastRow & "..."
    Next i

    Module_Utils.SpeedOff
    Module_Utils.ClearStatus
    MsgBox successCount & " de " & (lastRow - 1) & " registros enviados com sucesso.", _
           vbInformation, "Envio Concluído"
End Sub

' ── Importa dados da API para planilha ──────────────────────
Public Sub ImportApiData(endpoint As String, targetSheet As String)
    Dim response As String
    Dim ws       As Worksheet

    Module_Utils.ShowStatus "Buscando dados da API..."
    response = HttpGet(endpoint)

    If response = "" Then
        MsgBox "Sem dados retornados da API.", vbExclamation, "API"
        Module_Utils.ClearStatus
        Exit Sub
    End If

    Set ws = Module_Utils.GetOrCreateSheet(targetSheet)
    ws.Cells.ClearContents

    ' Cabeçalhos
    ws.Cells(1, 1).Value = "Timestamp"
    ws.Cells(1, 2).Value = "Dados Brutos"

    ws.Cells(2, 1).Value = Now()
    ws.Cells(2, 2).Value = response

    Module_Utils.ClearStatus
    MsgBox "Dados importados para a aba '" & targetSheet & "'.", vbInformation, "Sucesso"
End Sub