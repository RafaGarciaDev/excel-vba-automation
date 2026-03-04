Attribute VB_Name = "Module_Utils"
Option Explicit

' ============================================================
' MODULE_UTILS — Funções utilitárias reutilizáveis
' ============================================================

' ── Desativa eventos/tela para performance ──────────────────
Public Sub SpeedOn()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
End Sub

Public Sub SpeedOff()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
End Sub

' ── Limpa um intervalo com segurança ────────────────────────
Public Sub ClearRange(ws As Worksheet, rng As String)
    On Error Resume Next
    ws.Range(rng).ClearContents
    On Error GoTo 0
End Sub

' ── Verifica se uma planilha existe ─────────────────────────
Public Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' ── Cria planilha se não existir ────────────────────────────
Public Function GetOrCreateSheet(sheetName As String) As Worksheet
    If SheetExists(sheetName) Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    Else
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

' ── Última linha preenchida de uma coluna ───────────────────
Public Function LastRow(ws As Worksheet, Optional col As Integer = 1) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' ── Última coluna preenchida de uma linha ───────────────────
Public Function LastCol(ws As Worksheet, Optional row As Integer = 1) As Long
    LastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
End Function

' ── Formata um intervalo como tabela ────────────────────────
Public Sub FormatAsTable(ws As Worksheet, rng As Range, tableName As String)
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = tableName
        tbl.TableStyle = "TableStyleMedium2"
    End If
End Sub

' ── Exibe mensagem de status na barra do Excel ──────────────
Public Sub ShowStatus(msg As String)
    Application.StatusBar = msg
    DoEvents
End Sub

Public Sub ClearStatus()
    Application.StatusBar = False
End Sub

' ── Converte coluna numérica para letra (1 = A, 2 = B...) ──
Public Function ColLetter(colNum As Integer) As String
    ColLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

' ── Log de erros em planilha "Log" ──────────────────────────
Public Sub LogError(source As String, errNum As Long, errDesc As String)
    Dim wsLog As Worksheet
    Dim nextRow As Long

    Set wsLog = GetOrCreateSheet("Log")
    nextRow = LastRow(wsLog) + 1

    wsLog.Cells(nextRow, 1).Value = Now()
    wsLog.Cells(nextRow, 2).Value = source
    wsLog.Cells(nextRow, 3).Value = errNum
    wsLog.Cells(nextRow, 4).Value = errDesc

    If nextRow = 2 Then
        wsLog.Cells(1, 1).Value = "Timestamp"
        wsLog.Cells(1, 2).Value = "Source"
        wsLog.Cells(1, 3).Value = "ErrNum"
        wsLog.Cells(1, 4).Value = "Description"
    End If
End Sub

' ── Salva backup com timestamp ───────────────────────────────
Public Sub SaveBackup()
    Dim backupName As String
    backupName = ThisWorkbook.Path & "\backup_" & Format(Now(), "YYYYMMDD_HHMMSS") & ".xlsm"
    ThisWorkbook.SaveCopyAs backupName
    MsgBox "Backup salvo em:" & vbNewLine & backupName, vbInformation, "Backup"
End Sub

' ── Copia intervalo para outra planilha ─────────────────────
Public Sub CopyRangeToSheet(srcWs As Worksheet, srcRng As String, _
                             dstWs As Worksheet, dstCell As String)
    srcWs.Range(srcRng).Copy
    dstWs.Range(dstCell).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End Sub

' ── Validação: campo obrigatório não vazio ──────────────────
Public Function IsNotEmpty(value As Variant) As Boolean
    IsNotEmpty = Not (IsNull(value) Or Trim(CStr(value)) = "")
End Function

' ── Formata número como moeda BRL ───────────────────────────
Public Function ToBRL(value As Double) As String
    ToBRL = "R$ " & Format(value, "#,##0.00")
End Function