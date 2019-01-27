Attribute VB_Name = "templateDataCollection"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'
Dim rng As Range, cell As Range
Dim i As Integer

Sheets("Dest").Select
Columns("A:D").Select
Selection.ClearContents
Range("A1").Value = "Iniziativa"
Range("B1").Value = "Risorsa"
Range("C1").Value = "Pct"
Range("D1").Value = "Mese"
Sheets(1).Select
Set rng = Range("A2:A200")
i = 2
For Each cell In rng
    For j = 1 To 12
        If cell.Offset(0, j + 1) > 0 Then
            Worksheets("Dest").Cells(i, 1) = cell.Value
            Worksheets("Dest").Cells(i, 2) = cell.Offset(0, 1)
            Worksheets("Dest").Cells(i, 3) = cell.Offset(0, j + 1)
            Worksheets("Dest").Cells(i, 4) = Cells(1, j + 2)
            i = i + 1
        End If
    Next j
Next cell
  
  
End Sub

Sub Modulo1_formatConditional()
    Dim c As Range
    Dim nRiga As Integer, nCol As Integer
    Dim colName As String
    
    For Each c In Range("C2:N72")
        c.Select
        ' =E($P2<=C$1;$Q2>=C$1)
        nRiga = c.Row
        nCol = c.Column
        colName = Chr$(nCol + 64)
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
    "=E($P" & nRiga & "<=" & colName & "$1;$Q" & nRiga & ">=" & colName & "$1)"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.399945066682943
        End With
        Selection.FormatConditions(1).StopIfTrue = False
    Next c

    
 

End Sub
Function getColAsText(addr As String) As String
    Dim pos As Integer
    pos = InStr(2, addr, "$")
    getColAsText = UCase(Mid(addr, 2, pos - 2))
End Function

Function getRow(addr As String) As Integer
    Dim pos As Integer
    pos = InStr(2, addr, "$")
    getRow = Mid(addr, pos + 1, Len(addr) - pos)
End Function

Sub templateDataCollection_openForm()
    frmAddRow.Show
End Sub
Sub templateDataCollection_deleteRow()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TabDati")
    Dim newrow As ListRow
    ActiveSheet.Unprotect
    Dim r As Range
    Dim sRRange As String
    
    Set r = Application.Selection
    If r.Rows.Count > 1 Then
        ok = MsgBox("Vuoi eliminare " & r.Rows.Count & " righe ? ", vbOKCancel, "Conferma")
        If ok <> vbOK Then
            Exit Sub
        End If
    End If
    
    
    ' Range address of selected rows
    
    r.Rows.Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Protect
End Sub
Sub templateDataCollection_addRow(sIDTask As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("TabDati")
    Dim newrow As ListRow
    ActiveSheet.Unprotect
    Set newrow = tbl.ListRows.Add
    With newrow
        .Range(1) = sIDTask
    End With
    ActiveSheet.Protect
End Sub
