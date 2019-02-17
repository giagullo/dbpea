Attribute VB_Name = "modExcel"
Option Compare Database
Option Explicit
' ********************************************************************************************************
' Microsoft Access 2010 VBA Programming Inside Out
' Database Name         : ExcelAnalysis
'                       : Generates analysis in MS Excel
' Module Name           : modExcel
' Module Author         : Andrew Couch
' Module Version        : 1.0
' Module Revisions      :
' Module Description    :
'
' Copyright
' ---------
' You may add this code to your own applications without any acknowledgment of the source of the material
' You do not have the rights to make this code available on the internet without permission
' Copyright © Andrew Couch 2011, All rights Reserved
' ********************************************************************************************************

Dim appExcel As Excel.Application
Dim wkbExcel As Excel.Workbook
Dim wksExcel As Excel.Worksheet

Function modExcel_OpenExcel(UseExisting As Boolean) As Boolean
' open a copy of Excel
    If UseExisting Then
        On Error Resume Next
        Set appExcel = GetObject(, "Excel.Application")
        appExcel.Visible = True
        Err.Clear
        
    End If
    If appExcel Is Nothing Then
        Set appExcel = CreateObject("Excel.Application")
        appExcel.Visible = True
    End If
    If Err <> 0 Then
        MsgBox "An error occured trying to start MS Excel : " & _
                Err.Description, vbCritical, "Unable To Start MS Excel"
        modExcel_OpenExcel = False
    Else
        modExcel_OpenExcel = True
    End If
    On Error GoTo 0
End Function

Function modExcel_OpenWorkBook(strFileName As String) As Boolean
' Open an existing Excel document
    On Error Resume Next
    Set wkbExcel = appExcel.Workbooks.Open(strFileName)
    If Err <> 0 Then
        MsgBox "An error occured trying to start open the file : " & _
                strFileName & " : " & Err.Description, vbCritical, _
                "Unable To Open Workbook"
        modExcel_OpenWorkBook = False
    Else
        modExcel_OpenWorkBook = True
    End If
    On Error GoTo 0
End Function

Function modExcel_CreateNewWorkBook() As Boolean
' Open an existing word document
    On Error Resume Next
    Set wkbExcel = appExcel.Workbooks.Add
    If Err <> 0 Then
        MsgBox "An error occured trying to add new workbook : " & Err.Description, vbCritical, "Unable To Create Workbook"
        modExcel_CreateNewWorkBook = False
    Else
        modExcel_CreateNewWorkBook = True
    End If
    On Error GoTo 0
End Function
Function modExcel_OpenActiveWorkSheet() As Boolean
' Open an existing word document
    On Error Resume Next
    Set wksExcel = wkbExcel.ActiveSheet
    If Err <> 0 Then
        MsgBox "An error occured trying to choose active sheet in workbook :" _
                & Err.Description, vbCritical, "Unable To Select ActiveSheet"
        modExcel_OpenActiveWorkSheet = False
    Else
        modExcel_OpenActiveWorkSheet = True
    End If
    On Error GoTo 0
End Function
Function modExcel_SetWorkSheet(name As String) As Boolean
' Open a worksheet
    On Error Resume Next
    Set wksExcel = wkbExcel.Worksheets(name)
    If Err <> 0 Then
        MsgBox "An error occured trying to select active sheet in workbook :" _
                & Err.Description, vbCritical, "Unable To Select ActiveSheet"
        modExcel_SetWorkSheet = False
    Else
        modExcel_SetWorkSheet = True
    End If
    On Error GoTo 0
End Function

Function modExcel_WriteCell(strCol As String, lngRow As Long, _
                            strCellvalue As Variant) As Boolean
' Write to a cell
    On Error Resume Next
    wksExcel.Cells(lngRow, modExcel_MapLetterToColumn(strCol)) = strCellvalue
    If Err <> 0 Then
        MsgBox "An error occured trying to write to excel : " & Err.Description, _
            vbCritical, "Unable To Write Data"
        modExcel_WriteCell = False
    Else
        modExcel_WriteCell = True
    End If
    On Error GoTo 0
End Function
Function modExcel_ReadCell(strCol As String, lngRow As Long, _
                            ByRef strCellvalue As String) As Boolean
' Read from a cell
' Note we are ensuring that by using an Explicit ByVal that this routine can modify the value
    On Error Resume Next
    strCellvalue = wksExcel.Cells(lngRow, modExcel_MapLetterToColumn(strCol))
    If Err <> 0 Then
        MsgBox "An error occured trying to read from excel : " _
                & Err.Description, vbCritical, "Unable To Read Data"
        modExcel_ReadCell = False
    Else
        modExcel_ReadCell = True
    End If
    On Error GoTo 0
End Function

Function modExcel_CleanUp(boolCloseWorkSheet As Boolean, _
                            boolCloseExcel As Boolean) As Boolean
    If boolCloseWorkSheet Then
        wkbExcel.Close
    End If
    If boolCloseExcel Then
        appExcel.Quit
    End If
    Set appExcel = Nothing
    Set wkbExcel = Nothing
    Set wksExcel = Nothing
    If Err <> 0 Then
        modExcel_CleanUp = False
    Else
        modExcel_CleanUp = True
    End If
End Function

Function modExcel_SaveAs(strNewDocName As String) As Boolean
    wkbExcel.SaveAs strNewDocName
    If Err <> 0 Then
        modExcel_SaveAs = False
    Else
        modExcel_SaveAs = True
    End If
End Function

Function modExcel_MapLetterToColumn(strCol As String) As Long
    ' map excel columns to a number
    ' A......Z AA.....ZZ
    Dim lngCol As Long
    Dim strChar As String
    If Len(strCol) = 0 Or Len(strCol) > 2 Then
        modExcel_MapLetterToColumn = 0
        Exit Function
    End If
    strChar = UCase(Left(strCol, 1))
    lngCol = Asc(strChar) - 64
    If Len(strCol) > 1 Then
        strChar = UCase(Mid(strCol, 1, 1))
        lngCol = (lngCol + 1) * 25 + (Asc(strChar) - 64)
    End If
    modExcel_MapLetterToColumn = lngCol
End Function

Sub modExcel_SyntaxtForWorkSheets(strExpectedName As String)
    'Set wksExcel = wkbExcel.Sheets(i)
    Dim sheetExcel As Excel.Worksheet
    ' use code like the following to locate a specific worksheet
    For Each sheetExcel In wkbExcel.Worksheets
        If sheetExcel.name = strExpectedName Then
            Set wksExcel = sheetExcel
        End If
    Next
End Sub

Function modExcel_RefreshConnections() As Boolean
    On Error GoTo Err_Handler
    Dim lngCount As Long
'    wkbExcel.RefreshAll
    ' refresh all connections in the workbook
    For lngCount = 1 To wkbExcel.Connections.Count
        wkbExcel.Connections(lngCount).Refresh
    Next
    modExcel_RefreshConnections = True
    Exit Function
Err_Handler:
    modExcel_RefreshConnections = False
    Exit Function
End Function

Function modExcel_RefreshTablesAndPivots() As Boolean
    On Error GoTo Err_Handler
    Dim lngCount As Long
    Dim pvt As Excel.PivotTable
    Dim qtbl As Excel.QueryTable
'    wkbExcel.RefreshAll
    ' refresh any pivot tables
    For Each pvt In wksExcel.PivotTables
        pvt.RefreshTable
    Next
    ' refresh any query tables
    For Each qtbl In wksExcel.QueryTables
        qtbl.RefreshTable
    Next
    modExcel_RefreshTablesAndPivots = True
    Exit Function
Err_Handler:
    modExcel_RefreshTablesAndPivots = False
    Exit Function
End Function

