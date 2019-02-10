Attribute VB_Name = "modExtract"
' modExtract
' Extraction functions

' Populates the Collect data sheet, from the template
' saves it in a new excel file
' Input data:
'   scenario name: String - name of scenario to extract
'   portfolios: list of portfolios to include in the result file
'   iMonthIni, iMonthFin: Integer, staring months to report
'   nMonths, number of months to report, max = 12


Option Compare Database
Sub modExtract_populateDataSheet(scenario As String, portfolios() As String, _
                                iMonthIni As Long, nMonths As Long)

    Dim ok As Boolean
    ' First data row in TabDati
    Const iFirstDataRow = 4
    ' First month col in Tabdati (F)
    Const iFirstDataCol = 6
    
     ' check input
    Dim year As Long
    Dim mon As Long

    Debug.Print "Processing " & scenario
    If UBound(portfolios) - LBound(portfolios) + 1 = 0 Then
        Err.Raise 513, Description:="Bad input: zero length array"
        Exit Sub
    End If
    year = iMonthIni / 100
    mon = iMonthIni - year * 100
    If year < 2010 Then
        Err.Raise 513, Description:="Bad input: year too old"
        Exit Sub
    End If
    If mon < 1 Or mon > 12 Then
        Err.Raise 513, Description:="Bad input: bad month number"
        Exit Sub
    End If
    If nMonths < 1 Or nMonths > 12 Then
        Err.Raise 513, Description:="Bad input: bad number of months (1 to 12)"
        Exit Sub
    End If

         
     ' open excel template
    ok = modExcel_OpenExcel(True)
    If Not ok Then
        Stop
    End If
            
    ok = modExcel_OpenWorkBook(CurrentProject.Path & "\Template Foglio raccolta dati.xlsm")
    If Not ok Then
        Stop
    End If
    
    ' Populate TabTasks
    ok = modExcel_SetWorkSheet("Iniziative")
    If Not ok Then
        Stop
    End If
    
    Dim rstIniziativa As Recordset
    Dim iRow As Long
    Dim qdf As QueryDef
    Set qdf = CurrentDb.QueryDefs("qry_ScenarioDiPiano")
    qdf!scenario = scenario
    iRow = 2
    Set rstIniziativa = qdf.OpenRecordset
    Do While Not rstIniziativa.EOF
        If Not IsNull(rstIniziativa!codPortfolio) Then
            If IsInArray(rstIniziativa!codPortfolio, portfolios) Then
                ok = modExcel_WriteCell("A", iRow, rstIniziativa!codTask)
                ok = modExcel_WriteCell("B", iRow, rstIniziativa!descProgetto & " - " & rstIniziativa!descTask)
                ok = modExcel_WriteCell("D", iRow, rstIniziativa!dtInizio)
                ok = modExcel_WriteCell("E", iRow, rstIniziativa!dtFine)
                ' Debug.Print rstIniziativa!descProgetto & " - " & rstIniziativa!descTask
                If Not ok Then
                    Exit Do
                End If
                iRow = iRow + 1
            End If
        End If
        rstIniziativa.MoveNext
    Loop
    rstIniziativa.Close
    qdf.Close
    Set rstIniziativa = Nothing
    Set qdf = Nothing
    If Not ok Then
        Stop
    End If
    
    ' Populate TabRisorse
    ok = modExcel_SetWorkSheet("Personale DAPI")
    If Not ok Then
        Stop
    End If
    Dim rstRisorse As Recordset
    Dim sQueryRes As String
    sQueryRes = "select nome from Risorsa order by nome"
    Set rstRisorse = CurrentDb.OpenRecordset(sQueryRes, dbOpenDynaset)
    rstRisorse.MoveFirst
    iRow = 2
    Do While Not rstRisorse.EOF
        ok = modExcel_WriteCell("A", iRow, rstRisorse!Nome)
        rstRisorse.MoveNext
        iRow = iRow + 1
    Loop
    rstRisorse.Close
    Set rstRisorse = Nothing
    
    ' Populate dati sheet
    ' create a cursor on  qry_TabelloneAllocazioniXScenario
    ok = modExcel_SetWorkSheet("Dati")
    If Not ok Then
        Stop
    End If

    Dim rstQry As Recordset
    Dim prevTask As String, curTask As String
    Dim prevRes As String, curRes As String
    Dim iMonthFin As Long
    Dim iMonth As Long
    Dim sPortfolio As String
    Dim dblPct As Double
    iMonthFin = iMonthIni + nMonths - 1
    
    ' at row 3, from "F" onward, write date 1/month/year
    Dim i As Long
    Dim iCol As Integer
    For i = iMonthIni To iMonthFin
        Dim d As Date
        
        iCol = iFirstDataCol + i - iMonthIni
        year = i / 100
        mon = i - year * 100
        d = DateSerial(year, mon, 1)
        ok = modExcel_WriteCell(Chr(64 + iCol), 3, d)
        If Not ok Then
            Exit Sub
        End If
    Next i
    

    Set qdf = CurrentDb.QueryDefs("qry_TabelloneAllocazioniXScenario")
    qdf!scenario = scenario
    Set rstQry = qdf.OpenRecordset(dbOpenDynaset)
    rstQry.MoveFirst
    prevTask = ""
    prevRes = ""
    iRow = iFirstDataRow - 1
    
    ' loop through records
    Do While Not rstQry.EOF
        curTask = rstQry![Codice task]
        curRes = rstQry!Nome
        iMonth = rstQry!mese
        dblPct = rstQry!pct / 100
        sPortfolio = Nz(rstQry!codPortfolio, "   ")
        
        ' portfolio is ok and month is the right interval ?
        If IsInArray(sPortfolio, portfolios) And (iMonth >= iMonthIni And iMonth <= iMonthFin) Then
            If curTask <> prevTask Or curRes <> prevRes Then
                ' different task AND/OR different resource? change row
                iRow = iRow + 1
                ok = modExcel_WriteCell("A", iRow, curTask)
                If ok Then
                    ok = modExcel_WriteCell("C", iRow, curRes)
                End If
                If Not ok Then
                    Exit Do
                End If
            End If
            ' write month on the correct column (col iFirstDataCol + month - initial month)
            
            iCol = iFirstDataCol + iMonth - iMonthIni
            ok = modExcel_WriteCell(Chr(64 + iCol), iRow, dblPct)
            If Not ok Then
                Exit Do
            End If
            prevTask = curTask
            prevRes = curRes
        End If
        rstQry.MoveNext
    Loop

    rstQry.Close
    qdf.Close
    Set qdf = Nothing
    Set rstQry = Nothing
    Debug.Print "modExtract_populateDataSheet: exit"
End Sub

Private Sub test1()
    Dim a(1) As String
     ' a(0) = "BIR"
     ' a(1) = "SDP"
    modExtract_populateDataSheet "4Q2018", a, 201910, 1
End Sub
