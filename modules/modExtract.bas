' modExtract
' Extraction functions

' Populates the Collect data sheet, from the template
' saves it in a new excel file
' Input data:
'   scenario name: String - name of scenario to extract
'   portfolios: list of portfolios to include in the result file
'   year, month: Integer, starting months to report
'   nMonths, number of months to report, 1 to 12


Option Compare Database
Function modExtract_populateDataSheet(scenario As String, portfolios() As String, _
                                year As Integer, month As Integer, nMonths As Integer) As Boolean

    On Error GoTo ErrorHandler
    
    Dim ok As Boolean
    ' First data row in TabDati
    Const iFirstDataRow = 4
    ' First month col in Tabdati (F)
    Const iFirstDataCol = 6
    
     ' check input
    Dim iMonthIni As Long
    Dim i As Long
    
    Debug.Print "Processing " & scenario
    If UBound(portfolios) - LBound(portfolios) + 1 = 0 Then
        Err.Raise 513, Description:="Bad input: zero length array"
    End If
    If year < 2010 Then
        Err.Raise 513, Description:="Bad input: year too old"
    End If
    If month < 1 Or month > 12 Then
        Err.Raise 513, Description:="Bad input: bad month number"
    End If
    If nMonths < 1 Or nMonths > 12 Then
        Err.Raise 513, Description:="Bad input: bad number of months (1 to 12)"
    End If
    ' set month ini as YYYYMM numeric
    iMonthIni = CLng(year) * 100 + month
    Dim iMonthFinal As Integer, yearFinal As Integer
    Dim iMonthFin As Long
    iMonthFinal = (month + nMonths - 2) Mod 12 + 1
    yearFinal = IIf(month + nMonths - 1 > 12, year + 1, year)
    iMonthFin = iMonthFinal + CLng(yearFinal) * 100
     
     ' open excel template
    ok = modExcel_OpenExcel(True, True)
    If Not ok Then
        Err.Raise 514, Description:="Error opening excel"
    End If
            
    ok = modExcel_OpenWorkBook(CurrentProject.Path & "\Template.xlsm")
    If Not ok Then
        Err.Raise 515, Description:="template foglio raccolta dati: possibile file xls mancante in " & _
            CurrentProject.Path & "\Template Foglio raccolta dati.xlsm"
    End If
    
    ' --- Set metadata values in Meta sheet
    ok = modExcel_SetWorkSheet("meta")
    If Not ok Then
        Err.Raise 515, Description:="Template excel errato - Manca foglio Meta"
    End If
    ok = modExcel_WriteCell("B", 1, scenario)
    ok = modExcel_WriteCell("B", 2, iMonthIni)
    ok = modExcel_WriteCell("B", 3, iMonthFin)
    For i = LBound(portfolios) To UBound(portfolios)
        ok = modExcel_WriteCell(Chr(66 + i - LBound(portfolios)), 4, portfolios(i))
    Next i
    
    ' Populate TabTasks
    ok = modExcel_SetWorkSheet("Iniziative")
    If Not ok Then
        Err.Raise 515, Description:="Template excel errato - Manca foglio Iniziative"
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
        Err.Raise 514, Description:="Errore in scrittura su file excel"
    End If
    
    ' Populate TabRisorse
    ok = modExcel_SetWorkSheet("Personale DAPI")
    If Not ok Then
        Err.Raise 515, Description:="Template excel errato - Manca foglio Personale DAPI"
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
    ok = modExcel_SetWorkSheet("Dati")
    If Not ok Then
        Err.Raise 515, Description:="Template excel errato - Manca foglio Dati"
    End If

    Dim rstQry As Recordset
    Dim prevTask As String, curTask As String
    Dim prevRes As String, curRes As String
    Dim iMonth As Long
    Dim sPortfolio As String
    Dim dblPct As Double
    iMonthFin = iMonthIni + nMonths - 1
    
    ' at row 3, from "F" onward, write date 1/month/year
    
    Dim iCol As Integer
    Dim iYear As Long
    Dim iThisMon As Long
    For i = 0 To 11
        Dim d As Date
        
        iCol = iFirstDataCol + i
        iYear = IIf(month + i > 12, year + 1, year)
        ' iMonthFinal = (month + nMonths - 2) Mod 12 + 1
        iThisMon = (month - 1 + i) Mod 12 + 1
        d = DateSerial(iYear, iThisMon, 1)
        ok = modExcel_WriteCell(Chr(64 + iCol), 3, d)
        If Not ok Then
            Err.Raise 514, Description:="Errore in scrittura su file excel"
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
                Err.Raise 514, Description:="Errore in scrittura su file excel"
            End If
            prevTask = curTask
            prevRes = curRes
        End If
        rstQry.MoveNext
    Loop

    ' run macro to protect sheets
    ok = modExcel_RunMacro("templateDataCollection_ProtectSheets")
    If Not ok Then
        Err.Raise 514, Description:="Errore in esecuzione macro Protect su file excel"
    End If

    rstQry.Close
    qdf.Close
    Set qdf = Nothing
    Set rstQry = Nothing
    modExtract_populateDataSheet = True
    Exit Function

ErrorHandler:
    Select Case ErrorLogging(Err.Number, Err.Description, "modExtract_populateDataSheet")
    Case 1: Resume
    Case 2: Resume Next
    Case Else:
        If Not rstQry Is Nothing Then
            rstQry.Close
            Set rstQry = Nothing
        End If
        If Not qdf Is Nothing Then
            qdf.Close
            Set qdf = Nothing
        End If
        modExtract_populateDataSheet = False
    End Select
    Exit Function
End Function

Private Sub test1()
    Dim a(1) As String
     a(0) = "IVASS"
     a(1) = "SDP"
    modExtract_populateDataSheet "4Q2018", a, 2019, 2, 11
End Sub
Private Sub test2()
   Dim a(1) As String
     a(0) = "IVASS"
     a(1) = "SDP"
    Dim ret As Boolean
    ret = modExtract_populateDataSheet("4Q2018", a, 2019, 1, 11)
    Debug.Assert (Not ret)
    Debug.Print "test2 ok"
End Sub