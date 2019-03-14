Option Compare Database
' Input from excel workbook allocation data for a scenario & PF & months range
Type MetaDataScenario
    scenario As String
    portfolios() As String
    iniMonth As Long
    endMonth As Long
End Type

' Clone a scenario from an existing one
' Scenario must be registered in Scenario table
Function modScenario_Clone(origin As String, newOne As String) As Boolean

    On Error GoTo ErrorHandler
    
    Dim db As Database
    Set db = CurrentDb
    Dim qInsert As String
    
    qInsert = "INSERT INTO PianoTask ( IDTask, dtInizio, dtFine, scenario )" & _
      "SELECT PianoTask.IDTask, PianoTask.dtInizio, PianoTask.dtFine, '" & newOne & "'" & _
      " FROM PianoTask" & _
      " WHERE (((PianoTask.scenario)='" & origin & "'));"
    db.Execute qInsert, dbFailOnError
    
    qInsert = "INSERT INTO Allocazioni  (IDTask, scenario, IDRisorsa, mese, pct)" & _
      "SELECT Allocazioni.IDTask,'" & newOne & "', Allocazioni.IDRisorsa, Allocazioni.mese,Allocazioni.pct " & _
      "FROM Allocazioni  WHERE (((Allocazioni.scenario)='" & origin & "'));"
    db.Execute qInsert, dbFailOnError
 
    db.Close
    Set db = Nothing
    modScenario_Clone = True
    Exit Function
ErrorHandler:
    Select Case ErrorLogging(Err.Number, Err.Description, "modScenario_Clone")
    Case 1: Resume
    Case 2: Resume Next
    Case Else:
        If Not db Is Nothing Then
            db.Close
            Set db = Nothing
        End If
        modScenario_Clone = False
    End Select
    Exit Function
End Function



Function modScenario_GetAllocation(fileName As String, dry As Boolean, ByRef infos As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim metaData As MetaDataScenario
    ' Dim infos As String
   
    ' open excel file
    ok = modExcel_OpenExcel(True, True)
    If Not ok Then
        Err.Raise 514, Description:="Error opening excel"
    End If
            
    ok = modExcel_OpenWorkBook(fileName)
    If Not ok Then
        Err.Raise 515, Description:="errore in apertura file excel " & _
            fileName
    End If
    ' read Metadata from excel
    ok = modExcel_SetWorkSheet("Meta")
    If Not ok Then
        Err.Raise 515, Description:="File excel errato - Manca foglio Meta"
    End If

    ok = modSScenario_ReadMetadata(metaData, infos)
    If Not ok Then
        Err.Raise 515, "File excel errato: formato foglio Meta: " & infos
    End If
    
    If Not dry Then
        DAO.BeginTrans
    End If
    If Not dry Then
        ok = modScenario_DeletePrev(metaData)
        If Not ok Then
            modScenario_GetAllocation = False
            Exit Function
        End If
    End If
    
    ' cycle through rows from 4 to the end of "Dati" sheet
    ok = modExcel_SetWorkSheet("Dati")
    If Not ok Then
        Err.Raise 515, Description:="File excel errato - Manca foglio Dati"
    End If
    Dim i As Long, j As Long
    Dim taskSipros As String
    Dim resName As String
    Dim lngMonth As Long
    Dim dblAlloc As Double
    Dim strCell As String
    Dim nInserted As Long
    Dim nMonths As Double
    
    nInserted = 0
    nMonths = 0
    i = 4
    ok = modExcel_ReadCell("A", i, taskSipros)
    If Not ok Then
        Err.Raise 514, Description:="Errore in lettura su file excel riga " & i
    End If
    Do While taskSipros <> ""
        Debug.Print taskSipros
        ok = modExcel_ReadCell("C", i, resName)
        If Not ok Then
            Err.Raise 514, Description:="Errore in lettura su file excel riga " & i
        End If
        ' Cols "F" to "Q"
        For j = 6 To 17
            ' month num in row 2
            ok = modExcel_ReadCell(Chr(64 + j), 2, strCell)
            lngMonth = CLng(strCell)
            If lngMonth >= metaData.iniMonth And lngMonth <= metaData.endMonth Then
                ' (double safely!) checked in range
                ok = modExcel_ReadCell(Chr(64 + j), i, strCell)
                If Not ok Then
                    Err.Raise 514, Description:="Errore in lettura su file excel riga " & i
                End If
                ' ready to insert Allocazioni record
                If strCell <> "" Then
                    dblAlloc = CDbl(Nz(strCell, 0))
                    If Not dry Then
                        ok = modScenario_InsertAlloc(taskSipros, resName, lngMonth, dblAlloc, metaData.scenario)
                    End If
                    If Not ok Then
                        modScenario_GetAllocation = False
                        Exit Function
                    End If
                    nInserted = nInserted + 1
                    nMonths = nMonths + dblAlloc
                End If
            End If
        Next j
        ' read data for the next row
        i = i + 1
        ok = modExcel_ReadCell("A", i, taskSipros)
        If Not ok Then
            Err.Raise 514, Description:="Errore in lettura su file excel riga " & i
        End If
    Loop
    If Not dry Then
        DAO.CommitTrans
    End If
    ' cleanup
    ok = modExcel_CleanUp(True, False)
    If Not ok Then
        Err.Raise 514, Description:="Errore chiusura cartella excel"
    End If
    ' build infos for caller
    With metaData
        infos = IIf(dry, "CARICAMENTO SIMULATO, NESSUN MODIFICA ESEGUITA", "ESEGUITO")
        infos = infos & vbCrLf & "Scenario: " & .scenario
        infos = infos & vbCrLf & "Da: " & .iniMonth
        infos = infos & vbCrLf & "A : " & .endMonth
        infos = infos & vbCrLf & "Portafogli: " & Join(.portfolios, ";") & vbCrLf
        infos = infos & vbCrLf & "Inserito " & nInserted & " record di allocazioni." & vbCrLf & "Totale " & nMonths & " mesi persona"
    End With
    ' exit with ok
    modScenario_GetAllocation = True
    Exit Function
ErrorHandler:
    DAO.Rollback
    Select Case ErrorLogging(Err.Number, Err.Description, "modScenario_GetAllocation")
    Case 1: Resume
    Case 2: Resume Next
    Case Else:
        modScenario_GetAllocation = False
    End Select
    Exit Function
End Function
Private Function modScenario_DeletePrev(m As MetaDataScenario) As Boolean
    On Error GoTo ErrorHandler
    Dim pfList As String
    pfList = ""
    For i = 0 To UBound(m.portfolios)
        pfList = pfList & "'" & m.portfolios(i) & "'"
        If i < UBound(m.portfolios) Then
            pfList = pfList & ","
        End If
    Next i
    Dim s As String
    s = "DELETE FROM Allocazioni WHERE scenario = '" & m.scenario & "'"
    s = s & " AND mese BETWEEN " & m.iniMonth & " AND " & m.endMonth
    s = s & " AND IDtask IN (SELECT ID FROM Task WHERE codPortfolio IN (" & pfList & "))"
    Dim db As Database
    Set db = CurrentDb
    Debug.Print "Executing: " & s
    db.Execute s, dbFailOnError
    db.Close
    Set db = Nothing
    ' exit ok
    modScenario_DeletePrev = True
    Exit Function
ErrorHandler:
    Select Case ErrorLogging(Err.Number, Err.Description, "modScenario_DeletePrev")
    Case 1: Resume
    Case 2: Resume Next
    Case Else:
        If Not db Is Nothing Then
            db.Close
            Set db = Nothing
        End If
        modScenario_DeletePrev = False
    End Select
End Function
Private Function modScenario_InsertAlloc _
    (t As String, res As String, mon As Long, alloc As Double, scenario As String) _
    As Boolean
    
    On Error GoTo ErrorHandler
    Dim db As Database
    Dim s As String
    ' lookup Task ID
    Dim taskID As Long
    taskID = DLookup("ID", "Task", "codSIPROS = '" & t & "'")
    Dim resID As Long
    resID = DLookup("ID", "Risorsa", "Nome = '" & doubleApex(res) & "'")
    Debug.Print "Insert: " & t & "(" & taskID & ")", res & "(" & resID & ")", mon, alloc
    s = "INSERT INTO Allocazioni (IDTask, scenario, IDRisorsa, mese, pct) VALUES (" & _
        taskID & ",'" & scenario & "'," & resID & "," & mon & "," & alloc * 100 & ")"
    ' Debug.Print s
    Set db = CurrentDb
    db.Execute s, dbFailOnError
    
    ' exit ok
    db.Close
    Set db = Nothing
    modScenario_InsertAlloc = True
    Exit Function
ErrorHandler:
    Select Case ErrorLogging(Err.Number, Err.Description, "modScenario_InsertAlloc")
    Case 1: Resume
    Case 2: Resume Next
    Case Else:
        If Not db Is Nothing Then
            db.Close
            Set db = Nothing
        End If
        modScenario_InsertAlloc = False
    End Select

End Function
Private Function modSScenario_ReadMetadata(ByRef m As MetaDataScenario, _
            ByRef infos As String) As Boolean
            Dim strVal As String
            With m
                ok = modExcel_ReadCell("B", 1, strVal)
                If Not ok Then
                    infos = "Scenario"
                    modSScenario_ReadMetadata = False
                    Exit Function
                End If
                .scenario = strVal
                
                ok = modExcel_ReadCell("B", 2, strVal)
                If Not ok Then
                    infos = "Mese iniziale"
                    modSScenario_ReadMetadata = False
                    Exit Function
                End If
                .iniMonth = CLng(strVal)
                
                ok = modExcel_ReadCell("B", 3, strVal)
                If Not ok Then
                    infos = "Mese finale"
                    modSScenario_ReadMetadata = False
                    Exit Function
                End If
                .endMonth = CLng(strVal)
                
                Dim i As Integer
                i = 0
                ReDim m.portfolios(12)
                Do
                    ok = modExcel_ReadCell(Chr(Asc("B") + i), 4, strVal)
                    If Not ok Then
                        infos = "Portfolio"
                        modSScenario_ReadMetadata = False
                        Exit Function
                    End If
                    If strVal <> "" Then
                        m.portfolios(i) = strVal
                        i = i + 1
                    End If
                Loop While strVal <> ""
            End With
            ReDim Preserve m.portfolios(i - 1)
            modSScenario_ReadMetadata = True
End Function