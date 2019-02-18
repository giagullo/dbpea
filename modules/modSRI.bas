Attribute VB_Name = "modSRI"
Option Compare Database
Dim db As Database

' Import data from excel sFile into Utilizzo for an year and a month
' deleting records first if override
' Returns BYRef number of record committed, number of record discarded

Sub modSRI_importData(aYear As Integer, aMonth As Integer, sFile As String, override As Boolean, _
        ByRef numCommitted As Long, ByRef numDiscarded As Long)

    ' modSRI_importRowData
    
    ' Start processing imported table
    
    Dim rstTask As Recordset
    Dim rstRisorsa As Recordset
    Dim idTask As Integer, idRisorsa As Integer
    Dim sSqlInsert As String
    
    Set db = CurrentDb

    Set rstTask = db.OpenRecordset("Task", dbOpenDynaset)
    Set rstRisorsa = db.OpenRecordset("Risorsa", dbOpenDynaset)
    Dim startTime As Date
    startTime = Now()
    
    ' Clean current records if needed
    If override Then
        Debug.Print "deleting rows from Utilizzo"
        db.Execute ("delete from Utilizzo where mese >= " & CLng(aYear) * 100 + aMonth)
    End If
    
    ' open excel file
    Dim ok As Boolean
    ok = modExcel_OpenExcel(True)
    If Not ok Then
        Err.Raise 555, Description:="Errore apertura Excel"
    End If
        
    ok = modExcel_OpenWorkBook(sFile)
    If Not ok Then
        Err.Raise 555, Description:="Errore apertura file Excel " & sFile
    End If
    modExcel_OpenActiveWorkSheet
    Dim r As Long
    r = 2
    numCommitted = 0
    numDiscarded = 0
    
    Dim sXlPrg As String
    Dim sXlTask As String, sXlBusinessPartner As String, sXlMonth As String, sXlAllocated As String
    ok = modExcel_ReadCell("A", r, sXlPrg)
    Do While sXlPrg <> ""
        ' read row r from excel
        ' TODO substitute with true column names
        ok = modExcel_ReadCell("I", r, sXlTask)
        ok = modExcel_ReadCell("D", r, sXlBusinessPartner)
        ok = modExcel_ReadCell("N", r, sXlMonth)
        ok = modExcel_ReadCell("L", r, sXlAllocated)
        If Not ok Then
            Err.Raise 555, Description:="Errore lettura da excel"
        End If
        Debug.Print "Excel data: ", sXlTask, sXlBusinessPartner, sXlMonth, sXlAllocated
        
        ' build month in AAAAMM format
        If Not IsNumeric(sXlMonth) Then
             logError "modSRI_importData", "mese non numerico", r
             GoTo avanti
        End If
        If CDec(sXlMonth) <> aMonth Then
            GoTo avanti
        End If
        s = CLng(aYear) * 100 + aMonth
        
        ' recupera ID task
        rstTask.FindFirst ("codSIPROS = '" & sXlTask & "'")
        If rstTask.NoMatch Then
            logError "modSRI_importData", "Task " & sXlTask & " non trovata", r
            GoTo avanti
        End If
        idTask = rstTask!ID
        
        ' recupera ID risorsa
        rstRisorsa.FindFirst ("Nome = """ & sXlBusinessPartner & """")
                
        If rstRisorsa.NoMatch Then
            logError "modSRI_importData", "Risorsa " & doubleApex(sXlBusinessPartner) & " non trovata", r
            GoTo avanti
        End If
        idRisorsa = rstRisorsa!ID
        
        ' find utiizzo
        Dim dblUtilizzo As Double
        If Not IsNumeric(Nz(sXlAllocated, "x")) Then
            logError "modSRI_importData", "dato numerico non valido", r
            GoTo avanti
        End If
        dblUtilizzo = CDbl(sXlAllocated) * 100
        If dblUtilizzo < 0 Or dblUtilizzo > 100 Then
            logError "modSRI_importData", "dato numerico fuori range", r
            GoTo avanti
        End If
                        
        ' inserisci utilizzo
        sSqlInsert = "INSERT INTO Utilizzo (idTask, idRisorsa, mese, pct) VALUES (" & _
                    idTask & "," & _
                    idRisorsa & "," & _
                    s & "," & _
                    dblUtilizzo & ")"
        ' Stop
        
        Debug.Print sSqlInsert
        db.Execute sSqlInsert, dbFailOnError
        numCommitted = numCommitted + 1
avanti:
        r = r + 1
        ok = modExcel_ReadCell("A", r, sXlPrg)
        If Not ok Then
            Err.Raise 555, Description:="Errore lettura da excel"
        End If
    Loop
    
    ' rstImport.Close
    rstTask.Close
    rstRisorsa.Close
    db.Close
    ' Set rstImport = Nothing
    Set rstTask = Nothing
    Set rstRisorsa = Nothing
    Set db = Nothing
    
    numDiscarded = logError_count("modSRI_importData", startTime)
    
End Sub
Function modSRI_verifyOverride(aMonth As Integer, aYear As Integer) As Long
    Dim db As Database
    Dim rs As Recordset
    Dim numExisting As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("select count(*) from Utilizzo where mese = " & CLng(aYear) * 100 + aMonth)
    numExisting = rs(0)
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    modSRI_verifyOverride = numExisting
End Function




