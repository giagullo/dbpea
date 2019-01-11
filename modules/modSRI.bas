Attribute VB_Name = "modSRI"
Option Compare Database
Dim db As Database

Sub modSRI_importData()
    
    modSRI_importRowData
    
    ' Start processing imported table
    
    Dim rstImport As Recordset
    Dim rstUtilizzo As Recordset
    Dim rstTask As Recordset
    Dim rstRisorsa As Recordset
    Dim idTask As Integer, idRisorsa As Integer
    Dim sAnno As String
    Dim rowNum  As Integer
    rowNum = 1
    sAnno = InputBox("Indicare l'anno di cui si sta elaborando il consuntivo", "Input richiesto")
    If Not IsNumeric(sAnno) Then
        MsgBox "Digitare un anno in formato AAAA", vbCritical, "Errore input"
        Exit Sub
    End If
    Dim anno As Integer
    anno = CDec(sAnno)
    If anno < 2018 Or anno > Year(Now) Then
        MsgBox "Anno fuori range", vbCritical, "Errore input"
        Exit Sub
    End If
    
    Set db = CurrentDb
    Set rstImport = db.OpenRecordset("tblTempSRI", dbOpenDynaset)
    Set rstTask = db.OpenRecordset("Task", dbOpenDynaset)
    Set rstRisorsa = db.OpenRecordset("Risorsa", dbOpenDynaset)
    
    
    ' On Error GoTo finally
    
    ' Clean current year, replace all data from SRI
    db.Execute ("delete from Utilizzo where mese >= " & sAnno & "00")
    
    Set rstUtilizzo = db.OpenRecordset("Utilizzo", dbOpenDynaset)
    
    Do While Not rstImport.EOF
        ' recupera ID task
        rstTask.FindFirst ("codSIPROS = '" & rstImport!Task & "'")
        If rstTask.NoMatch Then
            logError "modSRI_importData", "Task " & rstImport!Task & " non trovata", rowNum
            GoTo avanti
        End If
        idTask = rstTask!ID
        
        ' recupera ID risorsa
        rstRisorsa.FindFirst ("Nome = """ & rstImport![Business partner] & """")
                
        If rstRisorsa.NoMatch Then
            logError "modSRI_importData", "Risorsa " & doubleApex(rstImport![Business partner]) & " non trovata", rowNum
            GoTo avanti
        End If
        idRisorsa = rstRisorsa!ID
        
        
        ' build month in AAAAMM format
        s = anno & Format(rstImport![Mese Fine], "00")
        
        ' find utiizzo
        Dim dblUtilizzo As Double
        If Not IsNumeric(rstImport![Consuntivi di mesi allocati]) Then
            If rstImport![Consuntivi di mesi allocati] <> "" Then
                logError "modSRI_importData", "dato numerico non valido", rowNum
            End If
            GoTo avanti
        End If
        dblUtilizzo = CDbl(rstImport![Consuntivi di mesi allocati]) * 100
            
        ' inserisci utilizzo
        With rstUtilizzo
            .AddNew
            !idTask = idTask
            !idRisorsa = idRisorsa
            !mese = s
            !pct = dblUtilizzo
            .Update
        End With
        
        ' Debug.Print "Inserito "; idTask, idRisorsa, s, dblUtilizzo
avanti:
        rstImport.MoveNext
        rowNum = rowNum + 1
    Loop

finally:
    rstImport.Close
    rstTask.Close
    rstRisorsa.Close
    Set rstImport = Nothing
    Set rstTask = Nothing
    Set rstRisorsa = Nothing
    
End Sub



Private Sub modSRI_importRowData()

    ' Remove temp table if exists
    On Error Resume Next
    DoCmd.DeleteObject acTable, "tblTempSRI"
    On Error GoTo 0
    
    ' run saved import of SRI excel report
    DoCmd.RunSavedImportExport "Importa-Cruscotto2018"
End Sub



