Attribute VB_Name = "ModOneOff"
Option Compare Database

Sub mod1_importUtilizzo()
    Dim db As Database
    Dim rstImport As Recordset
    Dim rstUtilizzo As Recordset
    Dim rstTask As Recordset
    Dim rstRisorsa As Recordset
    Dim idTask As Integer, idRisorsa As Integer
    Const anno = "2018"
    Set db = CurrentDb
    Set rstImport = db.OpenRecordset("tblSRI", dbOpenDynaset)
    Set rstTask = db.OpenRecordset("Task", dbOpenDynaset)
    Set rstRisorsa = db.OpenRecordset("Risorsa", dbOpenDynaset)
    Set rstUtilizzo = db.OpenRecordset("Utilizzo", dbOpenDynaset)
    ' On Error Resume Next
    Do While Not rstImport.EOF
        ' recupera ID task
        rstTask.FindFirst ("codSIPROS = '" & rstImport!task & "'")
        If rstTask.NoMatch Then
            Debug.Print "Task " & rstImport!task & " non trovata"
            GoTo avanti
        End If
        idTask = rstTask!ID
        
        ' recupera ID risorsa
        rstRisorsa.FindFirst ("Nome = '" & rstImport!risorsa & "'")
        If rstRisorsa.NoMatch Then
            Debug.Print "Task " & rstImport!risorsa & " non trovata"
            GoTo avanti
        End If
        idRisorsa = rstRisorsa!ID
        
        
        ' costruisci mese
        s = Format(rstImport!mese, "00") & anno
        
        ' inserisci utilizzo
        With rstUtilizzo
            .AddNew
            !idTask = idTask
            !idRisorsa = idRisorsa
            !mese = s
            !pct = rstImport!alloc * 100
            .Update
        End With
        
        Debug.Print "Inserito "; idTask, idRisorsa, s, rstImport!alloc
avanti:
        rstImport.MoveNext
    Loop
    
    rstImport.Close
    rstTask.Close
    rstRisorsa.Close
    Set rstImport = Nothing
    Set rstTask = Nothing
    Set rstRisorsa = Nothing
    
End Sub

