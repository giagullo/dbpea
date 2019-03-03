Option Compare Database

' Clone a scenario from an existing one
Sub modScenario_Clone(origin As String, newOne As String)

    On Error GoTo 0
    
    Dim db As Database
    Set db = CurrentDb
    Dim qInsert As String
    
    qInsert = "INSERT INTO PianoTask ( IDTask, dtInizio, dtFine, scenario )" & _
      "SELECT PianoTask.IDTask, PianoTask.dtInizio, PianoTask.dtFine, '" & newOne & "'" & _
      " FROM PianoTask" & _
      " WHERE (((PianoTask.scenario)='" & origin & "'));"
    db.Execute qInsert, dbFailOnError
    db.Close
    Set db = Nothing
End Sub