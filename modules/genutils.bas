Attribute VB_Name = "genutils"
Option Compare Database

Sub logError(pgm As String, msg As String, rowNum As Integer)
    Dim db As Database
    Set db = CurrentDb
    Debug.Print "Errore in riga " & rowNum & ": " & msg
    db.Execute "INSERT INTO LogErrori ( nomPgm, numRiga, msg) VALUES ('" & pgm & "'," & rowNum & ",'" & msg & "')"
    db.Close
    Set db = Nothing
End Sub

Function doubleApex(s As String)
    doubleApex = Replace(s, "'", "''")
End Function
