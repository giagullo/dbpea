Attribute VB_Name = "genutils"
Option Compare Database

Sub logError(pgm As String, msg As String, rowNum As Long)
    Dim db As Database
    Set db = CurrentDb
    Debug.Print "Errore in riga " & rowNum & ": " & msg
    db.Execute "INSERT INTO LogErrori ( nomPgm, numRiga, msg) VALUES ('" & pgm & "'," & rowNum & ",'" & msg & "')"
    db.Close
    Set db = Nothing
End Sub
Function logError_count(pgm As String, time As Date) As Long
    Dim db As Database
    Set db = CurrentDb
    Dim r As Recordset
    Dim s As String
    'SELECT *
    'FROM LogErrori
    ' where ID > #18/02/2019 10:52:00#;
    s = "SELECT count(*) from LogErrori where nomPgm = '" & _
        pgm & "' AND ID >= #" & time & "#"
    Set r = db.OpenRecordset(s, dbOpenDynaset)
    logError_count = r(0)
    r.Close
    Set r = Nothing
    db.Close
    Set db = Nothing
End Function
Function doubleApex(s As String)
    doubleApex = Replace(s, "'", "''")
End Function




Function IsInArray(s As String, arr() As String) As Boolean
    
    newarr = Filter(arr, s)
    n = UBound(newarr) - LBound(newarr) + 1
   
    IsInArray = (n > 0)
End Function

Private Sub testA()
    Dim arr(3) As String
    arr(0) = "AAA"
    arr(1) = "BBB"
    arr(2) = "CCC"
    arr(3) = "DDD"
    
    Const s = "CCC"
    
    Const s1 = "KKK"
    
    Debug.Print "Per s " & IsInArray(s, arr)
    Debug.Print "Per s1 " & IsInArray(s1, arr)
    Debug.Print "Per ""A"" " & IsInArray("A", arr)
End Sub

Private Sub testIqLog()
    Dim d As Date
    d = Now()
    
    Stop
    
    
    Dim l As Long
    l = logError_count("pippo", d)
    Debug.Print l
End Sub
