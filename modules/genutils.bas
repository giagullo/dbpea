Option Compare Database
Declare Function CSM_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nsize As Long) As Long
Global global_lngErrorAction As Long


Sub logError(pgm As String, msg As String, rowNum As Long)
    Dim db As Database
    Set db = CurrentDb
    Debug.Print "Errore in riga " & rowNum & ": " & msg
    db.Execute "INSERT INTO LogErrori ( nomPgm, numRiga, msg) VALUES ('" & pgm & "'," & rowNum & ",'" & msg & "')"
    db.Close
    Set db = Nothing
End Sub
Function ErrorLogging(lngErrorNo As Long, _
                                        strErrorText As String, _
                                        strCallingcode As String) As Long
' This is an example of a generic error handler
' It logs the error in a table and then allows the calling procedure
' to decide on the action to take
    Dim db As DAO.Database
    Dim rst As Recordset
    Set db = CurrentDb
    Set rst = db.OpenRecordset("tblErrorLog", dbOpenDynaset)
    With rst
        .AddNew
        !ErrorNo = lngErrorNo
        !ErrorMessage = strErrorText
        !ErrorProc = strCallingcode
        ' Also very useful to log the id of the user generating the error
        !WindowsUserName = modErrorHandler_GetUserName()
        .Update
    End With

' now allow the user to decide what happens next
    DoCmd.OpenForm "frmError", , , , , acDialog, strErrorText
    ' now as this is a dialog form our code stops here
    ' we will use a global variable to work out
    ' what action the user wants to happen
    ErrorLogging = global_lngErrorAction
End Function


Function modErrorHandler_GetUserName() As String
    Dim USER As String
    USER = Space(255)
    If CSM_GetUserName(USER, Len(USER) + 1) <> 1 Then
        modErrorHandler_GetUserName = ""
    Else
        USER = Trim$(USER)
        USER = Left(USER, Len(USER) - 1)
        modErrorHandler_GetUserName = USER
    End If
End Function
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