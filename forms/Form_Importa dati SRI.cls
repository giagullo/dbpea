VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Importa dati SRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOk_Click()
    If Nz(txtAnno, 0) < 2018 Then
        MsgBox "Inserire un anno valido", Title:="Errore input"
        Exit Sub
    End If
    If Nz(cmbMese, 0) = 0 Then
        MsgBox "Inserire un mese valido", Title:="Errore input"
        Exit Sub
    End If
    Debug.Print "Call modSRI_verifyOverride"
    
    Dim nExisting  As Long
    nExisting = modSRI_verifyOverride(cmbMese.Value, txtAnno.Value)
    If nExisting Then
        conf = MsgBox("Esistono " & nExisting & " record Utilizzo per lo stesso mese/anno. Cancellarli e sostituirli?", vbOKCancel, "Conferma")
        If conf <> vbOK Then
            Exit Sub
        End If
    End If
    
    Dim sFile As String
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show
        sFile = .SelectedItems.Item(1)
        If sFile = "" Then
            Exit Sub
        End If
        
    End With
    Debug.Print "file: " & sFile
    Dim nInserted As Long, nDiscard As Long
        
    modSRI_importData txtAnno.Value, cmbMese.Value, sFile, (nExisting > 0), nInserted, nDiscard
    
    MsgBox "Importazione completata. Inseriti " & nInserted & " record. Scartarti " & nDiscard & " record", vbOKOnly, "Operazione completata"
    
    
End Sub

Private Sub Form_Load()
    txtAnno = year(Now())
    cmbMese = 1
End Sub
