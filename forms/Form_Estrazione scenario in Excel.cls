VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Estrazione scenario in Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOk_Click()
    If IsNull(lstScenario.Value) Then
        MsgBox "Selezionare uno scenario"
        Exit Sub
    End If
    If lstPortfolio.ItemsSelected.Count = 0 Then
        MsgBox "Selezionare almeno un portafoglio"
        Exit Sub
    End If
    If Not Nz(cmbMese.Value, 0) > 0 Then
        MsgBox "Selezionare un mese"
        Exit Sub
    End If
    If Nz(txtAnno, 0) < 2018 Then
        MsgBox "Digitare un anno(dal 2018 in poi)"
        Exit Sub
    End If
    If Nz(txtNumMesi, 0) < 1 Or Nz(txtNumMesi, 0) > 12 Then
        MsgBox "Digitare il numero di mesi (da 1 a 12)"
        Exit Sub
    End If
    
    Debug.Print lstScenario.Value, cmbMese.Value, txtAnno, txtNumMesi
    
    For Each v In lstPortfolio.ItemsSelected
        Debug.Print lstPortfolio.ItemData(v)
    Next v
End Sub

Private Sub Corpo_Click()

End Sub
