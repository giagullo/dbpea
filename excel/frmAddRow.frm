VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddRow 
   Caption         =   "Nuova Riga Allocazione"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475
   OleObjectBlob   =   "frmAddRow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

    ' MsgBox lstTasks.List(lstTasks.ListIndex)
    If lstTasks.ListIndex = -1 Then
        MsgBox "Selezionare un task ", vbCritical, "Errore input"
        Exit Sub
    End If
    templateDataCollection_addRow (lstTasks.List(lstTasks.ListIndex))
    Me.Hide
End Sub

Private Sub cmdAnnulla_Click()
    Unload Me
End Sub

Private Sub lstTasks_Click()

End Sub

Private Sub lstTasks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    templateDataCollection_addRow (lstTasks.List(lstTasks.ListIndex))
    Me.Hide
End Sub
