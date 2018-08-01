VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form Picker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    Me.cbxsht.RowSource = lista_ark
    Me.brak = missing
End Sub

Private Sub cbxsht_BeforeUpdate(Cancel As Integer)
    shtpick = cbxsht.Value
End Sub

Private Sub Polecenie11_Click()
    DoCmd.Close
End Sub
