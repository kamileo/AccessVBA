VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Weryfikacja Komentarzy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Filtr_Handlowca_AfterUpdate()

Dim strFilter As String

strFilter = "[Sales representative]=" & """" & Filtr_Handlowca.Value & """" & " And " & _
"[Czy temat jest zamkniêty?]=" & 0

DoCmd.ApplyFilter , strFilter


End Sub

Private Sub Form_Open(Cancel As Integer)
DoCmd.Maximize
End Sub

Private Sub ZatwierdŸ_Click()

If IsNull(Komentarz_U) Then
    MsgBox "brak treœci do przeniesienia"
    GoTo FINITO
End If

    Komentarz_do_Raportu.Value = Komentarz_Uzupe³niaj¹cy.Value & Chr(13) & Chr(10) & Komentarz_do_Raportu.Value
    Komentarz_Uzupe³niaj¹cy.Value = ""

FINITO:

End Sub


Private Sub Obraz517_Click()
Dim strFilter As String

strFilter = "[Result K]=" & """" & Result_wybór.Value & """" & " And " & _
"[Sales representative]=" & """" & Filtr_Handlowca.Value & """" & " And " & _
"[Czy temat jest zamkniêty?]=" & 0

DoCmd.ApplyFilter , strFilter
End Sub

Private Sub Result_wybór_Click()

End Sub
