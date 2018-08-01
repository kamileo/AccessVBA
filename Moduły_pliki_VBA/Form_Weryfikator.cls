VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Weryfikator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()

DoCmd.OpenForm "Dodaj/Aktualizuj"

End Sub

Private Sub Obraz470_Click()
    Home = InStrRev(Me.Hiper³¹cze, "\", -1, vbTextCompare)
    Call Shell("Explorer.exe" & " " & Replace(Left(Me.Hiper³¹cze, Home), "#", "", , , vbTextCompare), vbNormalFocus)
    'MsgBox ("Explorer.exe" & " " & Replace(Left(Me.Hiper³¹cze, Home), "#", "", , , vbTextCompare))
End Sub
