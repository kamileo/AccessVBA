VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Przypominajka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Szukaj_h_Click()

If IsNull(handl_h.Value) Then
    MsgBox "Brak wybranego handlowca!"
    Exit Sub
End If

Dim rs As DAO.Recordset

i = 1
Lista_h.Value = ""

For x = 1 To 5
Select Case x
    Case 1
        Type_ = "New Awards"
        Rotatas = ">=40000"
    Case 2
        Type_ = "Renewal Awarded"
        Rotatas = ">=40000"
    Case 3
        Type_ = "Lost Client"
        Rotatas = ">=0"
    Case 4
        Type_ = "Bid Lost"
        Rotatas = ">=0"
    Case 5
        Type_ = "Offer Submitted"
        Rotatas = ">=20000"
    End Select

SQL_q = "SELECT * FROM Zapytania WHERE [Komentarz Uzupe³niaj¹cy] is null AND [Waste mass produced] >417 AND [Sales representative] = " & """" & handl_h.Value & """"
SQL_q = SQL_q & "AND [Result K] = " & """" & Type_ & """" & " AND [Total revenue] " & Rotatas

If x = 1 Then MsgBox SQL_q
Set rs = CurrentDb.OpenRecordset(SQL_q, dbOpenDynaset, dbSeeChanges)
'Set rs = CurrentDb.OpenRecordset(SQL_q)

SQL_q = ""
i = 1
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    Do Until rs.EOF = True
        If rs("Czy temat jest zamkniêty?") = 0 Then
            If x = 1 Then Lista_h.Value = "<b>" & Type_ & "</b>" & "<div>" & rs("No") & "_" & rs("Company name") & "_" & rs("Location")
            If i = 1 And x > 1 Then Lista_h.Value = Lista_h.Value & "<div>" & "<b>" & Type_ & "</b>" & "<div>" & rs("No") & "_" & rs("Company name") & "_" & rs("Location")
            If i > 1 Then Lista_h.Value = Lista_h.Value & "<div>" & rs("No") & "_" & rs("Company name") & "_" & rs("Location")
            i = i + 1
        End If
        rs.MoveNext
    Loop
End If

Next x

MsgBox "Szukanie zakoñczone. Znaleziono " & i & " Wpisów!"

'×××××××××××××××××××××××××××××××××××××TWORZENIE MAILA DO HANDLOWCA×××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
Dim oApp As Outlook.Application
Dim oMail As MailItem
Set oApp = CreateObject("Outlook.application")
Set oMail = oApp.CreateItem(olMailItem)
oMail.Display
Signature = oMail.HTMLBody
oMail.HTMLBody = Lista_h.Value & Signature
oMail.Subject = "Braki komentarzy"
oMail.To = handl_h.Value
oMail.CC = "Kalkulacje@sitapolska.com.pl"
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

rs.Close
Set rs = Nothing
End Sub
