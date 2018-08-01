Attribute VB_Name = "New Mail"
Option Compare Database
Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)

Sub sendOutlookEmail(Opcja As Integer, Typ As String) '(opcja 1 - z hiper��cza, opcja 2 - manualnie)

'�������������������������������������DEFINIOWANIE ZMIENNYCH���������������������������������������������������������������
Dim xl_app As Object, xl_wkb As Object, xl_wks, f, objItem As Object, xl_wks_p As Object
Dim tre��, temat, maildo, maildw, za��czenik As String
Set xl_app = CreateObject("Excel.Application")
Set rs = Forms("Dodaj/Aktualizuj").Recordset
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'�������������������������������������WYB�R PLIKU - HP, W�ASNY WYB�R�������������������������������������������������������������������������
Select Case Opcja
Case 1:
    If rs("Hiper��cze") <> "" Then
        On Error GoTo WrongFormat
        Att = CStr(Replace(rs("Hiper��cze"), "#", "", 1, , vbTextCompare))
        On Error GoTo 0
    Else
        MsgBox ("Brak hiper��cza, wybierz plik r�cznie!")
        GoTo Manualnie
    End If
Case 2:
Manualnie:
    Set f = Application.FileDialog(1)
    f.AllowMultiSelect = False
    If f.Show Then Att = CStr(f.SelectedItems(1))
    If Att = "" Then GoTo NO_FILE
End Select

On Error GoTo WrongFormat
Set xl_wkb = xl_app.Workbooks.Open(Att)
Set xl_wks = xl_wkb.Sheets(arkusze("KSO_odpady", xl_wkb))
Set xl_wkt = xl_wkb.Sheets(arkusze("Tabela informacyjna", xl_wkb))
Set xl_wks_p = xl_wkb.Sheets(arkusze("Podsumowanie na konrakt", xl_wkb))
Set xl_wkc = xl_wkb.Sheets(arkusze("Cennik do oferty", xl_wkb))
On Error GoTo 0
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
rs.Edit
If Typ = "oferta" Then rs("Result K") = "New Awards"
rs.Update
'�������������������������������������ZMIENNE DO MAILA�����������������������������������������������������������������������������������������
Select Case Typ
    Case "akceptacja"
        tre�� = "Prosz� o akceptacj�:" & "<br>" & "<br>" & "Obr�t: " & CStr(xl_wks_p.ListObjects(Tabela(xl_wks_p, "Podsumowanie_KON")).ListColumns("Obr�t [PLN/Kontrakt]").Total) & _
            "<br>" & "Zysk: " & CStr(xl_wks_p.ListObjects(Tabela(xl_wks_p, "Podsumowanie_KON")).ListColumns("Zysk [PLN/kontrakt]").Total) & _
            "<br>" & "Rentowno��: " & Format(CStr(xl_wks_p.ListObjects(Tabela(xl_wks_p, "Podsumowanie_KON")).ListColumns("Rentowno�� [%]").Total), "Percent") & _
            "<br>" & "CAPEX: " & CStr(xl_wks_p.ListObjects(Tabela(xl_wks_p, "Capex_KON")).DataBodyRange(1, 2))
        temat = "Do akceptacji - " & rs("Company name") & " ," & rs("Location") & " ," & rs("Index_number")
        maildo = "aneta.niemczak@suez.com"
        maildw = "Kalkulacje@sitapolska.com.pl"
        za��cznik = Att
    Case "oferta"
        tre�� = "zgodnie z poni�szym prosze o realizacj�."
        maildo = "umowyksc@sitapolska.com.pl"
        maildw = "Kalkulacje@sitapolska.com.pl"
        za��cznik = Att
    End Select
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'�������������������������������������W�A�CIWE MAKRO�����������������������������������������������������������������������������
Dim oApp As Outlook.Application
Dim oMail As MailItem
Set oApp = CreateObject("Outlook.application")

Select Case Typ
    Case "akceptacja"
        Set oMail = oApp.CreateItem(olMailItem)
        oMail.Subject = temat
    Case "oferta"
        For Each objItem In ActiveExplorer.Selection
            If objItem.MessageClass = "IPM.Note" Then
                Set oMail = objItem.Forward
                oMail.Subject = oMail.Subject
            End If
        Next
End Select

oMail.Display
Signature = oMail.HTMLBody
oMail.HTMLBody = tre�� & oMail.HTMLBody
oMail.To = maildo
oMail.CC = maildw
oMail.Attachments.Add (za��cznik)
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'�������������������������������������ERROR HANDLING�������������������������������������������������������������������������������������������
FINITO:
On Error Resume Next
Set oMail = Nothing
Set oApp = Nothing
xl_app.DisplayAlerts = False
If xl_wkb <> Empty Then xl_wkb.Close ' do poprawy
xl_app.DisplayAlerts = True
Set xl_wkb = Nothing
Set xl_app = Nothing
Exit Sub

WrongFormat:
MsgBox ("Nieprawid�owa nazwa arkusza/stara formatka/b��dne hiper��cze")
GoTo FINITO

NO_FILE:
MsgBox ("Plik nie zosta� wybrany, Spr�buj ponownie!")
GoTo FINITO
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
End Sub


