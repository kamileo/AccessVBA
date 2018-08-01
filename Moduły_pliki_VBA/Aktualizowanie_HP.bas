Attribute VB_Name = "Aktualizowanie_HP"
Option Compare Database

Public Sub Get_Data_UPHP(Opcja As Long, Rodzaj_dzia³ania As String, Rodzaj_Pliku As String, Kal_Kso As String)
'Deklaracja zmiennych
Dim Location, Number, nazwa_p, Hp As String
Dim Path_Main As String: Path_Main = "O:\Klient Sieciowy\KALKULACJE\"
Dim objAtt As Outlook.Attachment
Set objOL = CreateObject("Outlook.Application")
Set objNS = objOL.GetNamespace("MAPI")
Dim oMail As Outlook.MailItem
Dim xl_app As Object, xlwkb As Object, xl_wks, f, objItem As Object
'Inicjalizowanie Excela
Set xl_app = CreateObject("Excel.Application")
xl_app.EnableEvents = False
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
    Select Case Rodzaj_dzia³ania
    Case "edycja"
        'Wybór aktualnego pliku na bazie Hp z formularza
        Set rs = Forms("Dodaj/Aktualizuj").Recordset
        Forms("Dodaj/Aktualizuj").Recordset.Edit
    End Select
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
    Select Case Rodzaj_Pliku
    Case "wybór" 'wybór pliku z okna dialogowego
        Set f = Application.FileDialog(1)
        f.AllowMultiSelect = False
        If f.Show Then FileName = CStr(f.SelectedItems(1))
        'Error Handling dla braku wybranego pliku
        If FileName = "" Then GoTo NO_FILE
    End Select
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

'przypisywanie skoroszytu i arkuszy
Set xl_wkb = xl_app.Workbooks.Open(FileName)

'Deklarowanie g³ówny czêœci nazewnictwa
Loc_main = Path_Main & CInt(DatePart("yyyy", Date)) & "\" & rs("Company name").Value & "\"
Name_main = "(" & rs("Company name").Value & ")" & "(" & rs("Location").Value & ")" & "(" & Date & ")" & "(" & rs("Details").Value & ")"

lok_str = rs("Location").Value
Select Case Kal_Kso
    Case "KSO" 'Zapisywanie pliku w folderze KSO
        nazwa_p = Loc_main & lok_str & "\" & "KSO\" & "(" & rs("No").Value & ")" & Name_main & "_" & "KSO" & "_" & rs("Sub_No_1") & ".xlsm"
        Shell "cmd /c md " & """" & Loc_main & lok_str & "\" & "KSO\" & """": Sleep 5000
        xl_wkb.SaveAs (nazwa_p)
        rs("Hiper³¹cze") = "#" & nazwa_p
    Case "Kalkulacja" 'Zapisywanie pliku w folderze Kalkulacje + opcja na miasto
        If InStr(1, lok_str, "kodów", vbTextCompare) > 0 Then
            nazwa_p = Loc_main & rs("No") & "_" & lok_str & "\" & "(" & rs("No").Value & ")" & Name_main & "_" & rs("Sub_No_1") & "ver" & rs("Sub_No_2") & ".xlsm"
            Shell "cmd /c md " & """" & Loc_main & rs("No") & "_" & lok_str & "\" & """": Sleep 5000
        Else
            nazwa_p = Loc_main & lok_str & "\" & "(" & rs("No").Value & ")" & Name_main & "_" & rs("Sub_No_1") & "ver" & rs("Sub_No_2") & ".xlsm"
            Shell "cmd /c md " & """" & Loc_main & lok_str & "\" & """": Sleep 5000
        End If
        rs("Hiper³¹cze") = "#" & nazwa_p
        xl_wkb.SaveAs (nazwa_p)
    Case "ND"
End Select
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
SQL_r = "[INDEX]=" & rs![INDEX]
rs.Update
'Koñczenie - zamkniêcie skoroszytu
xl_app.DisplayAlerts = False
xl_wkb.Close
xl_app.DisplayAlerts = True
Set xl_wkb = Nothing
Set xl_app = Nothing
MsgBox "Gotowe - sprawdŸ teraz poprawnoœæ danych"
DoCmd.Close acForm, "Dodaj/Aktualizuj"
DoCmd.OpenForm "Weryfikator", acNormal, , SQL_r
Exit Sub
'Error Handling section
NO_FILE:
MsgBox ("Nale¿y wybraæ plik - spróbuj jeszcze raz!")
End Sub

