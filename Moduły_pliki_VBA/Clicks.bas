Attribute VB_Name = "Clicks"
Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)
Public lista_ark As String
Public shtpick As String
Public missing As String

Public Sub Get_Data_Clean(Opcja As Long, Rodzaj_dzia³ania As String, Rodzaj_Pliku As String, Kal_Kso As String)
'×××××××××××××××××××××××DEKLARACJA ZMIENNYCH×××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
Dim Location, Number, nazwa_p, Hp, Lista As String
Dim Path_Main As String: Path_Main = "\\WAWFS01\Groups$\All\Klient Sieciowy\KALKULACJE\"
Dim objAtt As Outlook.Attachment
Set objOL = CreateObject("Outlook.Application")
Set objNS = objOL.GetNamespace("MAPI")
Dim oMail As Outlook.MailItem
Dim xl_app As Object, xl_wkb As Object, xl_wks As Object, f, objItem As Object, xl_wks_p As Object
Set xl_app = CreateObject("Excel.Application")
xl_app.EnableEvents = False
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'×××××××××××××××××××××××WYBÓR RODZAJU DZIA£ANIA - NOWY REKORD, EDYCJA××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
Select Case Rodzaj_dzia³ania
    Case "nowy"
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrentDb()
        Set rs = db.OpenRecordset("Zapytania", dbOpenDynaset, dbSeeChanges)
        rs.AddNew
    Case "edycja"
        Set rs = Forms("Dodaj/Aktualizuj").Recordset
        Forms("Dodaj/Aktualizuj").Recordset.Edit
End Select
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'×××××××××××××××××××××××WYBÓR RÓD£A PLIKU - HP,RÊCZNIE,MAIL×××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
Select Case Rodzaj_Pliku
    Case "wybór" 'wybór pliku z okna dialogowego
        Set f = Application.FileDialog(1)
        f.AllowMultiSelect = False
        If f.Show Then FileName = CStr(f.SelectedItems(1))
        If FileName = "" Then GoTo NO_FILE
    Case "Hp"
        Set Rso = Forms("Dodaj/Aktualizuj").Recordset
        Hp = Rso("Hiper³¹cze").Value
        FileName = Right(Hp, Len(Hp) - 1)
    Case "mail" 'wybór pliku z maila, który jest aktualnie otwarty
        For Each objItem In ActiveExplorer.Selection
            If objItem.MessageClass = "IPM.Note" Then
                Set oMail = objItem
                For Each objAtt In oMail.Attachments
                    If Right(objAtt, 4) = "xlsm" Then
                        FileName = "O:\Klient Sieciowy\KALKULACJE\2018\TEMP\" & Replace(Now(), ":", "_", 1) & objAtt.DisplayName
                        objAtt.SaveAsFile FileName
                    End If
                Next
            End If
        Next
        If FileName = "" Then
            MsgBox ("Wybierz E-mail z odpowiednim plikiem w za³¹czniku!")
            Exit Sub
        End If
End Select
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'×××××××××××××××××××××××WILD CARD - OPCJA DLA STAREJ FORMATKI - DOCELOWO DO USUNIÊCIA××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
If Opcja = 8 Then
    rs("Priorytet") = "3 - Niski"
    Set xl_wkb = xl_app.Workbooks.Open(FileName)
    rs("Location").Value = InputBox("Podaj lokalizacjê")
    rs("Details").Value = InputBox("Podaj kody odpadów")
    lok_str = rs("Location").Value
    Kod_Str = rs("Details").Value
    Set xl_wks = xl_wkb.Sheets("KSO_Wycena_odpadów")
    rs("Company name").Value = xl_wks.Range("C3").Value
    rs("MARKET").Value = xl_wks.Range("C4").Value
    rs("Segment").Value = xl_wks.Range("C5").Value
    rs("Sub-segment").Value = xl_wks.Range("C6").Value
    rs("SUB-SUB-SEGMENT").Value = xl_wks.Range("C7").Value
    rs("Month").Value = CInt(DatePart("m", Date))
    rs("Year").Value = CInt(DatePart("yyyy", Date))
    rs("Data utworzenia") = Now()
    rs("Zarejestrowane przez") = Environ("username")
    rs("Bids deadline") = xl_wks.Range("C13").Value
    GoTo MANUAL
End If
If Opcja = 9 Then
    Set xl_wkb = xl_app.Workbooks.Open(FileName)
    GoTo MANUAL
End If
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'×××××××××××××××××××××××PRZYPISYWANIE SKOROSZYTU I ARKUSZY DO ZMIENNYCH××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
'dodana funkcja - w przypadku braku szukanego arkusza pojawia siê formularz z mo¿iwoœci¹ wyboru prawid³owego
Set xl_wkb = xl_app.Workbooks.Open(FileName)
Set xl_wks = xl_wkb.Sheets(arkusze("KSO_odpady", xl_wkb))
Set xl_wkt = xl_wkb.Sheets(arkusze("Tabela informacyjna", xl_wkb))
Set xl_wks_p = xl_wkb.Sheets(arkusze("Podsumowanie na konrakt", xl_wkb))
Set xl_wkc = xl_wkb.Sheets(arkusze("Cennik do oferty", xl_wkb))
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'×××××××××××××××××××××××ZACI¥GANIE DANYCH Z FORMATKI W PRZYPADKU NOWEGO REKORDU/NOWEJ KALKULACJI×××××××××××××××××××××××××××××××××××××××××××××××××
Kod_Str = Funkcje.Policz(Tabela(xl_wks, "Kalkulacja") & "[Kod odpadu *w ofercie SUEZ]", xl_wks, "Strumieni")
lok_str = Funkcje.Policz(Tabela(xl_wks, "Kalkulacja") & "[Miasto]", xl_wks, "Lokalizacji")

If Rodzaj_dzia³ania = "nowy" Then
    'Weryfikacja×××××××××××××××××××××××××××××××××××
    rs("Location").Value = InputBox("SprawdŸ poprawnoœæ lokalizaji: " & lok_str, "Sprawdzaczek", Trim(lok_str))
    rs("Company name").Value = InputBox("SprawdŸ poprawnoœæ nazwy firmy: " & xl_wkt.Range("B1").Value, "Sprawdzaczek", Trim(xl_wkt.Range("B1").Value))
    rs("Contract duration") = InputBox("SprawdŸ poprawnoœæ d³ugoœci kontraktu: ", "Sprawdzaczek", xl_wkt.Range("B13").Value)
    If xl_wkt.Range("B11").Value = "" Or xl_wkt.Range("B11").Value = 0 Then
    Else
        rs("Bids deadline") = InputBox("SprawdŸ poprawnoœæ Deadline: ", "Sprawdzaczek", Replace(xl_wkt.Range("B11").Value, ".", "-"))
    End If
    'Sta³e×××××××××××××××××××××××××××××××××××××××××
    rs("Priorytet") = "3 - Niski"
    rs("Month").Value = CInt(DatePart("m", Date))
    rs("Year").Value = CInt(DatePart("yyyy", Date))
    rs("Zarejestrowane przez") = Environ("username")
    rs("Data utworzenia") = Now()
    rs("Inqury date from sales rep") = Date
    rs("Result K") = "On-Going"
    rs("Result H") = "On-Going"
    rs("Czy temat jest zamkniêty?") = 0
    rs("Business Line") = "N/A"
    rs("Country") = "SUEZ Polska"
    'Pobieranie danych××××××××××××××××××××××××××××××
    rs("Details").Value = Kod_Str
    rs("MARKET").Value = xl_wkt.Range("B4").Value
    rs("Segment").Value = xl_wkt.Range("B5").Value
    rs("Sub-segment").Value = xl_wkt.Range("B6").Value
    rs("SUB-SUB-SEGMENT").Value = xl_wkt.Range("B7").Value
    rs("Sales representative").Value = xl_wkt.Range("B9").Value
    rs("Waste mass produced") = xl_wks.ListObjects(Tabela(xl_wks, "Kalkulacja")).ListColumns("Szacunkowa iloœæ odpadu /rok").Total
    rs("Kod pocztowy") = Funkcje.Policz(Tabela(xl_wks, "Kalkulacja") & "[Kod pocztowy]", xl_wks, "Kodów")
    rs("Adres") = Funkcje.Policz(Tabela(xl_wks, "Kalkulacja") & "[Ulica]", xl_wks, "Lokalizacji")
    rs("nazwa/numer lokalizacji") = Funkcje.Policz(Tabela(xl_wks, "Kalkulacja") & "[Nazwa/Numer Obiektu]", xl_wks, "numerów/nazw")
End If
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'××××××××××××××××××××××××DEKLAROWANIE G£ÓWNYCH CZÊŒCI NAZEWNICTWA××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
MANUAL:
Loc_main = Path_Main & CInt(DatePart("yyyy", Date)) & "\" & rs("Company name").Value & "\"
Name_main = "(" & rs("Company name").Value & ")" & "(" & rs("Location").Value & ")" & "(" & Date & ")" & "(" & rs("Details").Value & ")"
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'××××××××××××××××××××××××DODATKOWE OPCJE/MO¯LIWOŒCU ZMIAN ZALE¯NIE OD PRZYPADKU×××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
Select Case Opcja
    Case 1 'nowe KSO, œwie¿e - zwiêkszenie g³ównego numeru kalkulacji, Sub no1 i Sub no2 jako jedynki "1"
        rs("No") = DMax("[No]", "Zapytania") + 1
        rs("Sub_No_1") = 1
        rs("Sub_No_2") = 1
    Case 2 'nowe KSO, ponowne - iteracja Sub no2, nowe zapytanie w tej samej lokalizacji ( nowe strumienie )
        rs("No") = CInt(Forms![Dodaj/Aktualizuj]![p_no])
        rs("Sub_No_1") = DMax("[Sub_No_1]", "Zapytania", "[No] = CInt(Forms![Dodaj/Aktualizuj]![p_no])") + 1
        rs("Sub_No_2") = 1
     Case 3 'Tak zwane z marz¹_x
        rs("No") = CInt(Forms![Dodaj/Aktualizuj]![p_no])
        rs("Sub_No_1") = Forms("Dodaj/Aktualizuj").Recordset("Sub_No_1")
        rs("Sub_No_2") = Forms("Dodaj/Aktualizuj").Recordset("Sub_No_2") + 1
        Set Rso = Forms("Dodaj/Aktualizuj").Recordset
        Rso.Edit
        Rso("Result K") = "Calculation update"
        Rso("Result H") = "Calculation update"
        Rso.Update
    Case 4 ' Wrzucanie pliku z wycen
        xl_wkt.Range("B16").Value = rs("No").Value & "_" & rs("Sub_No_1") & "ver" & rs("Sub_No_2")
        rs("Wycenione przez") = xl_wkt.Range("B17").Value
    Case 5 'Generowanie oferty handlowej - uzupe³nianie danych do formatki (do dodania dodatkowe info)
        rs("CAPEX") = xl_wks_p.ListObjects(Tabela(xl_wks_p, "Capex_KON")).DataBodyRange(1, 2)
        rs("Waste mass produced") = xl_wks.ListObjects(Tabela(xl_wks, "Kalkulacja")).ListColumns("Szacunkowa iloœæ odpadu /rok").Total
        rs("GM (%)") = xl_wks_p.ListObjects(Tabela(xl_wks_p, "Podsumowanie_KON")).ListColumns("Rentownoœæ [%]").Total
        rs("Total revenue") = xl_wks_p.ListObjects(Tabela(xl_wks_p, "Podsumowanie_KON")).ListColumns("Obrót [PLN/Kontrakt]").Total
        rs("Zysk") = xl_wks_p.ListObjects(Tabela(xl_wks_p, "Podsumowanie_KON")).ListColumns("Zysk [PLN/kontrakt]").Total
        rs("Suez p³aci Klientowi") = xl_wks_p.ListObjects(Tabela(xl_wks_p, "Podsumowanie_KON")).ListColumns("SUEZ p³aci klientowi [PLN/kontrakt]").Total
        rs("Klient p³aci Suez") = xl_wks_p.ListObjects(Tabela(xl_wks_p, "Podsumowanie_KON")).ListColumns("Klient p³aci SUEZ [PLN/kontrakt]").Total
        rs("Result K") = "Offer Submitted"
        rs("Result H") = "Offer Submitted"
        rs("Offer_No") = DMax("[Offer_No]", "Zapytania") + 1
        rs("Numer Oferty") = rs("Offer_No") & "." & rs("No") & "." & rs("Sub_No_1") & "." & rs("Sub_No_2")
        xl_wkc.Range("D2").Value = rs("Numer Oferty")
        xl_app.Run ("'" & xl_wkb.Name & "'!Generuj_Handlowiec")
    Case 6 'z mar¿¹ 2 na podstawie wpisu
    Case 7 'Generuj plik dla realizacji + zmiana statusu na NEW Awards - dodac makro do generacji w formatce
        rs("Result K") = "New Awards"
        rs("Result H") = "New Awards"
        xl_app.EnableEvents = False
        xl_app.DisplayAlerts = False
        xl_app.ScreenUpdating = False
        xl_app.Run "Hide_start", 4
        xl_app.Run "BP_h", [Kalkulacja], "dkp", "dkk"
            If Len(Dir(xl_wkb.Path & "\" & "REALIZACJA", vbDirectory)) = 0 Then
               MkDir xl_wkb.Path & "\" & "REALIZACJA"
            End If
        xl_wkb.SaveAs FileName:=xl_wkb.Path & "\" & "REALIZACJA\" & Replace(xl_wkb.Name, ".xlsm", "_R.xlsx"), FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        xl_app.ScreenUpdating = True
        xl_app.DisplayAlerts = True
        xl_app.EnableEvents = True
    Case 8 'rêcznie
        rs("No") = DMax("[No]", "Zapytania") + 1
        rs("Sub_No_1") = 1
        rs("Sub_No_2") = 1
    Case 9 ' rêcznie plik z wycen
End Select
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'××××××××××××××××××××××××ZAPISYWANIE PLIKU/TWORZENIE FOLDERÓW××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
Select Case Kal_Kso
    Case "KSO" 'Zapisywanie pliku w folderze KSO
        nazwa_p = Trim(Loc_main) & rs("Location").Value & "\" & "KSO\" & "(" & rs("No").Value & ")" & Name_main & "_" & "KSO" & "_" & rs("Sub_No_1") & ".xlsm"
        Shell "cmd /c md " & """" & Trim(Loc_main) & rs("Location").Value & "\" & "KSO\" & """"
        Do Until Dir(Trim(Loc_main) & rs("Location").Value & "\" & "KSO\", vbDirectory) <> ""
            Sleep 500
        Loop
        xl_wkb.SaveAs (nazwa_p)
        rs("Hiper³¹cze") = "#" & nazwa_p
    Case "Kalkulacja" 'Zapisywanie pliku w folderze Kalkulacje + opcja na miasto
        If Opcja = 8 Or Opcja = 9 Then
            Loc_main = Loc_main & rs("Location").Value & "\"
        End If
        If InStr(1, lok_str, "Strumieni", vbTextCompare) > 0 Then
            nazwa_p = Loc_main & rs("No") & "_" & rs("Location").Value & "\" & "(" & rs("No").Value & ")" & Name_main & "_" & rs("Sub_No_1") & "ver" & rs("Sub_No_2") & ".xlsm"
            Shell "cmd /c md " & """" & Loc_main & rs("No") & "_" & rs("Location").Value & "\" & """"
            Do Until Dir(Loc_main & rs("No") & "_" & rs("Location").Value & "\", vbDirectory) <> ""
                Sleep 500
            Loop
        Else
            nazwa_p = Loc_main & rs("Location").Value & "\" & "(" & rs("No").Value & ")" & Name_main & "_" & rs("Sub_No_1") & "ver" & rs("Sub_No_2") & ".xlsm"
            Shell "cmd /c md " & """" & Loc_main & rs("Location").Value & "\" & """"
            Do Until Dir(Loc_main & rs("Location").Value & "\", vbDirectory) <> ""
                Sleep 500
            Loop
        End If
        rs("Hiper³¹cze") = "#" & nazwa_p
        xl_wkb.SaveAs (nazwa_p)
    Case "ND"
End Select
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'××××××××××××××××××××××××PRZYPISYWANIE NUMERÓW/DANE DO WERYFIKATORA××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
If Rodzaj_dzia³ania = "nowy" Then rs("Index_number") = rs("No") & "_" & rs("Sub_No_1") & "_" & rs("Sub_No_2")
If IsNull(rs("Index")) Then rs("Index").Value = DMax("[INDEX]", "Zapytania") + 1
SQL_r = "[INDEX]=" & rs![INDEX]
rs.Update
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'××××××××××××××××××××××××ZAMYKANIE EXCELA, KOÑCZENIE, OTWIERANIE WERYFIKATORA××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
xl_app.DisplayAlerts = False
xl_wkb.Close
xl_app.DisplayAlerts = True
Set xl_wkb = Nothing
Set xl_app = Nothing
MsgBox "Gotowe - sprawdŸ teraz poprawnoœæ danych"
DoCmd.Close acForm, "Dodaj/Aktualizuj"
DoCmd.OpenForm "Weryfikator", acNormal, , SQL_r
Exit Sub
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'××××××××××××××××××××××××ERROR HANDLING××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
NO_FILE:
MsgBox ("Nale¿y wybraæ plik - spróbuj jeszcze raz!")
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
End Sub

Public Function arkusze(sht_ As String, ByRef ex As Object)
For i = 1 To ex.Sheets.Count
    If ex.Sheets(i).Name = sht_ Then
        Lista = ""
        shtpick = sht_
        GoTo MAMYTO:
    Else
        Lista = Lista & Chr(34) & ex.Sheets(i).Name & Chr(34) & ";"
    End If
Next
MsgBox ("Inna ni¿ standardowa nazwa jednego z arkuszy. Proszê o wybranie zastêpczego.")
Lista = Left(Lista, Len(Lista) - 1)
lista_ark = Lista
missing = sht_
DoCmd.OpenForm "Form Picker", WindowMode:=acDialog
MAMYTO:
arkusze = shtpick
End Function

Public Function Tabela(ByRef sht As Object, searchtb As String)
Dim tbl As ListObject
    For Each tbl In sht.ListObjects
    If InStr(1, tbl.Name, searchtb, vbTextCompare) > 0 Then
        Tabela = tbl.Name
        GoTo FINITO
    End If
    Next tbl
FINITO:
End Function


