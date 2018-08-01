Attribute VB_Name = "z_OLD_Clicks"
''DOPASOWAÆ NAZWY NAG£ÓWKÓW W FORMATCE ALBO WYSZUKAÆ JAK WPISAÆ TEN JE***Y ENTER W NAZWIE ¯EBY £APA£O!!!
''POPRAWIÆ MAKRO GENEROWANIA PLIKU DLA HANDLOWCA
'
'Declare Sub Sleep Lib "kernel32" _
'(ByVal dwMilliseconds As Long)
'
'Public Sub Get_Data(Opcja As Long)
'
''-----------inicjalizowanie excela---------------------------------------------
'
'Dim Location As String
'Dim Number As String
'Dim Switch_Lok As Boolean
'
'Dim xl_app As Object, xlwkb As Object, xl_wks As Object
'Dim source_empl_col As Integer
'
'Dim f As Object
'Set f = Application.FileDialog(1)
'f.AllowMultiSelect = False
'If f.Show Then FileName = CStr(f.SelectedItems(1))
'
''Error Handling dla braku wybranego pliku
'If FileName = "" Then GoTo NO_FILE
'
'Set xl_app = CreateObject("Excel.Application")
'xl_app.EnableEvents = False
'Set xl_wkb = xl_app.Workbooks.Open(FileName)
'Set xl_wks = xl_wkb.Sheets("KSO_Wycena_odpadów")
'
''Wyci¹ganie kodów i lokalizacji - Funkcja Policz
'Kod_Str = Policz("Kalkulacja[Kod odpadu *wyceny]", xl_app.Workbooks.Open(FileName), "Kodów")
'Lok_Str = Policz("Kalkulacja[Miasto]", xl_app.Workbooks.Open(FileName), "Lokalizacji")
'
'Company_name = xl_wks.Range("A1").Value
'Set db = CurrentDb()
'Set Rs = db.OpenRecordset("Zapytania")
'Rs.AddNew
'
''Weryfikacja danych z KSO
'
'Rs("Location").Value = InputBox("SprawdŸ poprawnoœæ lokalizaji: " & Lok_Str, "Sprawdzaczek", Lok_Str)
'Rs("Company name").Value = InputBox("SprawdŸ poprawnoœæ nazwy firmy: " & xl_wks.Range("D3").Value, "Sprawdzaczek", xl_wks.Range("D3").Value)
'
'Rs("Details").Value = Kod_Str
'Rs("MARKET").Value = xl_wks.Range("D4").Value
'Rs("Segment").Value = xl_wks.Range("D5").Value
'Rs("Sub-segment").Value = xl_wks.Range("D6").Value
'Rs("SUB-SUB-SEGMENT").Value = xl_wks.Range("D7").Value
'Rs("Sales representative").Value = xl_wks.Range("D9").Value
'Rs("Month").Value = CInt(DatePart("m", Date))
'Rs("Year").Value = CInt(DatePart("yyyy", Date))
'Rs("Zarejestrowane przez") = Environ("username")
'Rs("Data utworzenia") = Now()
'Rs("Waste mass produced") = xl_wks.ListObjects("Kalkulacja").ListColumns("Szacunkowa iloœæ odpadu /rok").Total
'Rs("Contract duration") = xl_wks.Range("H4").Value
'Rs("Bids deadline") = xl_wks.Range("H2").Value
'Rs("Inqury date from sales rep") = Now()
'Rs("Result K") = "On-Going"
'Rs("Result H") = "On-Going"
'
'Loc_main = "C:\Users\kbronkow\Desktop\TESTY\Access\Zapytania\" & CInt(DatePart("yyyy", Date)) & "\" & Rs("Company name").Value & "\"
'Name_main = "(" & Rs("Company name").Value & ")" & "(" & Rs("Location").Value & ")" & _
'    "(" & Date & ")" & "(" & Rs("Details").Value & ")"
'
'Select Case Opcja
'    Case 1 'nowe KSO, œwie¿e - zwiêkszenie g³ównego numeru kalkulacji, Sub no1 i Sub no2 jako jedynki "1"
'
'        Rs("No") = DMax("[No]", "Zapytania") + 1
'        Rs("Sub_No_1") = 1
'        Rs("Sub_No_2") = 1
'
'        Shell "cmd /c md " & """" & Loc_main & "KSO\" & """"
'            Sleep 5000
'        xl_wkb.SaveAs (Loc_main & "KSO\" & "(" & Rs("No").Value & ")" & Name_main & "_" & "KSO" & ".xlsm")
'
'        Rs("Hiper³¹cze") = "#" & Loc_main & "KSO\" & "(" & Rs("No").Value & ")" & Name_main & "_" & "KSO" & ".xlsm"
'
'        Rs("Index_number") = Rs("No") & "_" & Rs("Sub_No_1") & "_" & Rs("Sub_No_2")
'
'    Case 2 'nowe KSO, ponowne - iteracja Sub no2, nowe zapytanie w tej samej lokalizacji ( nowe strumienie )
'        Rs("No") = CInt(Forms![Dodaj/Aktualizuj]![kso_up])
'        Rs("Sub_No_1") = DMax("[Sub_No_1]", "Zapytania", "[No] = CInt(Forms![Dodaj/Aktualizuj]![kso_up])") + 1
'        Rs("Sub_No_2") = 1
'
'        Shell "cmd /c md " & """" & Loc_main & "KSO\" & """"
'            Sleep 5000
'        xl_wkb.SaveAs (Loc_main & "KSO\" & "(" & Rs("No").Value & ")" & Name_main & "_" & "KSO" & "_" & Rs("Sub_No_1") & ".xlsm")
'
'        Rs("Hiper³¹cze") = "#" & Loc_main & "KSO\" & "(" & Rs("No").Value & ")" & Name_main & "_" & "KSO" & "_" & Rs("Sub_No_1") & ".xlsm"
'
'    Case 3 'Tak zwane z marz¹_x
'        'DODAÆ ERROR HANDLER!
'        Rs("No") = CInt(Forms![Dodaj/Aktualizuj]![kso_up])
'        Rs("Sub_No_1") = Forms("Dodaj/Aktualizuj").Recordset("Sub_No_1")
'        Rs("Sub_No_2") = Forms("Dodaj/Aktualizuj").Recordset("Sub_No_2") + 1
'
'        Lok_name = "Zapytania\" & Lok_Str & "\"
'
'        Shell "cmd /c md " & """" & Loc_main & "Zapytania\" & Lok_Str & "\" & """"
'            Sleep 5000
'        xl_wkb.SaveAs (Loc_main & "Zapytania\" & Lok_Str & "\" & "(" & Rs("No").Value & ")" & Name_main & "_" & Rs("Sub_No_1") & "ver" & Rs("Sub_No_2") & ".xlsm")
'
'          Rs("Hiper³¹cze") = "#" & Loc_main & "Zapytania\" & Lok_Str & "\" & "(" & Rs("No").Value & ")" & Name_main & "_" & Rs("Sub_No_1") & "ver" & Rs("Sub_No_2") & ".xlsm"
'
'End Select
'
'Rs("Index_number") = Rs("No") & "_" & Rs("Sub_No_1") & "_" & Rs("Sub_No_2")
'Rs.Update
'
'    xl_app.DisplayAlerts = False
'    xl_wkb.Close
'    xl_app.DisplayAlerts = True
'
'Set xl_wkb = Nothing
'Set xl_app = Nothing
'
'MsgBox ("gotowe")
'Exit Sub
'
'NO_FILE:
'MsgBox ("Nale¿y wybraæ plik - spróbuj jeszcze raz!")
'
'End Sub
'
'Public Sub Up_Data(Opcja As Long)
'
'Set Rs = Forms("Dodaj/Aktualizuj").Recordset
'Dim Switch_Lok As Boolean
'Dim Location As String
'Dim Number As String
'Dim xl_app As Object, xlwkb As Object, xl_wks As Object
'Dim source_empl_col As Integer
'Dim f As Object
'
'Forms("Dodaj/Aktualizuj").Recordset.Edit
'
'Loc_main = "C:\Users\kbronkow\Desktop\TESTY\Access\Zapytania\" & CInt(DatePart("yyyy", Date)) & "\" & Rs("Company name").Value & "\"
'Name_main = "(" & Rs("Company name").Value & ")" & "(" & Rs("Location").Value & ")" & "(" & Date & ")" & "(" & Rs("Details").Value & ")"
'
'Select Case Opcja
'    Case 1 'wrzuæ plik z wycen
'
'        Set f = Application.FileDialog(1)
'        f.AllowMultiSelect = False
'        If f.Show Then FileName = CStr(f.SelectedItems(1))
'        Set xl_app = CreateObject("Excel.Application")
'        xl_app.EnableEvents = False
'        Set xl_wkb = xl_app.Workbooks.Open(FileName)
'        Set xl_wks = xl_wkb.Sheets("KSO_Wycena_odpadów")
'        Set xl_wks_p = xl_wkb.Sheets("Podsumowanie na konrakt")
'        '
'        xl_wks.Range("H7").Value = Rs("No").Value & "_" & Rs("Sub_No_1") & "ver" & Rs("Sub_No_2")
'        '
'        Rs("Hiper³¹cze") = ""
'        Rs("Hiper³¹cze") = "#" & Loc_main & "Zapytania\" & Lok_Str & "\" & "(" & Rs("No").Value & ")" & Name_main & "_" & Rs("Sub_No_1") & "ver" & Rs("Sub_No_2") & ".xlsm"
'
'        Shell "cmd /c md " & """" & Loc_main & "Zapytania\" & Lok_Str & "\" & """"
'             Sleep 5000
'        xl_wkb.SaveAs FileName:=Loc_main & "Zapytania\" & Lok_Str & "\" & "(" & Rs("No").Value & ")" & Name_main & "_" & Rs("Sub_No_1") & "ver" & Rs("Sub_No_2") & ".xlsm"
'
'        Rs("Wycenione przez") = xl_wks.Range("H8").Value
'
'
'    Case 2
'
'        Hp = Rs("Hiper³¹cze").Value
'
'        Set xl_app = CreateObject("Excel.Application")
'        xl_app.EnableEvents = False
'        Set xl_wkb = xl_app.Workbooks.Open(Right(Hp, Len(Hp) - 1))
'        Set xl_wks = xl_wkb.Sheets("KSO_Wycena_odpadów")
'        Set xl_wks_p = xl_wkb.Sheets("Podsumowanie na konrakt")
'
'        Rs("Waste mass produced") = xl_wks.ListObjects("Kalkulacja").ListColumns("Szacunkowa iloœæ odpadu /rok").Total
'        Rs("GM (%)") = xl_wks_p.ListObjects("Podsumowanie_KON").ListColumns("Rentownoœæ [%]").Total
'        Rs("Total revenue") = xl_wks_p.ListObjects("Podsumowanie_KON").ListColumns("Obrót [PLN/Kontrakt]").Total
'
'        Rs("Zysk") = xl_wks_p.ListObjects("Podsumowanie_KON").ListColumns("Zysk [PLN/kontrakt]").Total
'        Rs("Suez p³aci Klientowi") = xl_wks_p.ListObjects("Podsumowanie_KON").ListColumns("SUEZ p³aci klientowi [PLN/kontrakt]").Total
'        Rs("Klient p³aci Suez") = xl_wks_p.ListObjects("Podsumowanie_KON").ListColumns("Klient p³aci SUEZ [PLN/kontrakt]").Total
'
'        Rs("Result K") = "Offer Submitted"
'        Rs("Result H") = "Offer Submitted"
'
'        Rs("Offer_No") = DMax("[Offer_No]", "Zapytania") + 1
'        Rs("Numer Oferty") = Rs("Offer_No") & "." & Rs("No") & "." & Rs("Sub_No_1") & "." & Rs("Sub_No_2")
'
'        'uruchamianie makra z formatki
'        xl_app.Run ("'" & xl_wkb.Name & "'!Generuj_Handlowiec")
'
'        'Inny odbiorca itp? Service
'        'CAPEX
'    Case 3 'Generuj plik dla realizacji + zmiana statusu na NEW Awards - dodac makro do generacji w formatce
'
''    Case 4 'z mar¿¹ 2 na podstawie wpisu
'''        Rs.Update
'''        Rs.AddNew
''
''        Hp = Rs("Hiper³¹cze").Value
''
''        Set xl_app = CreateObject("Excel.Application")
''        xl_app.EnableEvents = False
''        Set xl_wkb = xl_app.Workbooks.Open(Right(Hp, Len(Hp) - 1))
''
''        Rs("Hiper³¹cze") = "#" & Loc_main & "Zapytania\" & "(" & Rs("No").Value & ")" & Name_main & "_" & Rs("Sub_No_1") & "ver" & Rs("Sub_No_2") & ".xlsm"
''
''        Shell "cmd /c md " & """" & Loc_main & "Zapytania\" & """"
''             Sleep 5000
''        xl_wkb.SaveAs FileName:=Loc_main & "Zapytania\" & "(" & Rs("No").Value & ")" & Name_main & "_" & Rs("Sub_No_1") & "ver" & Rs("Sub_No_2") & ".xlsm"
''
''        Set db = CurrentDb()
''        Set Rs = db.OpenRecordset("Zapytania")
''
''        MsgBox Me.Recordset("Details")
''
''        For Each fld In db.TableDefs!Zapytania.Fields
''            If Me.Recordset(fld.Name) <> "" Then Rs(fld.Name) = Me.Recordset(fld.Name)
''        Next
''
''        Rs("Result K") = "On-Going"
''        Rs("Result H") = "On-Going"
''        Rs("Index_number") = Rs("No") & "_" & Rs("Sub_No_1") & "_" & Rs("Sub_No_2")
'
'End Select
'
'Forms("Dodaj/Aktualizuj").Recordset.Update
'
'    xl_app.DisplayAlerts = False
'    xl_wkb.Close
'    xl_app.DisplayAlerts = True
'
'Set xl_wkb = Nothing
'Set xl_app = Nothing
'
'MsgBox ("Gotowe")
'Exit Sub
'
'End Sub
'
'
'
'
'Sub Get_mail()
'
'Dim Location As String
'Dim Number As String
'Dim Switch_Lok As Boolean
'Dim xl_app As Object, xlwkb As Object, xl_wks As Object
'Dim source_empl_col As Integer
'Dim objAtt As Outlook.Attachment
'Set objOL = CreateObject("Outlook.Application")
'Set objNs = objOL.GetNamespace("MAPI")
'Dim oMail As Outlook.MailItem
'Dim objItem As Object
'
'Set xl_app = CreateObject("Excel.Application")
'xl_app.EnableEvents = False
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''
'
'For Each objItem In ActiveExplorer.Selection
'    If objItem.MessageClass = "IPM.Note" Then
'    Set oMail = objItem
'        For Each objAtt In oMail.Attachments
'            If Right(objAtt, 4) = "xlsm" Then
'
'                g = Replace(Now(), ":", "_", 1)
'                Name_temp = "C:\Users\kbronkow\Desktop\TEST\TEMP\" & g & objAtt.DisplayName
'                objAtt.SaveAsFile Name_temp
'
'            End If
'        Next
'    End If
'Next
'
'Set xl_wkb = xl_app.Workbooks.Open(Name_temp)
'Set xl_wks = xl_wkb.Sheets("KSO_Wycena_odpadów")
'
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''Wyci¹ganie kodów i lokalizacji - Funkcja Policz
'Kod_Str = Policz("Kalkulacja[Kod odpadu *wyceny]", xl_app.Workbooks.Open(Name_temp), "Kodów")
'Lok_Str = Policz("Kalkulacja[Miasto]", xl_app.Workbooks.Open(Name_temp), "Lokalizacji")
'
'Company_name = xl_wks.Range("A1").Value
'Set db = CurrentDb()
'Set Rs = db.OpenRecordset("Zapytania")
'Rs.AddNew
'
'Rs("Location").Value = InputBox("SprawdŸ poprawnoœæ lokalizaji: " & Lok_Str, "Sprawdzaczek", Lok_Str)
'Rs("Company name").Value = InputBox("SprawdŸ poprawnoœæ nazwy firmy: " & xl_wks.Range("D3").Value, "Sprawdzaczek", xl_wks.Range("D3").Value)
'
'Rs("Details").Value = Kod_Str
'Rs("MARKET").Value = xl_wks.Range("D4").Value
'Rs("Segment").Value = xl_wks.Range("D5").Value
'Rs("Sub-segment").Value = xl_wks.Range("D6").Value
'Rs("SUB-SUB-SEGMENT").Value = xl_wks.Range("D7").Value
'Rs("Sales representative").Value = xl_wks.Range("D9").Value
'Rs("Month").Value = CInt(DatePart("m", Date))
'Rs("Year").Value = CInt(DatePart("yyyy", Date))
'Rs("Zarejestrowane przez") = Environ("username")
'Rs("Data utworzenia") = Now()
'Rs("Waste mass produced") = xl_wks.ListObjects("Kalkulacja").ListColumns("Szacunkowa iloœæ odpadu /rok").Total
'Rs("Contract duration") = xl_wks.Range("H4").Value
'Rs("Bids deadline") = xl_wks.Range("H2").Value
'Rs("Inqury date from sales rep") = Now()
'Rs("Result K") = "On-Going"
'Rs("Result H") = "On-Going"
'
'Loc_main = "C:\Users\kbronkow\Desktop\TESTY\Access\Zapytania\" & CInt(DatePart("yyyy", Date)) & "\" & Rs("Company name").Value & "\"
'Name_main = "(" & Rs("Company name").Value & ")" & "(" & Rs("Location").Value & ")" & _
'    "(" & Date & ")" & "(" & Rs("Details").Value & ")"
'
'       Rs("No") = DMax("[No]", "Zapytania") + 1
'        Rs("Sub_No_1") = 1
'        Rs("Sub_No_2") = 1
'
'        Shell "cmd /c md " & """" & Loc_main & "KSO\" & """"
'            Sleep 5000
'        xl_wkb.SaveAs (Loc_main & "KSO\" & "(" & Rs("No").Value & ")" & Name_main & "_" & "KSO" & ".xlsm")
'
'        Rs("Hiper³¹cze") = "#" & Loc_main & "KSO\" & "(" & Rs("No").Value & ")" & Name_main & "_" & "KSO" & ".xlsm"
'
'        Rs("Index_number") = Rs("No") & "_" & Rs("Sub_No_1") & "_" & Rs("Sub_No_2")
'
'Rs("Index_number") = Rs("No") & "_" & Rs("Sub_No_1") & "_" & Rs("Sub_No_2")
'Rs.Update
'
'    xl_app.DisplayAlerts = False
'    xl_wkb.Close
'    xl_app.DisplayAlerts = True
'
'Set xl_wkb = Nothing
'Set xl_app = Nothing
'
'MsgBox ("gotowe")
'Exit Sub
'
'
'End Sub
'
'Public Function Policz(Col As String, Ex As Object, Name As String) As String
'
'Dim Counter_Str As String
'Dim Counter_Tab() As Variant
'Dim arr As New Collection, a
'With Ex.Sheets("KSO_Wycena_odpadów")
'On Error Resume Next
'    For x = 1 To .Range(Col).Rows.Count
'        If .Range(Col).Rows(x).Value <> "" Then _
'        arr.Add .Range(Col).Rows(x).Value, .Range(Col).Rows(x).Value
'    Next x
'
'If arr.Count <= 3 Then
'    For Each a In arr
'        If Counter_Str = "" Then
'            Counter_Str = a
'        Else
'            Counter_Str = Counter_Str & "," & a
'        End If
'    Next a
'Else
'    Counter_Str = arr.Count & " " & Name
'End If
'End With
'
'Policz = Counter_Str
'
'End Function
