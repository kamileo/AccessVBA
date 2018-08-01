Attribute VB_Name = "Wybór_arkszua_tabeli"
'Public lista_ark As String
'Public shtpick As String

Option Compare Database

Sub Get_Data_Cleangdfg()

'×××××××××××××××××××××××××××××××××××××DEKLARACJA ZMIENNYCH×××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
Dim Location, Number, nazwa_p, Hp As String
Dim Path_Main As String: Path_Main = "C:\Users\kbronkow\Desktop\TEST ACCESS\"
Dim objAtt As Outlook.Attachment
Set objOL = CreateObject("Outlook.Application")
Set objNS = objOL.GetNamespace("MAPI")
Dim oMail As Outlook.MailItem
Dim xl_app As Object, xl_wkb As Object, xl_wks As Object, f, objItem As Object
Set xl_app = CreateObject("Excel.Application")
xl_app.EnableEvents = False
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'××××××××××××××××××××××××××××××××××××PRZYPISYWANIE SKOROSZYTU I ARKUSZY××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××××
FileName = "C:\Users\kbronkow\Desktop\TestPob.xlsm"
Set xl_wkb = xl_app.Workbooks.Open(FileName)
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Set xl_wks = xl_wkb.Sheets(arkusze("KSO1", xl_wkb))
'MsgBox xl_wks.Range("A1").Value

MsgBox Tabela(xl_wks, "Kalkulacja")

'Debug.Print tbl.Name

End Sub


'Public Function Tabela(ByRef sht As Object, searchtb As String)
'Dim tbl As ListObject
'    For Each tbl In sht.ListObjects
'    If InStr(1, tbl.Name, searchtb, vbTextCompare) > 0 Then
'        Tabela = tbl.Name
'        GoTo FINITO
'    End If
'    Next tbl
'FINITO:
'End Function


'For i = 1 To xl_app.Sheets.Count
'    If xl_app.Sheets(i).Name = "KSO3" Then
'        Lista = ""
'        shtpick = "KSO1"
'        GoTo MAMYTO:
'    Else
'        Debug.Print xl_app.Sheets(i).Name
'        Lista = Lista & Chr(34) & xl_app.Sheets(i).Name & Chr(34) & ";"
'    End If
'Next
'    Lista = Left(Lista, Len(Lista) - 1)
'    lista_ark = Lista
'DoCmd.OpenForm "Form Picker", WindowMode:=acDialog
'MAMYTO:

'ark = arkusze("KSO1", xl_wkb)

'Set xl_wkt = xl_wkb.Sheets("Tabela informacyjna")
'Set xl_wks_p = xl_wkb.Sheets("Podsumowanie na konrakt")
'Set xl_wkc = xl_wkb.Sheets("Cennik do oferty")

