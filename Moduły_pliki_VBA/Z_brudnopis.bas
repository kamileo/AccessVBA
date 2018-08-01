Attribute VB_Name = "Z_brudnopis"
Sub TESTOOOR()

DoCmd.OpenForm "Weryfikator", acNormal, , "[INDEX]=1146"

End Sub


Public Sub createNewDirectory(directoryName As String)
     
    If Not DirExists(directoryName) Then
        MkDir (directoryName)
    End If
    
End Sub
 
Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
End Function


Sub Tessst()
Dim directoryName As String
directoryName = "C:\Users\kbronkow\Desktop\Access project 2.0\Zapytania\2017\Tesco\"
    If Not DirExists(directoryName) Then
        MkDir (directoryName)
    End If

End Sub

'Option Compare Database

'Private Sub Polecenie28_Click()
'    Dim cn As New ADODB.Connection
'    Dim Rs As New ADODB.Recordset
'    Dim i, j As Long
'    Dim Users As String
'    Set cn = CurrentProject.Connection
'
'    Set Rs = cn.OpenSchema(adSchemaProviderSpecific, _
'    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")
'    Users = ""
'    While Not Rs.EOF
'        Users = Users & " " & Trim(CStr(Rs.Fields(0)))
'        Rs.MoveNext
'    Wend
'
'Me.Tekst26 = Users
'End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Sample()
    Dim f As Object

    Set f = Application.FileDialog(1)

    f.AllowMultiSelect = False

    If f.Show Then MsgBox f.SelectedItems(1)

End Sub

Public Function Filename2(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        FileName = FileName(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function


End Sub



Sub test()
Dim Location As String
Dim Number As String
Dim z As Integer

Dim xl_app As Object, xlwkb As Object, xl_wks As Object
Dim source_empl_col As Integer

Dim f As Object
Set f = Application.FileDialog(1)
f.AllowMultiSelect = False
If f.Show Then FileName = CStr(f.SelectedItems(1))

Set xl_app = CreateObject("Excel.Application")
Set xl_wkb = xl_app.Workbooks.Open(FileName)
Set xl_wks = xl_wkb.Sheets("KSO_Wycena_odpadów")
Dim Kod_Str As String
Dim Kod_Tab() As Variant
Dim arr As New Collection, a
With xl_wkb.Sheets("KSO_Wycena_odpadów")
For i = 1 To 20 Step 1
    If .Range("B" & i).Interior.ColorIndex = 35 Then
        GoTo Zliczaj
    End If
Next i

Zliczaj:

On Error Resume Next
For z = i + 1 To .Range("B" & .Rows.Count).End(xlDown).Row
    arr.Add .Range("B" & z).Value, .Range("B" & z).Value
    MsgBox .Range("B" & z).Value
Next z

MsgBox arr.Count

If arr.Count <= 5 Then
    For Each a In arr
        If Kod_Str = "" Then
            Kod_Str = a
        Else
            Kod_Str = Kod_Str & "," & a
        End If
    Next a
Else
    Kod_Str = "wiele kodów"
End If
End With

End Sub



'For z = i + 1 To .Range("B" & .Rows.Count).End(xlDown).Row
'For i = 1 To ws.Range("C" & ws.Rows.Count).End(xlUp).Row
'Dim Users As String



Sub ShowUserRosterMultipleUsers()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i, j As Long

    Set cn = CurrentProject.Connection

    ' The user roster is exposed as a provider-specific schema rowset
    ' in the Jet 4.0 OLE DB provider.  You have to use a GUID to
    ' reference the schema, as provider-specific schemas are not
    ' listed in ADO's type library for schema rowsets

    Set rs = cn.OpenSchema(adSchemaProviderSpecific, _
    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")

    'Output the list of all users in the current database.

'    Debug.Print rs.Fields(0).Name, "", rs.Fields(1).Name, _
'    "", rs.Fields(2).Name, rs.Fields(3).Name
'
'    While Not rs.EOF
'        Debug.Print rs.Fields(0), rs.Fields(1), _
'        rs.Fields(2), rs.Fields(3)
'        rs.MoveNext
'    Wend

    While Not rs.EOF
        Users = CStr(rs.Fields(0)) & " "
        rs.MoveNext
    Wend
    MsgBox Users

End Sub



'        Shell "cmd /c md " & """" & Loc_main & "KSO\" & """"
'            Sleep 5000
'        xl_wkb.SaveAs (Loc_main & "KSO\" & "(" & Rs("No").Value & ")" & Name_main & "_" & Rs("Sub_No_1") & "ver" & Rs("Sub_No_2") & ".xlsm")
'
    
'Zapisywanie pliku
    'Location = "C:\Users\kbronkow\Desktop\Access project 2.0\Zapytania\" & _
    'CInt(DatePart("yyyy", Date)) & "\" & Rs("Company name").Value & "\"


    'Number = CStr(DMax("[No]", "Zapytania") + 1)
    'Location = "C:\Users\kbronkow\Desktop\Access project 2.0\Zapytania\" & Rs("Company name").Value & "-" & CStr(xl_wks.Range("C3").Value) & "(" & Lok_Str & ";" & Kod_Str & ")"
    'Call createNewDirectory(Location)
    
    'Shell "cmd /c md " & """" & Location & """"
  ''''''''''''''''''''''''''''''''''''''''''
    
'    xl_wkb.SaveAs (Location & _
'    "(" & Rs("No").Value & ")" & _
'    "(" & Rs("Company name").Value & ")" & _
'    "(" & Rs("Location").Value & ")" & _
'    "(" & Date & ")" & _
'    "(" & Rs("Details").Value & ")" & "_" & _
'    Rs("Sub_No_1") & "ver" & _
'    Rs("Sub_No_2") & _
'    ".xlsm")


Public Function Policz(Col As String, ex As Object, Name As String) As String

Dim Counter_Str As String
Dim Counter_Tab() As Variant
Dim arr As New Collection, a
With ex.Sheets("KSO_Wycena_odpadów")
On Error Resume Next
    For x = 1 To .Range(Col).Rows.Count
        If .Range(Col).Rows(x).Value <> "" Then _
        arr.Add .Range(Col).Rows(x).Value, .Range(Col).Rows(x).Value
        'MsgBox .Range(Col).Rows(x).Value
        'arr.Add .Range("B" & z).Value, .Range("B" & z).Value
    Next x

If arr.Count <= 3 Then
    For Each a In arr
        If Counter_Str = "" Then
            Counter_Str = a
        Else
            Counter_Str = Counter_Str & "," & a
        End If
    Next a
Else
    Counter_Str = arr.Count & " " & Name
End If
End With

Policz = Counter_Str

End Function


        'xl.Visible = True
        'xl_wkb.Application.Run ("!Generuj_Handlowiec")
        
        '''''XL.Run "congrats_macro.congrats"

        
        'xl_app.EnableEvents = True
        'xl_wks.Application.Run ("!Generuj_Handlowiec")
        'xl_app.EnableEvents = False
        
        'l.Run xl_wkb.Name & "!Generuj_Handlowiec"
        'xl_wkb.Application.Run ("!Generuj_Handlowiec")
        
        
        'MsgBox xl_wkb.Name
        'xl_wkb.Application.Run ("!Generuj_Handlowiec")
        'MsgBox ("'" & xl_wkb.Name & "'!Generuj_Handlowiec")
       ' Application.Run ("'" & xl_wkb.Name & "'!Generuj_Handlowiec")




'Set f = Application.FileDialog(1)
'f.AllowMultiSelect = False
'If f.Show Then FileName = CStr(f.SelectedItems(1))
'Set xl_app = CreateObject("Excel.Application")
'xl_app.EnableEvents = False
'Set xl_wkb = xl_app.Workbooks.Open(FileName)
'Set xl_wks = xl_wkb.Sheets("KSO_Wycena_odpadów")
'Set xl_wks_p = xl_wkb.Sheets("Podsumowanie na konrakt")
