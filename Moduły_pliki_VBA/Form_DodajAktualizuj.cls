VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Dodaj/Aktualizuj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Obraz109_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_UPHP(1, "edycja", "wybór", "Kalkulacja")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub

Private Sub Obraz144_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_UPHP(1, "edycja", "wybór", "KSO")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub

Private Sub Record_Based_KSO_h_Click()

    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(2, "nowy", "Hp", "KSO")
        'Forms("Dodaj/Aktualizuj").Requery
    End If

End Sub

Private Sub Rejestracja_z_m_2_h_Click()

    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(3, "nowy", "Hp", "Kalkulacja")
        'Forms("Dodaj/Aktualizuj").Requery
    End If

End Sub

Private Sub test_Click()
    Set rs = Forms("Dodaj/Aktualizuj").Recordset
    SQL_r = "[INDEX]=" & rs![INDEX]
    DoCmd.Close acForm, "Dodaj/Aktualizuj"
    DoCmd.OpenForm "Weryfikator", acNormal, , SQL_r

End Sub

Private Sub Wyceny_mail_old_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(9, "edycja", "mail", "Kalkulacja")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub

Private Sub wyceny_old_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(9, "edycja", "wybór", "Kalkulacja")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub

Private Sub Rêcznie_Mail_Click()
    Call Get_Data_Clean(8, "nowy", "mail", "KSO")
End Sub

Private Sub Rêcznie_Click()

    Call sendOutlookEmail(1, "oferta")
    'Call Get_Data_Clean(8, "nowy", "wybór", "KSO")
        
End Sub

'===================================================================================================

Private Sub Rejestracja_KSO_Click()

    Call Get_Data_Clean(1, "nowy", "wybór", "KSO")
    'Forms("Dodaj/Aktualizuj").Requery
End Sub

Private Sub Record_Based_KSO_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(2, "nowy", "wybór", "KSO")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub

Private Sub Record_Based_KSO_m_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(2, "nowy", "mail", "KSO")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub

Private Sub Rejestracja_z_m_2_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(3, "nowy", "wybór", "Kalkulacja")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub

Private Sub Rejestracja_z_m_2_m_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(3, "nowy", "mail", "Kalkulacja")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub

Private Sub wyceny_back_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(4, "edycja", "wybór", "Kalkulacja")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Plik_realizacja_Click()
    Call sendOutlookEmail(1, "akceptacja")
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub oferta_handlowa_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(5, "edycja", "Hp", "ND")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub

Private Sub KSO_mail_Click()
        Call Get_Data_Clean(1, "nowy", "mail", "KSO")
        'Forms("Dodaj/Aktualizuj").Requery
End Sub

Private Sub wyceny_back_mail_Click()
    If IsNull(kso_up) = True Then
        MsgBox ("Wybierz rekord")
        Exit Sub
    Else
        Call Get_Data_Clean(4, "edycja", "mail", "Kalkulacja")
        'Forms("Dodaj/Aktualizuj").Requery
    End If
End Sub

'-------------------------------------------------------------------------------------------------------------------

Private Sub kso_up_Change()
    p_firma.Value = Me.Recordset("Company name")
    p_lok.Value = Me.Recordset("Location")
    p_details.Value = Me.Recordset("Details")
    
    p_no.Value = Me.Recordset("No")
    p_subno1.Value = Me.Recordset("Sub_No_1")
    p_subno2.Value = Me.Recordset("Sub_No_2")
    
End Sub

Private Sub Form_Load()

obecni_nieobecni
    
End Sub

Private Sub Obraz47_Click()

obecni_nieobecni
    
End Sub

Public Sub obecni_nieobecni()

an.Visible = False
an_o.Visible = False
gd.Visible = False
gd_o.Visible = False
db.Visible = False
db_o.Visible = False
pa.Visible = False
pa_o.Visible = False
kb.Visible = False
kb_o.Visible = False
'ki.Visible = False
'ki_o.Visible = False

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i, j As Long
    Dim Users As String
    Set cn = CurrentProject.Connection

    Set rs = cn.OpenSchema(adSchemaProviderSpecific, _
    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")
    Users = ""
    While Not rs.EOF
        
        Select Case Trim(CStr(rs.Fields(0)))
        
            Case "WAWL1169"
                an.Visible = True
                an_o.Visible = True
            Case "WAWD0163"
                gd.Visible = True
                gd_o.Visible = True
            Case "WAWL1214"
                pa.Visible = True
                pa_o.Visible = True
            Case "WAWL1194"
                db.Visible = True
                db_o.Visible = True
            Case "WAWL0262"
                kb.Visible = True
                kb_o.Visible = True
        
        End Select
        rs.MoveNext
    Wend
    
End Sub
