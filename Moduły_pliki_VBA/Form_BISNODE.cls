VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_BISNODE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Szukaj_h_Click()

'lista_s
'rs("Client type")
'rs("Company name")
'Dir("O:\Klient Sieciowy\KALKULACJE\2018\Atalian\" & "*" & ".pdf")


Dim rs As DAO.Recordset
main_p = "\\WAWFS01\Groups$\All\Klient Sieciowy\KALKULACJE\2018\"
SQL_q = "SELECT * FROM Zapytania WHERE ([Client type] = ""new"" AND [Year] = 2018)"
Set rs = CurrentDb.OpenRecordset(SQL_q, dbOpenDynaset, dbSeeChanges)

'Set rs = CurrentDb.OpenRecordset(SQL_q)
'Set rs = db.OpenRecordset("Zapytania", dbOpenDynaset, dbSeeChanges)


If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    Do Until rs.EOF = True
        If Dir(main_p & rs("Company name") & "\" & "*" & ".pdf") = "" Then
           lista_s = lista_s & vbCrLf & rs("Company name")
        End If
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing
End Sub
