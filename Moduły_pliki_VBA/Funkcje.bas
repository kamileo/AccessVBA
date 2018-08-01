Attribute VB_Name = "Funkcje"
Option Compare Database

Public Function Policz(Col As String, ByRef sht As Object, Name As String) As String

Dim Counter_Str As String
Dim Counter_Tab() As Variant
Dim arr As New Collection, a
With sht
On Error Resume Next
    For x = 1 To .Range(Col).Rows.Count
        If .Range(Col).Rows(x).Value <> "" And .Range(Col).Rows(x).Value <> 0 Then
        g = Replace(.Range(Col).Rows(x).Value, "*", "o", 1, , vbTextCompare)
        arr.Add g, g
        End If
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
