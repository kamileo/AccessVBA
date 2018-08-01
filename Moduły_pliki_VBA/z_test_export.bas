Attribute VB_Name = "z_test_export"
Sub sdgsdg()

MsgBox Date


End Sub



Public Sub test_0()

    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i, j As Long
    Dim Users As String
    Set cn = CurrentProject.Connection

    Set rs = cn.OpenSchema(adSchemaProviderSpecific, _
    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")
    Users = ""
    While Not rs.EOF
        
        MsgBox Trim(CStr(rs.Fields(0)))
            
 
        rs.MoveNext
    Wend
    
End Sub




'Option Compare Database
'
'Sub TestOutputAndCopyWorksheet()
'Dim sPath As String
'sPath = "C:\Users\kbronkow\Downloads\Test.xls"
'DoCmd.OutputTo acOutputReport, "rptBook", acFormatXLS, sPath, False
'
'Dim oExcel As Object, oWb1 As Object, oWS1 As Object
'Set oExcel = New Excel.Application
'
'
'Set oWb1 = oExcel.Workbooks.Open(sPath)
'Set oWS1 = oWb1.Worksheets("rptBook")
'
'Dim oWB2 As Excel.Workbook
'Set oWB2 = oExcel.Workbooks.Add()
'
'oWS1.Copy Before:=oWB2.Sheets(1)
'
'oExcel.DisplayAlerts = False
'
'oWb1.Save
'oWB2.SaveAs "C:\Users\kbronkow\Downloads\TestTwo.xlsx", ConflictResolution:=xlLocalSessionChanges
'oExcel.DisplayAlerts = True
'
'oWb1.Close
'oWB2.Close
'oExcel.Quit
'
'Set oWb1 = Nothing
'Set oWB2 = Nothing
'Set oWS1 = Nothing
'
'Set oExcel = Nothing
'End Sub
