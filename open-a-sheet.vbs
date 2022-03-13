file_path = Wscript.Arguments.Item(0)
sheet_name = Wscript.Arguments.Item(1)
Dim Xl
Dim Wb

On Error Resume Next
set xl = GetObject(, "Excel.Application")
If Err.Number <> 0 then
Set Xl = CreateObject("Excel.Application")
xl.Visible = True
End If
On Error Goto 0

Set Wb = xl.Workbooks.Open(file_path)
wb.Sheets(sheet_name).Activate

