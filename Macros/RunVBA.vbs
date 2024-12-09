Dim Input_Excel, Output_Excel, VBA_Path, VBA_Name

Input_Excel 	= WScript.Arguments(0)
Output_Excel 	= WScript.Arguments(1)
VBA_Path 		= WScript.Arguments(2)
VBA_Name	 	= WScript.Arguments(3)

set objxl 		= CreateObject("Excel.Application") 
set objwk 		= objxl.Workbooks.Open(Input_Excel)


objxl.DisplayAlerts = wdAlertsNone
On Error Resume Next
objwk.Application.Run "'" + VBA_Path + "'!" + VBA_Name

If Err.Number <> 0 Then
  'WScript.Echo "You do not have access to the procedure. Please contact tran6!"
  WScript.Echo Err.Description
  Err.Clear
End If


If Input_Excel <> Output_Excel Then objxl.ActiveWorkbook.SaveAs Output_Excel, 51, , , , False
If Input_Excel = Output_Excel Then objxl.ActiveWorkbook.Save

objxl.Quit
Set objxl = nothing
Set objwk = nothing
