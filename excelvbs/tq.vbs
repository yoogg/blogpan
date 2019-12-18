Set objExcel = CreateObject("Excel.Application")
On Error Resume Next  '错误继续运行
objExcel.Visible = false  '隐藏excel窗口
Set objWorkbook = objExcel.Workbooks.Open ("D:\vbs\天气.xlsx") 
objWorkbook.RefreshAll  '更新
objExcel.DisplayAlerts = False  '关闭提示
objWorkbook.Save '保存工作表
objWorkbook.Close False '关闭工作表

objExcel.Quit  ' 退出 
Set objWorkbook = Nothing
Set objExcel = Nothing


Set ObjShell=CreateObject("Wscript.Shell")
SttCommand=("cmd.exe /C  Taskkill  /f /im Excel.exe")
ObjShell.Run SttCommand, 0, False

