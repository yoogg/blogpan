Set objExcel = CreateObject("Excel.Application")
On Error Resume Next  '�����������
objExcel.Visible = false  '����excel����
Set objWorkbook = objExcel.Workbooks.Open ("D:\vbs\����.xlsx") 
objWorkbook.RefreshAll  '����
objExcel.DisplayAlerts = False  '�ر���ʾ
objWorkbook.Save '���湤����
objWorkbook.Close False '�رչ�����

objExcel.Quit  ' �˳� 
Set objWorkbook = Nothing
Set objExcel = Nothing


Set ObjShell=CreateObject("Wscript.Shell")
SttCommand=("cmd.exe /C  Taskkill  /f /im Excel.exe")
ObjShell.Run SttCommand, 0, False

