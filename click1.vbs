Option Explicit

On Error Resume Next

ExcelMacro_Jikkou

Sub ExcelMacro_Jikkou() 

	Dim xlApp 
	Dim xlBook
	Dim cur_path 

	dim fso
	set fso = createObject("Scripting.FileSystemObject")
	cur_path=fso.getParentFolderName(WScript.ScriptFullName)
	set fso = nothing

	Set xlApp = CreateObject("Excel.Application") 
	Set xlBook = xlApp.Workbooks.Open(cur_path & "\click1.xlsm", 0, True) 
	xlApp.Run "click"
	xlBook.close(false)
	xlApp.Quit 
	
	Set xlBook = Nothing 
	Set xlApp = Nothing 

End Sub
