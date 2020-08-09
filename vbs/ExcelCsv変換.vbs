'Specify the directory where the excel files are stored here
pathLen= len(wscript.scriptfullname) - len(wscript.scriptname)
parPath = left(wscript.scriptfullname,pathLen)	
WorkingDir = parPath
Extension1 = ".XLS"
Extension2 = ".XLSX"

Dim fso, myFolder, fileColl, aFile, FileName, SaveName
Dim objExcel,objWorkbook

Set fso = CreateObject("Scripting.FilesystemObject")
Set myFolder = fso.GetFolder(WorkingDir)
Set fileColl = myFolder.Files

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = False
objExcel.DisplayAlerts= False

Dim local
local = true

For Each aFile In fileColl
	ext1 = Right(aFile.Name,4)
	ext2 = Right(aFile.Name,5)
	If UCase(ext1) = UCase(Extension1) OR UCase(ext2) = UCase(Extension2) Then
		'open excel
		FileName = Left(aFile,InStrRev(aFile,"."))
		Set objWorkbook = objExcel.Workbooks.Open(aFile)
		SaveName = FileName & "csv"
		objWorkbook.SaveAs SaveName, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, local
		objWorkbook.Close 
	End If	
Next

Set objWorkbook = Nothing
Set objExcel = Nothing
Set fso = Nothing
Set myFolder = Nothing
Set fileColl = Nothing
