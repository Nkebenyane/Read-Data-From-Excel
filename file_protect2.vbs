Dim objFSO, objFolder, objFile
Dim objExcel, objWB, stdout
Set objExcel = CreateObject("Excel.Application")
Set objFSO = CreateObject("scripting.filesystemobject")
Set stdout = objFSO.GetStandardStream (1)
Wscript.Echo "Start"
objStartFolder = "C:\OTM"

Set objFolder = objFSO.GetFolder(objStartFolder)

' stdout.WriteLine objFolder.Path

Set colFiles = objFolder.Files

For Each objFile in colFiles
	If (Right(objFile.Name,4) = "xlsx") and  (Left(objFile.Name,1) <> "~") Then
		Set objWB = objExcel.Workbooks.Open(objFile)
		objWB.save
		objWB.close
	End If
Next
ShowSubfolders objFSO.GetFolder(objStartFolder)

objExcel.Quit
Set objExcel = Nothing
Set objFSO = Nothing
Wscript.Echo "Done"


Sub ShowSubFolders(Folder)
	Set stdout = objFSO.GetStandardStream (1)
    For Each Subfolder in Folder.SubFolders
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
			If Right(objFile.Name,4) = "xlsx" and  (Left(objFile.Name,1) <> "~")  Then
				' stdout.WriteLine WsobjFile
				Set objWB = objExcel.Workbooks.Open(objFile)
				objWB.save
				objWB.close
			End If
        Next
        ShowSubFolders Subfolder
    Next
End Sub