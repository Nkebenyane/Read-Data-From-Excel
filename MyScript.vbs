Dim objFSO, objFolder, objFile
Dim objExcel, objWB
Set objExcel = CreateObject("Excel.Application")
Set objFSO = CreateObject("scripting.filesystemobject")
   MyFolder = "C:\OTM"
Set objFolder = objfso.getfolder(myfolder)
For Each objFile In objfolder.Files
   If Right(objFile.Name,4) = "xlsx" Then
      Set objWB = objExcel.Workbooks.Open(objFile)
      objWB.save
      objWB.close
   End If
Next
objExcel.Quit
Set objExcel = Nothing
Set objFSO = Nothing
Wscript.Echo "Done"