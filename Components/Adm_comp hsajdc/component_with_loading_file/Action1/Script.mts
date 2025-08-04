Option Explicit

Dim fileT, resource

'Set resource = GetResource(Trim("C:\res\ActiveWorkspace_OR.xml"))

'Set fileT=Eval(GetResource("C:\res\ActiveWorkspace_OR.xml").GetValue("Browser"))
Dim savePath
savePath = "Start: "

Dim testPath
testPath =Environment("TestDir")
Reporter.ReportEvent micDone, "Starting folder", testPath
Dim parts, i, partialPath
Dim fso, folder, subfolder, subfolders

Set fso = CreateObject("Scripting.FileSystemObject")

parts = Split(testPath, "\")

partialPath = parts(0)
For i = 1 To UBound(parts) - 3
	partialPath = partialPath & "\" & parts(i)	
Next

For i = UBound(parts) - 2 To UBound(parts)
	partialPath = partialPath & "\" & parts(i)
	Reporter.ReportEvent micDone, "Folder", partialPath
	Set folder = fso.GetFolder(partialPath)
	Set subfolders = folder.SubFolders
	For Each subfolder In subfolders
		Reporter.ReportEvent micDone, "Subfolder", subfolder.Path
	Next
Next


Reporter.ReportEvent micPass, savePath, "aaaa"
