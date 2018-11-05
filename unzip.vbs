'David Tsang
'run unzip.vbs to unzip multiple zip files into unzipped folder

Set oShell = CreateObject("WScript.Shell")
Set ofso = CreateObject("Scripting.FileSystemObject")
oShell.CurrentDirectory = oFSO.GetParentFolderName(Wscript.ScriptFullName)

Dim filesys, demofolder, fil, filecoll, filist 
Dim extName, destFolder
Set filesys = CreateObject("Scripting.FileSystemObject") 
Set demofolder = filesys.GetFolder(oShell.CurrentDirectory)  
Set filecoll = demofolder.Files 

destFolder = oShell.CurrentDirectory & "\unzipped"

For Each fil in filecoll  
	extName = ofso.GetExtensionName(oShell.CurrentDirectory & "\" & fil.Name)
	extName = lcase(extName)
	if extName = "zip" then
		ZipFile= oShell.CurrentDirectory & "\" & fil.Name
		call extract(zipFile, destFolder)
	end if
Next   

Sub extract(zipFile, destFolder)
'26 Dec 2012

'The location of the zip file.

'The folder the contents should be extracted to.
ExtractTo= destFolder

'If the extraction location does not exist create it.
Set fso = CreateObject("Scripting.FileSystemObject")
If NOT fso.FolderExists(ExtractTo) Then
   fso.CreateFolder(ExtractTo)
End If

'Extract the contants of the zip file.
set objShell = CreateObject("Shell.Application")
set FilesInZip=objShell.NameSpace(ZipFile).items
objShell.NameSpace(ExtractTo).CopyHere(FilesInZip)
Set fso = Nothing
Set objShell = Nothing

End sub