option explicit

!INC Local Scripts.EAConstants-VBScript
!INC EAScriptLib.VBScript-Logging


'
' Script Name: MDGRemoveXMLInAppData
' Author: Philippe Back
' Purpose: Remove XML files from the %APPDATA%\Sparx Systems\EA\MDGTechnologies\*.xml
' Date: 18/10/2011
'

sub main
 
 LOGLEVEL=LOGLEVEL_DEBUG
 
 Dim oShell, oEnv, strText
  
 Set oShell = CreateObject("WScript.Shell")
 If oShell Is Nothing then
   MsgBox "Cannot create object - Exiting"
 Else
   '"System", "User", "Volatile", or "Process"
    Set oEnv = oShell.Environment("Process")
  
    strText = oEnv("APPDATA")

    LOGInfo "APPDATA is located on: " & strText

	Dim objFSO
	Dim objFolder
	Dim objFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")

    Dim strLocation
	
	strLocation = strText & "\Sparx Systems\EA\MDGTechnologies\"
	
    If objFSO.FolderExists(strLocation) Then 
       Set objFolder = objFSO.GetFolder(strLocation)
	   Dim res
	   res = MsgBox("Folder " & strLocation & " found! Erase XML?", 1)
	   LOGDebug "You clicked on "& res
	   if (res=1) Then 'OK, erase
	     For each objFile In objFolder.Files
		   LOGInfo "Found: " & objFile.Name
		   objFSO.DeleteFile(strLocation & objFile.Name)
		   LOGInfo "Removed"
		 Next
         MsgBox "Done"
	   end if
    Else
       MsgBox "Folder " & strLocation & " doesn't exists"
    End If
	
 End If
 
end sub

main