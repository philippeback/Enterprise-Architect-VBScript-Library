option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: MDGFolderFind
' Author: Philippe Back
' Purpose: Find where a given MDG technology files lives and get rid of it if wanted
' Date: 13/10/2011
'

sub main
 
 Dim oShell, oEnv, strText
 
 Set oShell = CreateObject("WScript.Shell")
 If oShell Is Nothing then
   MsgBox "Cannot create object - Exiting"
 Else
   '"System", "User", "Volatile", or "Process"
    Set oEnv = oShell.Environment("Process")
  
    strText = oEnv("APPDATA")

    MsgBox("APPDATA is located on: " & strText) 

	Dim objFSO
	Dim objFolder
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")

    Dim strLocation
	
	strLocation = strText & "\Sparx Systems\EA\MDGTechnologies\DWH Technology.xml"
	
    If objFSO.FileExists(strLocation) Then 
       'Set objFolder = objFSO.GetFile(strLocation)
	   Dim res
	   res = MsgBox("File " & strLocation & " found! Erase?", 1)
	   'MsgBox "You clicked on "& res
	   if (res=1) Then 'OK, erase
	     objFSO.DeleteFile(strLocation)
		 MsgBox "Done"
	   end if
    Else
       MsgBox "File " & strLocation & " doesn't exists"
    End If
	
 End If
 
end sub

main