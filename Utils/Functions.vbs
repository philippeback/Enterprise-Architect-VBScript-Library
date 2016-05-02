!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
' Returns a replaced string: Pattern, Replacement, Source
Function ReplaceText(byVal patrn, byVal replStr, byVal textString)
	Dim regEx, Matches, Match           ' Create variables.
	ReplaceText = textString
	Set regEx = New RegExp            ' Create regular expression.
	regEx.Pattern = patrn            ' Set pattern.
	'regEx.IgnoreCase = True            ' Make case insensitive.
	regEx.Global = True         ' Set global applicability.

	Set Matches = regEx.Execute(textString)
	For Each Match in Matches
		ReplaceText = regEx.Replace(textString, replStr)   ' Make replacement.
	Next
End Function

Function FileExists(filespec)
    
    FileExists = False
    
    Dim fso
    Set fso=CreateObject("Scripting.FileSystemObject")
    
    'On Error Resume Next

    If fso.FileExists(filespec) then
      FileExists  = True
    End If
    
    Set fso=Nothing
End Function

