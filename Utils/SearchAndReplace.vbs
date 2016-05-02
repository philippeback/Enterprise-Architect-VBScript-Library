option explicit

!INC Local Scripts.EAConstants-VBScript

Dim fromString, toString

'Replace text in Name and Notes
 Function ReplaceString(obj)

    Dim modified
    modified = False

    'replace string in Name
    If InStr(obj.Name, fromString) > 0 Then
        obj.Name = Replace(obj.Name, fromString, toString)
        modified = True
    End If
   
    'replace string in Notes
    If InStr(obj.Notes, fromString) > 0 Then
        obj.Notes = Replace(obj.Notes, fromString, toString)
        modified = True
    End If
   
    'call Update if modified
    If modified Then
        obj.Update
    End If
   
    ReplaceString = modified
   
End Function

'Replace texts of an element
Sub ReplaceElement(elem)

    Dim method
    Dim attr
    Dim i
    Dim modified
    modified = False
   
    ' element methods
 For i = 0 To elem.Methods.Count - 1
  Set method = elem.Methods.GetAt(i)
    
  modified = modified Or ReplaceString(method)
 Next
           
    ' element attributes
 For i = 0 To elem.Attributes.Count - 1
  Set attr = elem.Attributes.GetAt(i)
   
  modified = modified Or ReplaceString(attr)
 Next

    ' element properties
 modified = modified Or ReplaceString(elem)
   
 ' tell EA changes if modified
    If modified Then
        Repository.AdviseElementChange elem.ElementID
    End If
       
       
End Sub

' Replace texts of a Package
Sub ReplaceInPackage(pkg)

    ' elements in the Package
    Dim elem
    Dim i
    For i = 0 To pkg.Elements.Count - 1
        Set elem = pkg.Elements.GetAt(i)
               
        ReplaceElement elem
    Next
    
    ' packages in the Package
    Dim childPkg
   
 For i = 0 To pkg.Packages.Count - 1
  Set childPkg = pkg.Packages.GetAt(i)
  
  ReplaceInPackage childPkg
 Next
   
    ' Package properties
 If ReplaceString(pkg) Then
  ' tell EA changes if modified
  Repository.AdviseElementChange pkg.Element.ElementID
 End If

End Sub

'Main function
Sub Main()

 fromString = InputBox("Find what:","Replace text")
 If fromString = "" Then Exit Sub
 toString = InputBox("Replace with:","Replace text")
 If toString = "" Then Exit Sub

 ReplaceInPackage Repository.GetTreeSelectedPackage
 
 Msgbox "Complete.",,"Replace text"

End Sub

'Call main function
Main