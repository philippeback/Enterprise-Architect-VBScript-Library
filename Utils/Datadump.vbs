!INC Local Scripts.EAConstants-VBScript
!INC EAScriptLib.VBScript-Logging

'
' Script Name: DataDump
' Author: Philippe Back


' Purpose: Dump a collection of collections for debugging purposes
' Date: 25/10/2011
Sub DataDump(title, oCollectionOfCollections)
  LOGInfo("Dumping:" & title)
  
  Dim oCollection, oItem
  
  For Each oCollection in oCollectionOfCollections
     LOGInfo("ROW")
	 For Each oItem in oCollection
	   LOGInfo("Item:" & CStr(oItem))
	 Next
  Next
End Sub

' Purpose: Dump a collection  for debugging purposes
' Date: 25/10/2011
Sub CollectionDump(title, oCollection)
  LOGInfo("Dumping:" & title)
  
  Dim oItem
  
   For Each oItem in oCollection
	   LOGInfo(oItem)
	 Next
  
End Sub

Sub CollectionOfElementsDump(title, oCollection)
  LOGInfo("Dumping Collection of Elements:" & title)
  
  Dim currentElement as EA.Element
  
   For Each currentElement in oCollection
	   LOGInfo("Element:" & currentElement.Name & _
			" (" & currentElement.Type & _
			", ID=" & currentElement.ElementID & ")")
	 Next
  
End Sub

Sub LOGElement(theElement) 
	LOGInfo("Element:" & theElement.Name & "["& theElement.ElementID &"] (" & theElement.Type & ")")
End Sub

Sub LOGPackage(thePackage) 
	LOGInfo("Package:" & thePackage.Name & "[" & thePackage.PackageID & "]")
End Sub


Sub ShowScriptOutputWindow
	Repository.EnsureOutputVisible "Script"
	Repository.ClearOutput "Script"
End Sub