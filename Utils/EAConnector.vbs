!INC Local Scripts.EAConstants-VBScript

'Repository is a globally available object pointing to the EA model currently opened.

Class EAConnector
'-------------------------------------------------------------
' Author:   Geert Bellekens
' Date:     10/10/2008
' Description: set some perfomance enhancing stuff
'-------------------------------------------------------------
Public Sub speedUp()
    Repository.BatchAppend = True
    Repository.EnableUIUpdates = False
End Sub

'-------------------------------------------------------------
' Author:   Geert Bellekens
' Date:     01/09/2009
' Description: refresh the modelView
'-------------------------------------------------------------
Public Sub refreshModelView(package) ' As EA.package)
 Repository.refreshModelView (package.PackageID)
End Sub


'-------------------------------------------------------------
' Author:   Geert Bellekens
' Date:     31/08/2009
' Description: adds or updates the class with the given name
'   stereotype and description in the given parent Package
'-------------------------------------------------------------
Public Function addOrUpdateClass(parentPackage, name, stereotype, description) 
'(parentPackage As EA.package, name As String, stereotype As String, description As String) As EA.Element
    Dim myClass As EA.Element
    'try to find existing class with the given name
    'this will only work correctly if there is only one element with the given name, and if it is a class
    Set myClass = getElementByName(parentPackage, name)
    If myClass Is Nothing Then
        'no existing class, create new
        Set myClass = parentPackage.Elements.AddNew(name, "Class")
    End If
    'set properties
    myClass.stereotype = stereotype
    myClass.Notes = description
    'save class
    myClass.Update
    'refresh elements collection
    parentPackage.Elements.Refresh
    'return class
    Set addOrUpdateClass = myClass
End Function

'-------------------------------------------------------------
' Author:   Geert Bellekens
' Date:     31/08/2009
' Description: adds or updates the attribute with the given name
'   stereotype and description and type in the given parent Class
'-------------------------------------------------------------
Public Function addOrUpdateAttribute(p_parentClass, name, stereotype, description, attrType) 
'parentClass As EA.Element, name As String, stereotype As String, description As String, attrType As String) As EA.Attribute
  Dim myAttribute As EA.Attribute
    'try to find existing attribute with the given name
  Dim parentClass As EA.Element
    
	Set parentClass = p_ParentClass
	LOGInfo("aoua->ParentClass=" & parentClass.Name & " AttName=" & name)
    Set myAttribute = getAttributeByName(parentClass, name)
    If myAttribute Is Nothing Then
        'no existing attribute, create new
        Set myAttribute = parentClass.Attributes.AddNew(name, "Attribute")
    End If
    'set properties
    myAttribute.stereotype = stereotype
    myAttribute.Notes = description
    myAttribute.Type = attrType
    'save attribute
    myAttribute.Update
    'refresh attributes collection
    parentClass.Attributes.Refresh
    'return attribute
    Set addOrUpdateAttribute = myAttribute
End Function
'-------------------------------------------------------------
' Author:   Geert Bellekens
' Date:     07/09/2009
' Description: adds or updates the attribute tag on the given attribute with
' the given name and value
'-------------------------------------------------------------
Public Function addOrUpdateAttributeTag(anAttribute, tagName, tagValue)
'anAttribute As EA.Attribute, tagName As String, tagValue As String) As EA.AttributeTag
    Dim currentAttributeTag 'As EA.AttributeTag
	Set addOrUpdateAttributeTag = Nothing
	On Error Resume Next
    'update all tagged values with the given name
    For Each currentAttributeTag In anAttribute.TaggedValues
        If currentAttributeTag.name = tagName Then
            currentAttributeTag.Value = tagValue
            Set addOrUpdateAttributeTag = currentAttributeTag
            currentAttributeTag.Update
        End If
    Next
	If Err.Number <> 0 Then
       Err.Clear
	   Set addOrUpdateAttributeTag = Nothing
    End If
    On Error Goto 0
	
    'no tagged value found, so create it
    If addOrUpdateAttributeTag Is Nothing Then
        Set addOrUpdateAttributeTag = anAttribute.TaggedValues.AddNew(tagName, "AttributeTag")
        addOrUpdateAttributeTag.Value = tagValue
        addOrUpdateAttributeTag.Update
    End If
End Function


Public Function addOrUpdateTag(theElement, tagName, tagValue)
'anElement As EA.Element, tagName As String, tagValue As String

	Dim currentTag as EA.TaggedValue
	
    Dim tags as EA.Collection
	Set tags = theElement.TaggedValues
		
	
	Set addOrUpdateTag = Nothing
	On Error Resume Next
    'update all tagged values with the given name
    For Each currentTag In theElement.TaggedValues
        If currentTag.name = tagName Then
            currentTag.Value = tagValue
            Set addOrUpdateTag = currentTag
            currentTag.Update
        End If
    Next
	If Err.Number <> 0 Then
       Err.Clear
	   Set addOrUpdateAttributeTag = Nothing
    End If
    On Error Goto 0
	
    'no tagged value found, so create it
    If addOrUpdateTag Is Nothing Then
        Set addOrUpdateTag = theElement.TaggedValues.AddNew(tagName, tagValue)
        addOrUpdateTag.Update
    End If
End Function
		

'-------------------------------------------------------------
' Author:   Geert Bellekens
' Date:     31/08/2009
' Description: gets the attribute with the given name from the given class
'-------------------------------------------------------------
Public Function getAttributeByName(parentClass, name)
'(parentClass As EA.Element, name As String) As EA.Attribute
    LOGInfo("getAttributeByName")
    Dim currentAttribute As EA.Attribute
	'LOGInfo("getAttributeByName2")
	Set getAttributeByName = Nothing
	'LOGInfo("getAttributeByName3")
	On Error Resume Next
	'LOGInfo("getAttributeByName4")
    For Each currentAttribute In parentClass.Attributes
	'LOGInfo("getAttributeByName5")
        If currentAttribute.name = name Then
		'LOGInfo("getAttributeByName6")
            Set getAttributeByName = currentAttribute
            Exit For
        End If
    Next
	'LOGInfo("getAttributeByName7")
	If Err.Number <> 0 Then
	'LOGInfo("getAttributeByName8")
       Err.Clear
	   Set getAttributeByName = Nothing
	Else
	  'LOGInfo("getAttributeByName8b")
    End If
	'LOGInfo("getAttributeByName9")
	On Error Goto 0
End Function
'-------------------------------------------------------------
' Author:   Geert Bellekens
' Date:     31/08/2009
' Description: gets the element with the given name from the given package
'-------------------------------------------------------------
Public Function getElementByName(parentPackage, name)
'(parentPackage As EA.package, name As String) As EA.Element
    On Error Resume Next
    Set getElementByName = parentPackage.Elements.GetByName(name)
    If Err.Number <> 0 Then
       Err.Clear
	   Set getElementByName = Nothing
    End If
    On Error Goto 0
End Function

'-------------------------------------------------------------
' Author:   Geert Bellekens
' Date:     17/12/2007
' Description: Return the selected package from the currently opened model
'-------------------------------------------------------------


Public Function getSelectedPackage() 'As EA.package
	
	' Get the currently selected package in the tree to work on
	dim thePackage as EA.Package
	set thePackage = Repository.GetTreeSelectedPackage()
		
	if not thePackage is nothing and thePackage.ParentID <> 0 then
		
		LOGInfo( "Working on package '" & thePackage.Name & "' (ID=" & _
			thePackage.PackageID & ")" )
		Set getSelectedPackage = thePackage
	else
	   LOGInfo("No package selected")
	end if
	
End Function

'-------------------------------------------------------------

'-------------------------------------------------------------
' Author:   Geert Bellekens
' Date:     14/01/2008
' Description: gets the classes for the given package
'-------------------------------------------------------------
Public Function getClasses(package) 'As Collection
Dim i 'As Integer
Dim allElements As EA.Collection
Dim aClass As EA.Element
'initialize return
Set getClasses = CreateObject("System.Collections.ArrayList")
'first get all elements
Set allElements = package.Elements
For i = 0 To allElements.Count - 1
    Set aClass = allElements.GetAt(i)
    'if the element is a class (and not an enumeration) add to the output collection
    If aClass.Type = "Class" And aClass.stereotype <> "enumeration" Then
        getClasses.Add aClass
    End If
Next

End Function
 

'relationshipType is one of:
'Aggregation Assembly Association Collaboration CommunicationPath Connector ControlFlow 
'Delegate Dependency Deployment ERLink Generalization InformationFlow Instantiation 
'InterruptFlow Manifest Nesting NoteLink ObjectFlow Package Realization Sequence 
'StateFlow UseCase 
Function CreateRelationship(sourceParentPackage, sourceElementName, _
                            targetParentPackage, targetElementName, _
							relationshipName, relationshipType)
	' Find source element by name
	' Find target element by name
	' Create a relationship between the two of the specified type and using the supplied stereotype
	LOGInfo("CreateRelationship")
	
	Dim con As EA.Connector
	
	Dim srcID
    srcID	= findElementIDByName(sourceParentPackage, sourceElementName)
	Dim trgID 
	trgID   = findElementIDByName(targetParentPackage, targetElementName)
	
	LOGInfo("src:" & srcID & " - trg:" & trgID)

	Set con = Repository.GetElementByID(srcID).Connectors.AddNew(relationshipName, relationshipType)
    con.SupplierID = trgID

   if (not con.Update) Then
       LOGError("Connector: (" & sourceElementName & "->" & targetElementName & ") - Fail: " & con.GetLastError)
	   Set CreateRelationship = Nothing
   Else
       LOGInfo("Connector: [" & relationshipType & "] (" & sourceElementName & "-- " & relationshipName & "-->" & targetElementName & ") created")
       Set CreateRelationship = con
   End If
End Function   

function findElementIDByName(p_thePackage, name) 
	
	Dim elem as EA.Element
	Dim thePackage as EA.Package
	
	Set thePackage = p_thePackage
	
	LOGInfo("findElementByName(" & name & ")")
	LOGPackage(thePackage)

	LOGInfo("About to search")
	findElementIDByName=-2
	Set elem = thePackage.Elements.GetByName(name)
	
	if (elem is Nothing ) Then
		LOGInfo("Not found")
		findElementIDByName=-1
	else 
	    LOGElement(elem)
		LOGInfo("Found:" & elem.ElementID)
		findElementIDByName= elem.ElementID
	end if

End Function



End Class