option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Utils.EAConnector
!INC Utils.Datadump
'
' Script Name: Unit Tests for EA Connector
' Author: Philippe Back
' Purpose: Check if the EA class works as advertised
' Date: 25/10/2011
'

Dim oEAConnector
    
Dim parentPackage as EA.Package

Const TESTCLASS_NAME = "MYTESTCLASS"
Const TESTATTRIBUTE_NAME = "MYTESTATTRIBUTE"
Const TESTTAG_NAME = "MYTESTTAG"
Const TESTTAG_VALUE = "MY VALUE FOR TAG"

Const TESTREL_NAME = "MYRELNAME"
Const TESTREL_TYPE = "Association"
Const TESTTRGCLASS_NAME = "MYTESTTRGCLASS"


Sub Setup
  Set oEAConnector = New EAConnector
  
  Set parentPackage = oEAConnector.getSelectedPackage()
  LOGInfo("Parent Package=" & parentPackage.Name)
End Sub

Sub Teardown
  oEAConnector.refreshModelView parentPackage
  Set parentPackage=Nothing
  Set oEAConnector=Nothing
End Sub

Sub OpenConnectorTest
  LOGInfo("OpenConnectorTest")
  oEAConnector.speedUp
  Dim oAllClasses
  
  Set oAllClasses = oEAConnector.getClasses(oEAConnector.getSelectedPackage())
  CollectionOfElementsDump "All classes of selected package",oAllClasses
End Sub

Sub AddOrUpdateElementToPackageTest

    LOGInfo("AddOrUpdateElementToPackageTest")
	
    Dim name, stereotype,description
	
	name = TESTCLASS_NAME
	
	stereotype = "mystereo"
	description = "A sample description of mine"
	oEAConnector.addOrUpdateClass parentPackage, name, stereotype, description
	description = "A modified description of mine"
	oEAConnector.addOrUpdateClass parentPackage, name, stereotype, description
	
	
	
End Sub

Sub AddAOrUpdatettributeToElementTest
	LOGInfo("AddOrUpdateAttributeToElementTest")
	
	
	Dim parentClass as EA.Element
	
	Set parentClass = oEAConnector.getElementByName(parentPackage, TESTCLASS_NAME)
	
	If parentClass Is Nothing Then
	  LOGError("Parent class is nothing, can't proceed")
	Else
	  LOGInfo("Parent class:" & parentClass.Name)
	End If
	
	Dim name, stereotype,description,attrType
	
	name = TESTATTRIBUTE_NAME
	stereotype = "tst"
	description = "A sample attribute description of mine"
	attrType = "SuperString"
	
	oEAConnector.addOrUpdateAttribute parentClass, name, stereotype, description, attrType
	
	attrType = "MegaString"
    oEAConnector.addOrUpdateAttribute parentClass, name, stereotype, description, attrType

End Sub

Sub AddOrUpdateAttributeTaggedValueTest
	LOGInfo("AddOrUpdateAttributeTaggedValueTest")
	
	
	Dim parentClass as EA.Element
	
	Set parentClass = oEAConnector.getElementByName(parentPackage, TESTCLASS_NAME)
	
	If parentClass Is Nothing Then
	  LOGError("Parent class is nothing, can't proceed")
	Else
	  LOGInfo("Parent class:" & parentClass.Name)
	End If
	
	Dim anAttribute as EA.Attribute
	
	Set anAttribute = oEAConnector.getAttributeByName(parentClass, TESTATTRIBUTE_NAME)
	
	Dim tagValue
	tagValue = TESTTAG_VALUE
	
    oEAConnector.addOrUpdateAttributeTag anAttribute, TESTTAG_NAME, tagValue
	
End Sub

Sub AddOrUpdateTaggedValueTest
  LOGInfo("AddOrUpdateTaggedValueTest")
  
  'Make sure we have the classes we need to connect together
  Dim theElement as EA.Element
  
  Set theElement = oEAConnector.getElementByName(parentPackage, TESTCLASS_NAME)
  
  If theElement Is Nothing Then
   	   Set theElement = oEAConnector.addOrUpdateClass(parentPackage, TESTCLASS_NAME, "", "")
  End If
	
  oEAConnector.addOrUpdateTag theElement, TESTTAG_NAME, TESTTAG_VALUE
  
End Sub

Sub CreateConnectorTest
    LOGInfo("CreateConnectorTest")
	
	'Make sure we have the classes we need to connect together
	If oEAConnector.getElementByName(parentPackage, TESTCLASS_NAME) Is Nothing Then
   	   oEAConnector.addOrUpdateClass parentPackage, TESTCLASS_NAME, "", ""
	End If
	
	If oEAConnector.getElementByName(parentPackage, TESTTRGCLASS_NAME) Is Nothing Then
	   oEAConnector.addOrUpdateClass parentPackage, TESTTRGCLASS_NAME, "", ""
	End If
	LOGInfo("Creating the relationship")
	oEAConnector.CreateRelationship parentPackage, TESTCLASS_NAME, parentPackage, TESTTRGCLASS_NAME, TESTREL_NAME, TESTREL_TYPE
	
End Sub

Sub TestSuite
  ShowScriptOutputWindow
  Setup
  OpenConnectorTest
  AddOrUpdateElementToPackageTest
  AddAOrUpdatettributeToElementTest
  AddOrUpdateAttributeTaggedValueTest
  AddOrUpdateTaggedValueTest 
  CreateConnectorTest
  Teardown
End Sub 

TestSuite


