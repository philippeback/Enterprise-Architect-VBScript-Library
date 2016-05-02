!INC Local Scripts.EAConstants-VBScript
!INC Utils.ExcelConnector
!INC Utils.EAConnector
!INC Utils.Datadump

Sub ImportFromExcel(excelFilespec)

	Dim oExcelConnector 
	Dim oEAConnector 
	Dim parentPackage

	Set oExcelConnector = New ExcelConnector
	Set oEAConnector = New EAConnector
	Set parentPackage = oEAConnector.getSelectedPackage()

	Dim oValuesSheet1, oValuesSheet2, oValuesSheet3, oValuesSheet4
  
  'Get everything back from Excel worksheets
  oExcelConnector.OpenExcel(excelFilespec)
  Set oValuesSheet1 = oExcelConnector.GetValues("Elements", 1, 1)
  Set oValuesSheet2 = oExcelConnector.GetValues("DatasetTags", 1, 1)
  Set oValuesSheet3 = oExcelConnector.GetValues("JobTags", 1, 1)
  Set oValuesSheet4 = oExcelConnector.GetValues("Relationships", 1, 1)
  oExcelConnector.CloseExcel()
  
  'Show what we've got
  DataDump "Elements", oValuesSheet1
  DataDump "DatasetTags", oValuesSheet2
  DataDump "JobTags", oValuesSheet3
  DataDump "Relationships", oValuesSheet4
  
  'Titles are in the first row
  LOGInfo("Titles are the first row - Sample sheet1")
  LOGInfo("Heading 1:" & oValuesSheet1.Item(0).Item(0)) 'Works!


   'Let's do some batchin'
   oEAConnector.Speedup
   

  
  LOGInfo("Loading Elements")
  
  Dim oCollection, oItem, oTopRow
  Dim row
  
  row = 0
  Set oTopRow = Nothing
  
  For Each oCollection in oValuesSheet1
     row = row + 1
     If row > 1 Then
	   LOGInfo("Processing row " & row)
	   ImportOneElement oEAConnector, parentPackage, oTopRow, oCollection
	 Else 
	   LOGInfo("Top row")
	   Set oTopRow = oCollection
	 End If
  Next

  
  LOGInfo("Loading Tagged Values for Datasets")
  
  row = 0
  Set oTopRow = Nothing
  
  For Each oCollection in oValuesSheet2
     row = row + 1
     If row > 1 Then
	   LOGInfo("Processing row " & row)
	   ImportOneSetOfTags oEAConnector, parentPackage, oTopRow, oCollection
	 Else 
	   LOGInfo("Top row")
	   Set oTopRow = oCollection
	 End If
  Next
  
  LOGInfo("Loading Tagged Values for Jobss")
  
  row = 0
  Set oTopRow = Nothing
  
  For Each oCollection in oValuesSheet3
     row = row + 1
     If row > 1 Then
	   LOGInfo("Processing row " & row)
	   ImportOneSetOfTags oEAConnector, parentPackage, oTopRow, oCollection
	 Else 
	   LOGInfo("Top row")
	   Set oTopRow = oCollection
	 End If
  Next
  
  
  LOGInfo("Loading Relationships")
  
  row = 0
  Set oTopRow = Nothing
  
  For Each oCollection in oValuesSheet4
     row = row + 1
     If row > 1 Then
	   LOGInfo("Processing row " & row)
	   ImportOneRelationship oEAConnector, parentPackage, oTopRow, oCollection
	 Else 
	   LOGInfo("Top row")
	   Set oTopRow = oCollection
	 End If
  Next
  
  'refresh the contents of the package so that the changes are seen in EA
  oEAConnector.refreshModelView parentPackage
End Sub

'Elements sheet: Stereotype,Element Name, Description
'Skip first row, it contains headings
Sub ImportOneElement(oEAConnector, parentPackage, oTopRow, oCollection)
   'We do not care much about oTopRow in here.
   
   Dim name, stereotype,description
   
   stereotype = oCollection.Item(0)
   name = oCollection.Item(1)
   description = oCollection.Item(2)
   
   oEAConnector.addOrUpdateClass parentPackage, name, stereotype, description
    
End Sub

'Tags sheet: Element name, Tag 1, Tag 2, Tag 3, Tag 4, Tag 5, Tag 6, Tag 7, Tag 8, Tag 9, Tag 10, ...
'Tags are optional
Sub ImportOneSetOfTags(oEAConnector, parentPackage, oTopRow, oCollection)

   Dim strItem, name, tagname, tagvalue
   
   name = oCollection.Item(0)
   
   Dim theElement as EA.Element
   
   Set theElement = oEAConnector.getElementByName(parentPackage, name)
   
   'TODO: if the element is nothing, what to do? Skip?
   
   Dim i
   'We skip the name of the element at index 0
   For i= 1 to oCollection.Count-1
     tagname = oTopRow.Item(i)
	 tagvalue = oCollection.Item(i)
     LOGInfo("Element: " & name & " (" & tagname & ")=" & tagvalue)
	 
	 oEAConnector.addOrUpdateTag theElement, tagName, tagValue
   Next

End Sub

'Relationships sheet
'Source Name, Relationship Type, Relationship Name, Target Name
'Relationsip Name optional
Sub ImportOneRelationship(oEAConnector, parentPackage, oTopRow, oCollection)
   
   'We do not care much about oTopRow in here.
   
   Dim sourceName, relationshipType, relationshipName, targetName
   sourceName = oCollection.Item(0)
   relationshipType = oCollection.Item(1)
   relationshipName = oCollection.Item(2)
   targetName = oCollection.Item(3)
	
	oEAConnector.CreateRelationship parentPackage, sourceName, parentPackage, targetName, relationshipName, relationshipType

End Sub



