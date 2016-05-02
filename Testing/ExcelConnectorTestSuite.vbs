option explicit

!INC Local Scripts.EAConstants-VBScript
!INC Utils.ExcelConnector
!INC Utils.Datadump
'
' Script Name: Unit Tests for Excel Connector
' Author: Philippe Back
' Purpose: Check if the ExcelConnector class works as advertised
' Date: 20/10/2011
'

Dim oExcelConnector

Sub Setup
  Set oExcelConnector = New ExcelConnector
End Sub

Sub Teardown
  Set oExcelConnector=Nothing
End Sub

Sub OpenCloseExcelTest
  oExcelConnector.OpenExcel("H:\TestExcel.xls")
  oExcelConnector.CloseExcel()
End Sub

Sub GetValuesTest

  Dim oValuesSheet1, oValuesSheet2, oValuesSheet3
  
  oExcelConnector.OpenExcel("H:\TestExcel.xls")
  Set oValuesSheet1 = oExcelConnector.GetValues("Elements", 1, 1)
  Set oValuesSheet2 = oExcelConnector.GetValues("Tags", 1, 1)
  Set oValuesSheet3 = oExcelConnector.GetValues("Relationships", 1, 1)
  oExcelConnector.CloseExcel()
  DataDump "Elements", oValuesSheet1
  DataDump "Tags", oValuesSheet2
  DataDump "Relationships", oValuesSheet3
  
  LOGInfo("Titles are the first row - Sample sheet1")
  LOGInfo("Heading 1:" & oValuesSheet1.Item(0).Item(0)) 'Works!
  
End Sub

Sub TestSuite
  Setup
  'OpenCloseExcelTest
  GetValuesTest
  Teardown
End Sub 

TestSuite