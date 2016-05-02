option explicit

!INC Local Scripts.EAConstants-VBScript
!INC EAScriptLib.VBScript-Logging
!INC Utils.ExcelImport
'
' Script Name: Unit Tests for importing into EA from Excel
' Author: Philippe Back
' Purpose: Test the full chain
' Date: 27/10/2011
'

Dim excelFileSpec

Sub Setup
 
  excelFileSpec = "H:\TestImportExcel.xls"
  
End Sub

Sub Teardown
  '
End Sub

Sub ExcelImportTest
	LOGInfo("ExcelImportTest")
	ImportFromExcel excelFilespec
End Sub

Sub TestSuite
  Setup
  ExcelImportTest
  Teardown
End Sub 

TestSuite
