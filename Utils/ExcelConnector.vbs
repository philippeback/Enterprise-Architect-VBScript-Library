!INC EAScriptLib.VBScript-Logging

Class ExcelConnector

  Public xlApp 
  Public xlWB 
  Public xlWS 

  Public xlSortOnValues
  Public xlAscending
  Public xlSortNormal
  Public xlYes

  Private Sub Class_Initialize
    Set xlApp = Nothing
	Set xlWB = Nothing
	Set xlWS = Nothing
	
	xlSortNormal = 0
	xlSortOnValues = 0
	xlAscending = 1
	xlYes=1
	
  End Sub

  Sub OpenExcel(XLS_FILE_LOCATION)

    LOGInfo("Open Excel")
    If (xlApp Is Nothing) Then
      LOGInfo("Opening")
      Set xlApp = CreateObject("Excel.Application")
      
      'xlApp.Visible = True
      'xlApp.ScreenUpdating = False
      Set xlWB = xlApp.Workbooks.Open(XLS_FILE_LOCATION,True)
      Set xlWS = xlWB.Worksheets(1)
      LOGInfo("Excel Opened:" & xlWB.Name)
    Else
        LOGInfo("Excel reused")
    End If

  End Sub

  Sub CloseExcel
    LOGInfo("Close Excel")
    If (xlApp Is Nothing) Then
       LOGInfo("Nothing to do")
    Else
       LOGInfo("Closing Excel")
       xlWB.Close False ' close the workbook without saving
       xlApp.Quit ' close the Excel application
       Set xlWB = Nothing
       Set xlApp = Nothing
    End If
  End Sub
  
  Sub SetWorksheet(worksheetNumber)
    LOGInfo("SetWorksheet(" & worksheetNumber & ")")
    Set xlWS = xlWB.Worksheets(worksheetNumber)
  End Sub

'-------------------------------------------------------------
' Author:   Philippe Back, from Geert Bellekens code
' Date:     20/10/2011 - 06/02/2008
' Description: Returns the values of a given sheet as a two-dimensional collection of strings
'-------------------------------------------------------------
 Function GetValues(worksheetName, startrow, startcol)

    Dim row 
    Dim col 
    Dim colValues 
	
	'go to proper sheet
	Set xlWS = xlWB.Sheets.Item(worksheetName)
	
    'initialize return
    Set GetValues = CreateObject("System.Collections.ArrayList") 'Magic!
	
	LOGInfo("GetValues")
    'initialize column and row counter
    row = startrow
    col = startcol
    Do Until xlWS.Cells(row, 1).Value = "" 'loop rows
	   LOGInfo("Reading row:" & row)
       'reset column values
       Set colValues = CreateObject("System.Collections.ArrayList")
       'reset column counter
       col = startcol
       'loop columns until two consecutive empty cells are found
       Do Until xlWS.Cells(row, col).Value = "" And xlWS.Cells(row, col + 1).Value = ""
	     LOGInfo("Reading cell(" & row & " , " & col & ")")
         colValues.Add CStr(xlWS.Cells(row, col).Value)
         col = col + 1 'up column counter
       Loop
       GetValues.Add colValues
       row = row + 1 'up row counter
    Loop
	
End Function

End Class
