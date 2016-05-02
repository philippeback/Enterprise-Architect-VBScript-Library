option explicit

!INC Local Scripts.EAConstants-VBScript
!INC EAScriptLib.VBScript-Logging

'
' Script Name: DotNetArrayListTest
' Author: Philippe Back
' Purpose: Check if .NET collection work from VBScript embedded in EA.
' Date: 20/10/2011
'
' Test looks okay, great!
''.Count useful
  'Methods: http://msdn.microsoft.com/en-us/library/system.collections.arraylist_methods%28v=VS.71%29.aspx
'returnValue = ArrayListObject.Item(index)
'ArrayListObject.Item(index) = returnValue
  

Dim DataList, strItem

Set DataList = CreateObject("System.Collections.ArrayList") 'Magic! - Die Dictionary, die!

DataList.Add "B"
DataList.Add "C"
DataList.Add "E"
DataList.Add "D"
DataList.Add "A"

DataList.Sort()

LOGInfo("Eaching out:")
For Each strItem in DataList
     LOGInfo(strItem)
Next

LOGInfo("Looping:")
Dim i
For i=0 to DataList.Count-1
     LOGInfo(DataList.Item(i))
Next

