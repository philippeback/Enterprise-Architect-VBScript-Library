option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
Const ForReading = 1

Dim DebugFlag
DebugFlag=False
'DebugFlag=True

Dim xlApp 
Dim xlWB 
Dim xlWS
Dim aDocTypes()
Dim iDocTypes 

Set xlApp = Nothing

Const XLS_FILE_LOCATION  = "Test.xls"
Const CACHE_FILENAME = "C:\TEMP\XXXXCache.xml"


Dim xlSortOnValues
xlSortOnValues = 0
Dim xlAscending
xlAscending = 1
Dim xlSortNormal
xlSortNormal = 0
Dim xlYes
xlYes=1

Dim oCache
Set oCache = CreateObject("Scripting.Dictionary")

Dim regEx
Set regEx = New RegExp         ' Create regular expression.
regEx.IgnoreCase = True

Function SanitizeFilename(byVal strFilename, byVal strReplChar)	
      
      Dim oRegExp 
      
      Set oRegExp = New RegExp ' Create new RegExp object
      
      ' Define regex pattern and set Global replacement property
      oRegExp.Pattern = "[\x00-\x1f\x22\/\*\?<>\|]"
      oRegExp.Global = True
 
      ' Check strReplChar parameter for invalid length or character,
      ' default to underscore
      If Len(strReplChar) = 1 Then 
         strReplChar = oRegExp.Replace(strReplChar, "_")
      Else 
         strReplChar = "_"
      End If
 
      ' Return clean filename
      SanitizeFilename = oRegExp.Replace(strFilename, strReplChar)
 
      Set oRegExp = Nothing
End Function



Sub OpenExcel

    LogInfo("Open Excel")
    If (xlApp Is Nothing) Then
      LogInfo("Opening")
      Set xlApp = CreateObject("Excel.Application")
      
      'xlApp.Visible = True
      'xlApp.ScreenUpdating = False
      Set xlWB = xlApp.Workbooks.Open(XLS_FILE_LOCATION,True)
      Set xlWS = xlWB.Worksheets(1)
      'MsgBox "Excel Opened:" & xlWB.Name
    Else
        LogInfo("Excel reused")
    End If

End Sub

Sub CloseExcel
    LogInfo("Close Excel")
    If (xlApp Is Nothing) Then
       LogInfo("Nothing to do")
    Else
       LogInfo("Closing Excel")
       xlWB.Close False ' close the workbook without saving
       xlApp.Quit ' close the Excel application
       Set xlWB = Nothing
       Set xlApp = Nothing
    End If
End Sub


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

'Clear the cache (memory only)
Function ClearCache
  oCache.RemoveAll()
End Function

'Clear the cache, disk and memory included.
Function ClearCacheAll()
  Stop
  'Not implemented
End Function

Function InCache(key)
  InCache = oCache.Exists(key)
End Function

'Cache content into the cache for the given key
Function StoreInCache(key,content)
  oCache.Add key, content
  StoreInCache = True
End Function

'Retrive content from the cache based on the key
Function RetrieveFromCache(key)
  LogInfo("RetreiveFromCache("&key&")")
  RetrieveFromCache = oCache.Item(key)
End Function

'Persists the cache on disk
Function PersistCache()

    Dim fso
    Set fso=CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.DeleteFile(CACHE_FILENAME)
    On Error Goto 0 'TODO: handle this nicer
    Dim cacheFile 
    Set cacheFile = fso.OpenTextFile(CACHE_FILENAME, 2, True) 'for writing, as Unicode
    
    'Dim oXmlParser
    'Set oXmlParser = CreateObject("MSXML.DOMDocument")
    
    cacheFile.WriteLine("<?xml version=""1.0""?>")
    cacheFile.WriteLine("<cache docrows=""" & tablerowId  & """>")
    
    Dim key, value
    For Each key In oCache.Keys
      cacheFile.WriteLine(vbTab & "<entry key=""" & key & """>")
      cacheFile.WriteLine(vbTab & "<![CDATA[")
      cacheFile.WriteLine(oCache.Item(key))
      cacheFile.WriteLine(vbTab & "]]>")
      cacheFile.WriteLine(vbTab & "</entry>")
    Next 
    cacheFile.WriteLine("</cache>")
    cacheFile.Close
    Set cacheFile = Nothing
    Set fso = Nothing
End Function

'Load the cache from disk
Function LoadCache()
    
    Dim fso
    Set fso=CreateObject("Scripting.FileSystemObject")
    
    
    If FileExists(CACHE_FILENAME) Then
       LogInfo "Cache file exists"
    Else
       LogInfo "Cache file not found"
    End If
    
    On Error Goto 0
    Dim oXmlDoc    
    Set oXmlDoc = CreateObject("Microsoft.XMLDOM") 

    oXmlDoc.async =False 
    oXmlDoc.validateOnParse=False

    Dim cacheFile 
    Set cacheFile = fso.OpenTextFile(CACHE_FILENAME, 1, True) 'for writing, as Unicode
    Dim strContents
    strContents = cacheFile.ReadAll
    cacheFile.Close
    
    Dim fLoad
    
    fLoad = oXmlDoc.loadXML(strContents) 
    
    'Loop through "entry" nodes and get the key attribute and the node value is the content
    
    oCache.RemoveAll
    
     
    Dim oTop
    Set oTop = oXmlDoc.selectNodes("//cache")
    
    tablerowId = oTop(0).getAttribute("docrows")
    LogInfo "docrows: " & tablerowId
    
    Set oTop = Nothing
    
    Dim oNodes
    
    Set oNodes = oXmlDoc.getElementsByTagName("entry")
    'Set oNodes = oXmlDoc.selectNodes("//entry")

    Dim i, key, value
    
    For i=0 To oNodes.Length-1
      key = oNodes(i).getAttribute("key")
      value = oNodes(i).Text
      ' LogInfo "From XML: " & key & "-->" & oNodes(i).xml & "..."
      oCache.Add key, value
    Next
    
    Set oNodes =  Nothing
    
    Set oXmlDoc = Nothing
    
    'cacheFile.Close
    Set cacheFile = Nothing
    
    Set fso = Nothing
    
    LogInfo "Cache file loaded"
    
    CreateOverview 'This is added here for ease of use: loads doc directly after cache for visual feedback.
    

End Function


Function CreateBox(title,trigram)

  LogWarn("CreateBox(" & trigram & ")")

  AddShortcut trigram, title

  If (InCache(trigram)=True) Then
       CreateBox = RetrieveFromCache(trigram)
       LogWarn("Exit CreateBox(" & trigram & ")  - caching used.")
       Exit Function
  Else
  
      Call OpenExcel

	  Dim strHtml
	
	  strHtml = "<a name=""" & trigram & """/><table class=""box""><tr><th colspan=""2"">" & title & " (" & trigram & ")&nbsp;<a href=""#top"">^Top</a></th></tr>"
	
	  Dim xlRange, xlRange2, xlRangeFiles
	
	  Set xlRange = xlWS.Range("A4:R2000")
	  Set xlRange2 = xlWS.Range("G4")
	  xlRange.Sort xlRange2, xlAscending
	  ' Sort not working yet...
	  
	
	  Dim r
	  r = 2
	
	  Dim theValue
	
	  While xlRange.Cells(r, 1) <> "" 'was .Value
	    theValue = xlRange.Cells(r,1)
	    if (trigram = theValue) Then
	      NextTableRow
	      strHtml = strHtml & "<tr id=""" & CreateTableRowId() & """><td id=""" & CreateTableCellId() & """>" & xlRange.Cells(r,7) & "</td><td width=""30%"">" & CreateActions(xlRange.Cells(r,18), xlRange.Cells(r,10)) & "</td></tr>"
	    end if 
	    r = r + 1
	  Wend
	
	  strHtml = strHtml & "</table>"
	
	  'MsgBox "TABLE: " & strHtml
	
	  StoreInCache trigram, strHtml
	  
	  CreateBox = strHtml
	  
	  LogWarn("Exit CreateBox(" & trigram & ")")
  End If	  


End Function

