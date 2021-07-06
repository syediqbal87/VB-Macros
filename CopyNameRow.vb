Option Explicit

Sub CopyNameRow()

' Declare variables
Dim ms As Worksheet
Dim ws As Worksheet
Dim tName As String
Dim sName As String
Dim colNum, i As Integer
Dim fCriteria As String
Dim sFolder As String
Dim hRows As String
Dim SelRange As String

' --------- User Input --------- '
tName = "Table1"                       ' Name of table that is being copied/pasted into different files (with "" around it)
colNum = 5                             ' Column NUMBER which is being filtered i.e. where the manager names are
sFolder = "C:\Users\siqbal\Desktop\"   ' Folder where to save all the results
hRows = "1:5"                          ' Header rows, beginning:end (with " around it)
' ---------------------------------

' -------------------------------------------- '
' ------------- CODE STARTS HERE ------------- '
' -------------------------------------------- '
Application.StatusBar = "Running..."

' Select the first worksheet, assumed to be the master sheet
Set ms = Worksheets(1)
ms.Select

' First remove any filters
ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum

' Get cell where table starts to copy/paste it in the same place later
Dim rng As Range
Dim rngAddr As String
ActiveSheet.ListObjects(tName).Range(1).Select ' Select first cell of table
Set rng = ActiveCell
rngAddr = rng.Address                           'Returns address of start of table

' Get unique names in column
Dim d As Object, c As Range, k, tmp As String
Set d = CreateObject("scripting.dictionary")
ActiveSheet.ListObjects(tName).ListColumns(colNum).Range.Select ' Select range that will be filtered over

For Each c In Selection ' Loop over names
    tmp = Trim(c.Value)
    If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
Next c

' Copy paste all unique names
i = 0
For Each k In d.keys
    ' Display status
    i = i + 1
    Application.StatusBar = "Working on: " & i & " of " & d.Count
 Debug.Print i
  
    ' -------------- Copy/Paste Headers -------------- '
    ' Copy headers first
    Rows(hRows).Select
    Selection.Copy

    ' Add new worksheet
    Sheets.Add After:=ActiveSheet
    Set ws = ActiveSheet
    
    ' Paste Format only (colors and column widths)
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        
    ' Paste formula and number format only
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False

    ' -------------- Filtered Table -------------- '
    ' Ensure we are in the master sheet
    ms.Select
        
    ' Criteria to filter against
    fCriteria = k
        
    ' Filter results based on fCriteria
    ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum, Criteria1:=fCriteria
    
    ' Select everything on the filtered table
    Range(tName & "[#All]").Select
    Selection.Copy
    
    ' Go to the newly added worksheet with headers already and paste table (same location as original table)
    ws.Select
    ActiveSheet.Range(rngAddr).Select
        
    ' Paste Format only (colors and column widths)
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    
    ' Paste formula and number format only
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    
    ' Activate filter
    Selection.AutoFilter
    
    ' Save Worksheet
    sName = Replace(fCriteria, "@oracle.com", "") ' Repalce the email address after @ to nothing
    sName = Replace(sName, "?", "")               ' Remove unallowed character
    sName = Replace(sName, ".", "_")              ' Remove unallowed character
    sName = Replace(sName, ":", "")               ' Remove unallowed character
    sName = Replace(sName, "*", "")               ' Remove unallowed character
    sName = Replace(sName, "\", "")               ' Remove unallowed character
    sName = Replace(sName, "/", "")               ' Remove unallowed character
    
    ws.Select
    ws.Move
    ActiveWorkbook.SaveAs Filename:=sFolder & sName & ".xlsx", _
          FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    

    ' Go back to master sheet and remove any filters
    ms.Select
    ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum
    
    
Next k

Application.StatusBar = False

End Sub


