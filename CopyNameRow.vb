Option Explicit

Sub CopyNameRow()

' Declare variables
Dim ms As Worksheet
Dim ws As Worksheet
Dim tName As String
Dim sName As String
Dim colNum As Integer
Dim fCriteria As String
Dim sFolder As String
Dim hRows As String

' --------- User Input --------- '
tName = "Table1"                       ' Name of table that is being copied/pasted into different files (with "" around it)
colNum = 5                             ' Column NUMBER which is being filtered i.e. where the manager names are
sFolder = "C:\Users\siqbal\Desktop\"   ' Folder where to save all the results
hRows = "1:5"                          ' Header rows, beginning:end (with " around it)
' ---------------------------------

' -------------------------------------------- '
' ------------- CODE STARTS HERE ------------- '
' -------------------------------------------- '
' Select the first worksheet, assumed to be the master sheet
Set ms = Worksheets(1)
ms.Select

' First remove any filters
ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum

' Get unique names in column
Dim d As Object, c As Range, k, tmp As String
Set d = CreateObject("scripting.dictionary")
ActiveSheet.ListObjects(tName).ListColumns(colNum).Range.Select ' Select range that will be filtered over
For Each c In Selection ' Loop over names
    tmp = Trim(c.Value)
    If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
Next c

For Each k In d.keys

    ' -------------- Header Copy/Paste -------------- '
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
            
    ' Go back to master worksheet
    ms.Select
    
    ' -------------- Filtered Table Copy/Paste -------------- '
    ' Criteria to filter against
    fCriteria = k
        
    ' Filter results based on fCriteria
    ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum, Criteria1:=fCriteria
    
    ' Select everything on the filtered table
    Range(tName & "[#All]").Select
    Selection.Copy
    
    ' Go to the newly added worksheet with headers already and on the last empty row
    ws.Select
    Range("A" & Rows.Count).End(xlUp).Offset(1).Select
    
    ' Paste Format only (colors and column widths)
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    
    ' Paste formula and number format only
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    
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

End Sub

