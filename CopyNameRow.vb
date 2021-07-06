Option Explicit

Sub CopyNameRow()
' Declare variables
Dim ms As Worksheet
Dim ws As Worksheet
Dim tName As String
Dim colNum As Integer
Dim fCriteria As String
Dim sFolder As String

' --------- User Input --------- '
tName = "Table1"                       ' Name of table that is being copied/pasted into different files (with "" around it)
colNum = 5                             ' Column NUMBER which is being filtered i.e. where the manager names are
sFolder = "C:\Users\siqbal\Desktop\"   ' Folder where to save all the results
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
    ' Criteria to filter against
    fCriteria = k
    
    ' Filter results based on fCriteria
    ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum, Criteria1:=fCriteria
    
    ' Select everything on the filtered table
    Range(tName & "[#All]").Select
    Selection.Copy
    
    ' Add new worksheet
    Sheets.Add After:=ActiveSheet
    Set ws = ActiveSheet
    
    ' Paste Format only (colors and column widths)
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    
    ' Pase formula and number format only
    Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
    
    ' Save Worksheet
    ws.Select
    ws.Move
    ActiveWorkbook.SaveAs Filename:=sFolder & fCriteria & ".xlsx", _
          FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    
    ' Go back to master sheet and remove any filters
    ms.Select
    ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum
Next k

End Sub
