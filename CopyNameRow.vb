Sub CopyNameRow()

' Declare variables
Dim tName As String
Dim cNum As Integer
Dim fCriteria As String
Dim sFolder As String

' --------- User Input --------- '
tName = "Table1"                       ' Name of table that is being copied/pasted into different files (with "" around it)
colNum = 5                             ' Column NUMBER which is being filtered
fCriteria = "bronco"                   ' Filter name i.e. manager name (with "" around it)
sFolder = "C:\Users\siqbal\Desktop\"   ' Folder where to save all the results

' Select the first worksheet, assumed to be the master sheet
Set ms = Worksheets(1)
ms.Select

' First remove any filters
ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum

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
End Sub
