Option Explicit

Sub CopyNameRow()
' Declare variables
Dim sWB As Workbook
Dim ms, rs, ws As Worksheet
Dim tName As String
Dim sName As String
Dim colNum, i, j, gN As Integer
Dim fCriteria As String
Dim sFolder As String
Dim hRows As String
Dim SelRange As String
Dim gCol As Variant
Dim sMaster, sReadme As Integer

' --------- User Input --------- '
sMaster = 2                                              ' Master sheet
sReadme = 3                                              ' README sheet
tName = "Master"                                         ' Name of table that is being copied/pasted into different files (with "" around it)
colNum = 5                                               ' Column NUMBER which is being filtered i.e. where the manager names are
sFolder = "C:\Users\siqbal\Desktop\Cboo Maco\Results\"   ' Folder where to save all the results

' -------------------------------------------- '
' ------------- CODE STARTS HERE ------------- '
' -------------------------------------------- '
Application.StatusBar = "Running..."
Application.ScreenUpdating = False

' Get master workbook
Dim mWB As Workbook
Set mWB = ThisWorkbook

' Select the master and the readme worksheet
Set ms = Worksheets(sMaster)
Set rs = Worksheets(sReadme)

' Get name of these sheets
Dim mName, rName As String
mName = ms.Name
rName = rs.Name

' Start ot filter
ms.Select

' First remove any filters
ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum

' Get unique names in column
Dim d As Object, c As Range, k, tmp As String
Set d = CreateObject("scripting.dictionary")
ActiveSheet.ListObjects(tName).ListColumns(colNum).DataBodyRange.Select         ' Select range that will be filtered over

For Each c In Selection
    ' Loop over names
    tmp = Trim(c.Value)
    If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
Next c

' Copy paste all unique names
i = 0
For Each k In d.keys
    ' Display status
    i = i + 1
    Application.StatusBar = "Working on: " & i & " of " & d.Count
  
    ' Criteria to filter against
    fCriteria = k
        
    ' -------------- Save table to a new file -------------- '
    ms.Select                   ' Ensure we are in the master sheet
    ms.Copy                     ' Move copy of worksheet to new file
                            
    ' Filter results based on fCriteria (exclude fCriteria)
    ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum, Criteria1:="<>" & fCriteria
    
    ' Select everything on the filtered table
    Range(tName).Select
    ActiveSheet.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8  ' Expand all group (have to do it)
    Selection.EntireRow.Delete
    ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1  ' Collapse them all again
    
    ' Deactivate filter
    ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum
    
    ' Select first cell, just to clear it up
    ActiveSheet.Cells(1, 1).Select
    
    ' -------------- Save README file -------------- '
    Set sWB = ActiveWorkbook            ' Temporary workbook
    mWB.Activate        ' Go back to master file
    rs.Select           ' Select readme tab
    rs.Copy Before:=Workbooks(sWB.Name).Sheets(1)
    sWB.Activate
    
    ' -------------- Save Worksheet -------------- '
    sName = Replace(fCriteria, "@oracle.com", "") ' Repalce the email address after @ to nothing
    sName = Replace(sName, "@ORACLE.COM", "")     ' Repalce the email address in capitol
    sName = Replace(sName, "?", "")               ' Remove unallowed character
    sName = Replace(sName, ".", "_")              ' Remove unallowed character
    sName = Replace(sName, ":", "")               ' Remove unallowed character
    sName = Replace(sName, "*", "")               ' Remove unallowed character
    sName = Replace(sName, "\", "")               ' Remove unallowed character
    sName = Replace(sName, "/", "")               ' Remove unallowed character
    ActiveWorkbook.SaveAs Filename:=sFolder & sName & ".xlsx", _
          FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    
    ' Go back to master sheet and remove any filters
    mWB.Activate        ' Go back to master file
    ms.Select
    ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum
Next k

' Clear status bar
Application.ScreenUpdating = True
Application.StatusBar = False
End Sub
