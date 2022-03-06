Attribute VB_Name = "CopyNameRowModule"
Option Explicit

Sub CopyNameRow()
' Declare variables
Dim sWB As Workbook
Dim ws As Worksheet
Dim i, j, gN As Integer
Dim sName As String
Dim hRows As String
Dim SelRange As String
Dim gCol As Variant


' Open workbook
Dim iWB As Workbook: Set iWB = ThisWorkbook
iWB.Activate

Application.StatusBar = "Running..."
Application.ScreenUpdating = False

' --------- Parse Data --------- '
Dim df, sMaster, sReadme, tName, colName, folName As String
df = Range("dFile").Value          ' Data file location
sReadme = Range("rSheet").Value    ' README sheet
sMaster = Range("mSheet").Value    ' Master sheet
tName = Range("dTable").Value      ' Name of table that is being copied/pasted into different files
colName = Range("FCol").Value       ' Column  which is being filtered i.e. where the manager names are
folName = Range("sFolder").Value   ' Folder where to save all the results

' ------- Error Checks -------'
' Read me
Dim bReadMe As Boolean
If IsEmpty(sReadme) = True Then
        bReadMe = False
    Else
        bReadMe = True
End If

' ---- Email fields ----'
Dim sEmail As String
Dim bEmail As Boolean
sEmail = Range("eActive").Value
If sEmail = "Yes" Then
    bEmail = True
Else
    bEmail = False
End If

If bEmail = True Then
    Dim sPreview As String
    Dim sSubject As String
    Dim sBody As String
    Dim bSend As Boolean
    Dim cc_col As String
    sPreview = Range("bDisplay").Value
    sSubject = Range("eSub").Value
    sBody = Range("eMessage").Value
    cc_col = Range("eCC").Value
    Dim CCNum As Integer: CCNum = ColumnNumber(cc_col)
    
    
    If sPreview = "Send" Then
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
        Msg = "Email will be sent! THERE IS NO TURNING BACK. Make sure you have previewed the results already. Do you want to continue? Click 'Yes' to confirm sending message or 'No' to preview?"    ' Define message.
        Style = vbYesNo Or vbExclamation Or vbDefaultButton2    ' Define buttons.
        Title = "Confirm Sending Email"    ' Define title.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            bSend = True
        Else    ' User chose No.
            bSend = False
        End If
    Else
        bSend = False
    End If

End If


' Output folder
MkDir (folName) ' Create output folder or check that it exits
If Right(folName, 1) <> "\" Then ' Make sure filepath for the save folder is correct
    folName = folName & "\"
End If


' Convert Input from letters to numbers
Dim colNum As Integer: colNum = ColumnNumber(colName)

' ----------- Master workbook ------- '
' Just get the name of the master workbook
Dim dfname As String
Dim fso As Object: Set fso = CreateObject("scripting.FileSystemObject")
dfname = fso.GetFileName(df)

' Open workbook
Dim mWB As Workbook: Set mWB = OpenWorkbook(df)
mWB.Activate

' Select the master and the readme worksheet
Dim ms, rs As Worksheet
Set ms = Worksheets(sMaster)

If bReadMe = True Then
    Set rs = Worksheets(sReadme)
End If

' First remove any filters
ms.Select
ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum

' Get unique names
Dim d As Object
Dim k As Variant
Set d = uniquenames(tName, colNum)

Dim sfullfile As String
Dim fCriteria As String
For Each k In d.keys
     ' Display status
     iWB.Activate
     i = i + 1
     Application.StatusBar = "Working on: " & i & " of " & d.Count
     mWB.Activate
     
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
    
    If bEmail = True Then
        Dim ccname As String
        ccname = Range(tName).Columns(CCNum).End(xlDown).Value
    End If
    
    ' Select first cell, just to clear it up
    ActiveSheet.Cells(1, 1).Select
    
    ' -------------- Save README file -------------- '
    Set sWB = ActiveWorkbook            ' Temporary workbook
    mWB.Activate                        ' Go back to master file
    
    ' Copy readme file if true
    If bReadMe = True Then
        rs.Select                           ' Select readme tab
        rs.Copy Before:=Workbooks(sWB.Name).Sheets(1)
    End If
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
 
    sfullfile = folName & sName & "_" & dfname
    ActiveWorkbook.SaveAs Filename:=sfullfile, _
          FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    
    If bEmail = True Then
        Call SendEmail(fCriteria, ccname, sfullfile, sSubject, sBody, bSend)
    End If
    
    ' Go back to master sheet and remove any filters
    mWB.Activate        ' Go back to master file
    ms.Select
    ActiveSheet.ListObjects(tName).Range.AutoFilter Field:=colNum
    
Next k
   
mWB.Close SaveChanges:=False

' Clear status bar
Application.ScreenUpdating = True
Application.StatusBar = False
End Sub

' Opening Workbook function
Function OpenWorkbook(df) As Workbook
    ' Hardcoded
    Sheets("Input").Select ' Ensure we are the right sheet
    
    ' Error check
    If IsEmpty(df) = True Then
        ' Message if there is no datafile value at the datafile location
        MsgBox "No data file locaiton detected"
        Exit Function
    End If
    
    ' Open file
    Set OpenWorkbook = Workbooks.Open(df)
    
    ' Returns workbook type varaiable
End Function


Function ColumnNumber(ColumnLetter) As Integer
    'PURPOSE: Convert a given letter into it's corresponding Numeric Reference
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
      
    'Convert To Column Number
       ColumnNumber = Range(ColumnLetter & 1).Column
End Function

' Get unique names in column
Function uniquenames(tName, colNum) As Object
    Dim d As Object, c As Range, tmp As String
    
    Set d = CreateObject("scripting.dictionary")
    ActiveSheet.ListObjects(tName).ListColumns(colNum).DataBodyRange.Select         ' Select range that will be filtered over
    
    For Each c In Selection
        ' Loop over names
        tmp = Trim(c.Value)
        If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
    Next c
  Set uniquenames = d
End Function

' Create folder if it doesn't exist
Function MkDir(path As String)
    
    Dim fso As Object
    Set fso = CreateObject("scripting.FileSystemObject")
    
    If Not fso.FolderExists(path) Then
        ' doesn't exist, so create the folder
        fso.CreateFolder path
    End If
End Function


Sub SendEmail(sTo As String, cName As String, aFile As String, sSub As String, sBody As String, bSend As Boolean)
   ' aFile is the attachment file (full location)
   
   'Setting up the Excel variables.
   Dim olApp As Object
   Dim olMailItm As Object
   Dim iCounter As Integer
   Dim Dest As Variant
   Dim SDest As String
   
   'Create the Outlook application and the empty email.
   Set olApp = CreateObject("Outlook.Application")
   Set olMailItm = olApp.CreateItem(0)
   
   'Using the email, add multiple recipients, using a list of addresses in column A.
   With olMailItm
       .To = sTo
       .CC = cName
       .Subject = sSub
       .Body = sBody
       .Attachments.Add aFile
       If bSend = True Then
        .Send
        Else
        .Display
        End If
   End With
   
   'Clean up the Outlook application.
   Set olMailItm = Nothing
   Set olApp = Nothing
End Sub
