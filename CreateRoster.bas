Attribute VB_Name = "CreateRoster"

Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

Sub CreateRoster()

'This macro needs a .csv file (provided by CANVAS) with the roster, it has to be located in the same folder as the target Excel file.
' The roster file must be named Roster.csv
' This macro will create a drop down list in cells C5, C6 and C7 of the first tab ("Intro").
  
    
    Dim wbkS As Workbook
    Dim wshS As Worksheet
    Dim wsht As Worksheet
    Dim wst As Worksheet
    Dim lastRow As Long
    Dim sectionCol As Long
    Dim myRoster() As String
    Dim i As Long
    Dim range1 As Range, rng As Range
    Dim visN As Long
    Dim footer As String
    Dim labName As String
    Dim sectionID As String
    
    'Check for the existence of a previous roster. If it exists delete it
    If sheetExists("Roster") Then
        Sheets("Roster").Visible = xlSheetVisible
        Application.DisplayAlerts = False
        Sheets("Roster").Delete
        Application.DisplayAlerts = True
    End If
    
    'Import the roster from a .csv file
    Set wsht = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    Set wbkS = Workbooks.Open(filename:=ActiveWorkbook.Path & "\Roster.csv")
    Set wshS = wbkS.Worksheets(1)
    
    sectionCol = -1 ' In case there is no column named "Section" in the Roster file.
    On Error Resume Next
    sectionCol = wshS.Cells(1, 1).EntireRow.Find(What:="Section", LookIn:=xlValues, LookAt:=xlPart, _
    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
    
    wshS.Range("A1:A2").EntireRow.Delete
    wshS.UsedRange.Copy Destination:=wsht.Range("A1")
    wbkS.Close SaveChanges:=False
    wsht.Name = "Roster"
    If sectionCol > -1 Then
        If IsEmpty(wsht.Cells(1, sectionCol)) = False Then
            sectionID = wsht.Cells(1, sectionCol).Value
        Else
            sectionID = "XX/XX-PHY-XXXXL-XXXXX"
        End If
    Else
        sectionID = "XX/XX-PHY-XXXXL-XXXXX"
    End If
    wsht.Range("B:XFD").ClearContents
    
    
    'Find the last non-blank cell in column A(1)
    lastRow = wsht.Cells(Rows.Count, 1).End(xlUp).Row
    
    Set wst = Sheets("Intro")
    
    Set range1 = wsht.Range("A1:A" & lastRow)
    
    For i = 5 To 7
        Set rng = wst.Range("C" & i)
        With rng.Validation
            .Delete 'delete previous validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:="='" & wsht.Name & "'!" & range1.Address
        End With
    Next i
    Sheets("Roster").Visible = xlSheetVeryHidden
    
    ' Write Page Numbers and Headers
        visN = 0
        labName = ActiveWorkbook.Worksheets(1).Range("A2").Value
        For i = 1 To ActiveWorkbook.Worksheets.Count
           If ActiveWorkbook.Worksheets(i).Visible = True Then
            visN = visN + 1
            footer = "PAGE " & visN
            With ActiveWorkbook.Worksheets(i)
                If i = 1 Then
                    .PageSetup.CenterHeader = sectionID
                Else
                    .PageSetup.CenterHeader = labName
                End If
                .PageSetup.CenterFooter = footer
                .Activate
                With ActiveWindow
                    .View = xlPageLayoutView
                End With
            End With
           End If
        Next i
            
      
End Sub


Sub CreateRosterForAllExcelFilesInFolder()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(filename:=myPath & myFile)
    
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    'Create or update the roster
      CreateRoster
    
    'Save and Close Workbook
      wb.Close SaveChanges:=True
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
