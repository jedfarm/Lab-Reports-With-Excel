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
    Dim myRoster() As String
    Dim i As Long
    Dim range1 As Range, rng As Range
    Dim visN As Long
    Dim footer As String
    Dim labName As String
    
    'Check for the existence of a previous roster. If it exists delete it
    If sheetExists("Roster") Then
        Sheets("Roster").Visible = xlSheetVisible
        Application.DisplayAlerts = False
        Sheets("Roster").Delete
        Application.DisplayAlerts = True
    End If
    
    'Import the roster from a .csv file
    Set wsht = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    Set wbkS = Workbooks.Open(Filename:=ActiveWorkbook.Path & "\Roster.csv")
    Set wshS = wbkS.Worksheets(1)
    wshS.Range("A1:A2").EntireRow.Delete
    wshS.UsedRange.Copy Destination:=wsht.Range("A1")
    wbkS.Close SaveChanges:=False
    wsht.Name = "Roster"
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
                If i > 1 Then
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

