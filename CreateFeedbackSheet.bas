Attribute VB_Name = "CreateFeedbackSheet"
Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function



Sub CreateGradingSheet()
'This macro creates a new worksheet called Feedback.
' REQUERIMENTS:
'   - The names of the students selected on the first worksheet (at least one). This macro has to be run after the lab report is submitted,
'   not before.
'   - A function named sheetExists
'   - A worksheet named RUBRIC, with a specific format


    Dim ws As Worksheet
    Dim wr As Worksheet
    Dim xSh As Variant
    Dim sheetName As String
    Dim header As String
    Dim i As Long
    Dim indicators
    Dim possPoints2048
    Dim possPoints1025
    Dim r1 As Range, r2 As Range
    Dim r3 As Range, r4 As Range
    Dim royalBlue As Long
    Dim courseLabel As String
    Dim LastRowRubric As Long
    Dim visShtNum As Long
    Dim rowRef As Integer   ' This a reference, the row where the indicators begin
    
    royalBlue = 6299648
    sheetName = "Feedback"
    
' First of all, there has to be a Sheet named RUBRIC
    If sheetExists("RUBRIC") = False Then
        MsgBox ("The RUBRIC Sheet is not present. This program will stop")
        Exit Sub
    Else
        Set wr = Sheets("RUBRIC")
    End If
  

'Check for the existence of a grade sheet. If exists, delete it
    If sheetExists(sheetName) = False Then
        Set ws = Sheets.Add(After:=Sheets(Worksheets.Count))
        ws.Name = sheetName
    End If
    
' Delete old grading sheet
    If sheetExists("GRADE") Then
        Application.DisplayAlerts = False
        Sheets("GRADE").Delete
        Application.DisplayAlerts = True
    End If
    
' Collect the lab name from the headers that already exist
  header = ActiveWorkbook.Worksheets(2).PageSetup.CenterHeader
    
' Page setup
    With Worksheets(sheetName).PageSetup
        .CenterHeader = header
        .Orientation = xlLandscape
        .LeftMargin = Application.InchesToPoints(0.1)
        .RightMargin = Application.InchesToPoints(0.1)
        .TopMargin = Application.InchesToPoints(0.65)
        .BottomMargin = Application.InchesToPoints(0.65)
        
    End With
    
    Set r1 = Worksheets(sheetName).Range("B1:D4")
    Set r2 = Worksheets(sheetName).Range("B7:D14")
    DrawAllBorders r1
    DrawAllBorders r2
    
 ' Working with the Rubric
 ' Because, in general, the number of indicators is variable, we must make our code flexible to that
 
 LastRowRubric = wr.Cells(wr.Rows.Count, 4).End(xlUp).Row
 
 wr.Range("B1:D" & LastRowRubric).Copy
 ws.Range("B7:D" & (6 + LastRowRubric)).PasteSpecial Paste:=xlPasteAll
 Application.CutCopyMode = False
 
 Set r3 = Worksheets(sheetName).Range("D7:D" & (6 + LastRowRubric))
 Set r4 = Worksheets(sheetName).Range("C7:C" & (6 + LastRowRubric))
    
    With Worksheets(sheetName)
        .Range("B1").Value = "TEAM MEMBERS"
        .Range("B1").Font.Bold = True
        .Range("B1").Interior.Color = royalBlue
        .Range("B1").Font.Color = vbWhite
        
        For i = 5 To 8
        ' TODO: Repeated students' names not allowed
            If Sheets("Intro").Range("C" & i).Value <> "" Then
                .Range("B" & i - 3).Value = Sheets("Intro").Range("C" & i).Value
            End If
        Next i
        
        .Range("C1").Value = "GRADE"
        .Range("D1").Value = "MAX PTS"
        .Range("E7").Value = "COMMENTS"
        
        .Columns("C:D").HorizontalAlignment = xlCenter
        .Range("E7").HorizontalAlignment = xlCenter
        .Columns("B:B").EntireColumn.AutoFit
        .Range("E8:K34").Merge
        .Range("E7:K7").Merge
        .Columns("E:K").VerticalAlignment = xlTop
        .Range("E8:K34").WrapText = True
        .Range("E8:K34").HorizontalAlignment = xlLeft
        
        
        
        For i = 2 To 5
            If .Range("B" & i).Value <> "" Then
                .Range("D" & i).Formula = "=SUM(" & r3.Address(0, 0) & ")"
                .Range("D" & i).Font.Color = vbRed
                .Range("C" & i).Formula = "=SUM(" & r4.Address(0, 0) & ")"
            End If
        Next i
        
        ' Add page number
        visShtNum = 0
        For Each xSht In ActiveWorkbook.Sheets
            If xSht.Visible = True Then visShtNum = visShtNum + 1
        Next
        .PageSetup.CenterFooter = "PAGE " & visShtNum
        
        ' Change the page view to page layout (so the page number becomes visible)
        .Activate
        With ActiveWindow
            .View = xlPageLayoutView
        End With
        
    End With
   TableHeadersHCC Worksheets(sheetName).Range("B1:D1")
   TableHeadersHCC Worksheets(sheetName).Range("E7:K7")
   
      

End Sub


Sub ChangeCentralHeaders()
  ' This sub changes the central headers in all worksheets of a given workbook
  
         Dim WS_Count As Integer
         Dim i As Integer
         Dim oldHeader As String, newHeader As String
         
         ' ############## Change Headers Here  ################
         
         oldHeader = "ELLASTIC COLLISIONS"
         newHeader = "ELASTIC COLLISIONS"
             
         ' #####################################################
     
         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For i = 1 To WS_Count
                If ActiveWorkbook.Worksheets(i).PageSetup.CenterHeader = oldHeader Then
                    ActiveWorkbook.Worksheets(i).PageSetup.CenterHeader = newHeader
                End If
                
            
         Next i

End Sub


Sub DrawAllBorders(R As Range)
 With R
    .Borders(xlEdgeTop).Color = RGB(0, 0, 0)
    .Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
    .Borders(xlEdgeRight).Color = RGB(0, 0, 0)
    .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    .Borders(xlInsideHorizontal).Color = RGB(0, 0, 0)
    .Borders(xlInsideVertical).Color = RGB(0, 0, 0)
    
 End With

End Sub

Sub TableHeadersHCC(R As Range)
    Dim royalBlue As Long
    royalBlue = 6299648
    With R
        .Font.Bold = True
        .Interior.Color = royalBlue
        .Font.Color = vbWhite
    End With

End Sub



Sub Test()
    Dim wr As Worksheet
    Dim LastRubric As Long
    
    Set wr = Sheets("RUBRIC")
    LastRowRubric = wr.Cells(wr.Rows.Count, 4).End(xlUp).Row
    MsgBox (LastRowRubric)
    
End Sub

