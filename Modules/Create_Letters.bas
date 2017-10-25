Attribute VB_Name = "Create_Letters"
Option Explicit

Private Sub CopyColumnWidths(TargetRange As Range, SourceRange As Range)
Dim c As Long
    With SourceRange
        For c = 1 To .Columns.Count
            TargetRange.Columns(c).ColumnWidth = .Columns(c).ColumnWidth
        Next c
    End With
End Sub

Public Sub Breakdown(RowNumber As Integer, theSheetRef As String, theStudentNumber As Integer)
Dim d As Range
Dim Cell As Range
'Dim lastColumn As String
'Dim lastColumNumber As Integer
Dim i As Integer
Dim j As Integer
Dim endflag As Integer


i = 0
j = 0
Application.DisplayAlerts = False       'Prevent merge warning from popping up
'Copy new student performance information from the source sheet


Worksheets("Letters_Home").Range("A" & RowNumber & ":AE" & RowNumber + 2).Value = Worksheets(theSheetRef).Range("A2:AE4").Value
Worksheets("Letters_Home").Range("A" & RowNumber + 3).Value = "Your Child's Score:"
Worksheets("Letters_Home").Range("B" & RowNumber + 3 & ":AF" & RowNumber + 3).Value = Worksheets(theSheetRef).Range("B" & theStudentNumber & ":AF" & theStudentNumber).Value
Worksheets("Letters_Home").Range("A" & RowNumber & ":AE" & RowNumber + 3).HorizontalAlignment = xlVAlignCenter
Worksheets("Letters_Home").Range("A" & RowNumber & ":AE" & RowNumber + 3).VerticalAlignment = xlVAlignCenter
 
With Range("A" & RowNumber & ":AE" & RowNumber + 3).Font
            .Name = "Arial"
            .Size = 10
            .Bold = False
End With



''''''''''''''''''''''''''''''''''''Set row width and height to match the source sheet'''''''''''''''''''''''''''''''''''''''''''''''''''

Call CopyColumnWidths(Worksheets("Letters_Home").Range("B1:AG1"), Worksheets(theSheetRef).Range("B1:AG1"))
Range("B" & RowNumber).RowHeight = (Worksheets(theSheetRef).Range("A2").RowHeight - 15)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Reset the counter "i" to its beginning value in the merge loop
i = 1

'This for loop re-merges all cells that were merged in the original sheet.
For Each d In Range("A" & RowNumber & ":AJ" & RowNumber)

     
     If (d.Offset(1, 0).Value <> "Points") Then
        If (d.Offset(0, 1).Value = 0) Then
             i = i + 1
        Else
             j = (i - 1) * -1
             Range(d.Offset(0, j), d).Merge
             i = 1
             d.WrapText = True
                          
             With Range(d.Offset(0, j), d)
                
                'Fill cell with light grey
                .Interior.ColorIndex = 15
                
                'Set border around cell
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeTop).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlThin
                'Set font in cell to bold
                .Font.Bold = True

            End With
        End If


    Else
    d.Borders(xlEdgeRight).Weight = xlThin
        'This is the last merged skill section.
'        j = (i - 1) * -1
'        Range(d.Offset(0, j), d.Offset(0, -1)).Merge
'
'        Range(d.Offset(3, 0), d.Offset(2, 0)).ColumnWidth = 5.17
'
'
'        i = 1
'        d.WrapText = True
'
'        With Range(d.Offset(0, j), d.Offset(0, -1))
'
'                'Fill cell with light grey
'                .Interior.ColorIndex = 15
'
'                'Set border around cell
'                .Borders(xlEdgeLeft).Weight = xlThin
'                .Borders(xlEdgeRight).Weight = xlThin
'                .Borders(xlEdgeTop).Weight = xlMedium
'
'
'                'Set font in cell to bold
''                .Font.Bold = True
''
'        End With


    End If
    
      
    If (d.Offset(1, 0).Value = "%Correct") Then
        'Get rid of print button
        d.Offset(0, -1).Value = ""
        d.ColumnWidth = 8.14
        d.Offset(3, 0).NumberFormat = "0%"
          
          
        d.Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -1).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -2).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -3).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -4).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -5).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -6).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -7).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -8).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -9).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -10).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -11).Borders(xlEdgeTop).Weight = xlMedium
        
        
        d.Borders(xlEdgeBottom).Weight = xlThin
        d.Borders(xlEdgeRight).Weight = xlMedium
        
        
        d.Offset(0, -1).Borders(xlEdgeTop).Weight = xlMedium
        d.Offset(0, -1).Borders(xlEdgeBottom).Weight = xlThin

        d.Offset(1, 0).Borders(xlEdgeRight).Weight = xlMedium
        d.Offset(2, 0).Borders(xlEdgeRight).Weight = xlMedium
        d.Offset(3, 0).Borders(xlEdgeRight).Weight = xlMedium
        
    End If
    d.WrapText = True
    Next d

endflag = 0
For Each d In Range("A" & RowNumber + 3 & ":AF" & RowNumber + 3)
    If (endflag = 0) Then
        d.Borders(xlEdgeBottom).Weight = xlMedium
        d.Borders(xlEdgeTop).Weight = xlThin
        d.Offset(-1, 0).Borders(xlEdgeRight).Weight = xlThin
        d.Offset(-1, 0).Borders(xlEdgeTop).Weight = xlThin
        d.Offset(-2, 0).Borders(xlEdgeRight).Weight = xlThin
        d.Offset(-2, 0).Borders(xlEdgeTop).Weight = xlThin
        d.Borders(xlEdgeRight).Weight = xlThin
        If (d.NumberFormat = "0%") Then
           d.Borders(xlEdgeBottom).Weight = xlMedium
           d.Borders(xlEdgeTop).Weight = xlThin
           d.Borders(xlEdgeRight).Weight = xlMedium
           d.Offset(-1, 0).Borders(xlEdgeRight).Weight = xlMedium
           d.Offset(-1, 0).Borders(xlEdgeTop).Weight = xlThin
           d.Offset(-2, 0).Borders(xlEdgeRight).Weight = xlMedium
           d.Offset(-2, 0).Borders(xlEdgeTop).Weight = xlThin
           endflag = 1
        End If
    End If
Next d

Range("A" & RowNumber).Interior.ColorIndex = 0
Range("A" & RowNumber + 1).Borders(xlEdgeLeft).Weight = xlThin
Range("A" & RowNumber + 2).Borders(xlEdgeLeft).Weight = xlThin
Range("A" & RowNumber + 3).Borders(xlEdgeLeft).Weight = xlThin



End Sub

'============================================================================

Function Extract_Number_from_Text(Phrase As String) As Double
Dim Length_of_String As Integer
Dim Current_Pos As Integer
Dim Temp As String
Length_of_String = Len(Phrase)
Temp = ""
For Current_Pos = 1 To Length_of_String
If (Mid(Phrase, Current_Pos, 1) = "-") Then
  Temp = Temp & Mid(Phrase, Current_Pos, 1)
End If
If (Mid(Phrase, Current_Pos, 1) = ".") Then
 Temp = Temp & Mid(Phrase, Current_Pos, 1)
End If
If (IsNumeric(Mid(Phrase, Current_Pos, 1))) = True Then
    Temp = Temp & Mid(Phrase, Current_Pos, 1)
 End If
Next Current_Pos
If Len(Temp) = 0 Then
    Extract_Number_from_Text = 0
Else
    Extract_Number_from_Text = CDbl(Temp)
End If
End Function

'============================================================================
Sub createLetters()
Dim RangeTempVar As Range
Dim cellCounter As Integer
Dim sheetname As String
Dim studentName As String
Dim studentScore As Integer
Dim c As Range
Dim topRowFinder As Integer
'Dim unitDescriptionOffset As Integer
Dim sheetNumber As Integer
Dim nextSheet As String
Dim studentCounter As Integer
Dim wsSheet As Worksheet
'These variables are used in determining the percent done for the progress bar
Dim ClassSize As Integer
Dim PctDone As Single


Application.DisplayAlerts = False  'Prevent merge warning from popping up
Application.ScreenUpdating = False 'disabling screen updating should speed the macro up


Application.StatusBar = "Please Wait"

Range("A2:A1000").RowHeight = 14
'Range("b2:ac2").ColumnWidth = 4





'Clear out old values
Range("A2:AC1000").Select
    Selection.Delete Shift:=xlUp
sheetname = Range("A1").Value
studentCounter = 6
cellCounter = 2

sheetNumber = Extract_Number_from_Text(Range("A1").Value)

'clear out all old pagebreaks
ActiveSheet.Cells.PageBreak = xlPageBreakNone

'Set pagebreak to start After A1
Range("A2").PageBreak = xlPageBreakManual




nextSheet = "Unit " & (sheetNumber + 1)

Range("A1").ColumnWidth = 15.5



' determine class size
For Each c In Worksheets(sheetname).Range("A6:A42")
    If (c.End(xlToLeft).Value <> "") Then
    ClassSize = ClassSize + 1
    End If
    Next c


ClassSize = ClassSize + 6



For Each c In Worksheets(sheetname).Range("A6:A42")
    If (c.End(xlToLeft).Value <> "") Then
 
        'This will help us point to the cell with the unit
        topRowFinder = (studentCounter - 1) * -1
        
        'Insert the date in the first line
        Range("A" & cellCounter, "AJ" & cellCounter).Merge
        Range("A" & cellCounter).RowHeight = 25
        
        With Range("A" & cellCounter).Font
            .Name = "Arial"
            .Size = 15
            .Bold = False
        End With
        Range("A" & cellCounter).HorizontalAlignment = xlLeft
        
        
        
        
        
        Range("A" & cellCounter).Value = ("                                                                                       " & Format(Date, "Long Date"))
        
        'iterate cellCounter to go to the next cell.  Here, we will print the meat of the letter.
        cellCounter = cellCounter + 1
        
  
        'Merge cell with the writing
        Range("A" & cellCounter, "AC" & cellCounter + 3).Merge
        
        Range("A" & cellCounter).WrapText = True
        Range("A" & cellCounter).VerticalAlignment = xlVAlignTop
        Range("A" & cellCounter).RowHeight = 120
  
  
        With Range("A" & cellCounter).Font
            .Name = "Arial"
            .Size = 15
            .Bold = False
        End With
  
        If (c.End(xlToRight).Value <> "") Then
        
            If (c.End(xlToRight).Value >= 0.7945) Then
                Range("A" & cellCounter).Value = "Dear Parents of " & c.Value & "," & _
                vbNewLine & vbNewLine & "We just finished our " & c.Offset(topRowFinder, 3).Value & " test on " & _
                StrConv(c.Offset(topRowFinder, 4).Value, vbLowerCase) & ". " & vbNewLine & _
                Split(c.Value, " ")(0) & " scored " & c.End(xlToRight).Offset(0, -1).Value & " out of " & _
                c.End(xlToRight).Offset(topRowFinder + 3, -1).Value & _
                ", which is " & Round((c.End(xlToRight).Value * 100), 0) & "%!  Congratulations!!" & _
                vbNewLine & "Below, you will find a breakdown of your child's performance."
            Else
                Range("A" & cellCounter).Value = "Dear Parents of " & c.Value & "," & _
                vbNewLine & vbNewLine & "We just finished our " & c.Offset(topRowFinder, 3).Value & " test on " & _
                StrConv(c.Offset(topRowFinder, 4).Value, vbLowerCase) & ". " & vbNewLine & _
                Split(c.Value, " ")(0) & " scored " & c.End(xlToRight).Offset(0, -1).Value & " out of " & _
                c.End(xlToRight).Offset(topRowFinder + 3, -1).Value & _
                ", which is " & Round((c.End(xlToRight).Value * 100), 0) & "%." & _
                vbNewLine & "Below, you will find a breakdown of your child's performance."
        
            End If
        End If
                
        Call Breakdown(cellCounter + 4, sheetname, studentCounter)
        
        
        With Range("A" & cellCounter + 9)
            .RowHeight = 23.25
            .Font.Name = "Arial"
            .Font.Size = 15
            .Font.Bold = False
        End With
        
        
        ' This will test if it is the last unit, avoiding an error:
        On Error Resume Next
        Set wsSheet = Sheets(nextSheet)
        On Error GoTo 0
        If Not wsSheet Is Nothing Then
            Range("A" & cellCounter + 9).Value = "Next up is " & Worksheets(nextSheet).Range("D1").Value & ": " & Worksheets(nextSheet).Range("E1").Value & "."
        Else
            Range("A" & cellCounter + 9).Value = "This was our last unit.  It has been fantastic working with your child this year!"
        End If


       
        With Range("A" & cellCounter + 11)
            .VerticalAlignment = xlTop
            .Font.Name = "Arial"
            .Font.Size = 15
            .Font.Bold = False
        End With
        
        Range("A" & cellCounter + 11 & ":AC" & cellCounter + 14).Merge
        
        Range("A" & cellCounter + 11).Value = "Thanks," & vbNewLine & Worksheets("Info").Range("B1").Value



        With Range("A" & cellCounter + 16)
            .Font.Name = "Arial"
            .Font.Size = 15
            .Font.Bold = False
        End With
             
             
        Range("A" & cellCounter + 16 & ":AC" & cellCounter + 23).Merge
        
        Range("A" & cellCounter + 16).Value = "Please sign and return to indicate that you have reviewed this information:" & vbNewLine & vbNewLine & "X__________________________________"
        Range("A" & cellCounter + 16).Activate

        ActiveWindow.SelectedSheets.HPageBreaks.Add before:=ActiveCell.Offset(1, 0)


        cellCounter = cellCounter + 24
        studentCounter = studentCounter + 1
    End If
'-------------------------------------------
'            userform code

        PctDone = studentCounter / ClassSize
        With Progress_Window
            .FrameProgress.Caption = Format(PctDone, "0%")
            .LabelProgress.Width = PctDone * (.FrameProgress.Width - 10)
        End With
'       The DoEvents statement is responsible for the form's updating
        DoEvents
'--------------------------------------------
    
Next c

'Unload Processing_Dialog
Unload Progress_Window

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.StatusBar = False
Range("A1").Activate

End Sub


