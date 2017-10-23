VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Call ZoomerSummary
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True   'turn off Edit mode when using ÒEdit directly in a cellÓ
      If Not Intersect(Range("A2"), Target) Is Nothing Then
          Call PrintArea
      End If
End Sub


Sub PrintArea()
Dim CF As Long, CV As Long, RF As Long, RV As Long
Dim Col As Long, Rw As Long

With ActiveSheet

'Set page to landscape printing
.PageSetup.Orientation = xlLandscape

'Find the limits in order to set the print area
CV = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

RV = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

Col = Application.WorksheetFunction.Max(CF, CV)
Rw = Application.WorksheetFunction.Max(RF, RV)

.PageSetup.PrintArea = "$A$1:" & Cells(Rw, Col).Address
  
     .PageSetup.Zoom = False
     .PageSetup.FitToPagesWide = 1
     .PageSetup.FitToPagesTall = 1


'Make print button invisible for printing
Range("A2").Interior.Color = RGB(255, 255, 255)

Application.Dialogs(xlDialogPrint).Show

Range("A2").Interior.Color = RGB(0, 112, 196)

Range("A1").Activate

End With




End Sub
Public Sub ShrinkColumns()
Columns("G:N").ColumnWidth = 6.71
End Sub

Public Sub AutoFitSheet()

    If ActiveWorkbook Is Nothing Then Exit Sub
    Dim i#
    If ActiveWindow.SelectedSheets.Count > 1 Then
        For i = 1 To ActiveWindow.SelectedSheets.Count
            ActiveWindow.SelectedSheets(i).Cells.EntireColumn.AutoFit
        Next
    Else
        Cells.EntireColumn.AutoFit
    End If

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
   If Not Application.Intersect(Target, Range("a1:b1")) Is Nothing Then
      Application.EnableEvents = False


Dim c As Range
Dim sheetname As String
Dim numProf As Integer
Dim Qone As Integer
Dim Qtwo As Integer
Dim Qthree As Integer
Dim Qfour As Integer

numProf = 0
Qone = 5
Qtwo = 5
Qthree = 5
Qfour = 5

sheetname = Range("A1").Value

'Clear out old data
Range("G5: N35").Delete
Range("G4:N4").Value = ""


'Start by shrinking the quartile columns as much as possible
Call ShrinkColumns


Range("G4").Value = "0 Students"
Range("I4").Value = "0 Students"
Range("K4").Value = "0 Students"
Range("M4").Value = "0 Students"

If sheetname <> "Select Unit" Then


    For Each c In Worksheets(sheetname).Range("A6:A34")
        If (c.Value <> "") Then
            If (c.End(xlToRight).Value <> "") Then
            
                'Quartile 1
                If (c.End(xlToRight).Value < 0.255) Then
           
                'Name of student
                Range("G" & Qone).Value = c.Value
            
                'Student's score
                Range("H" & Qone).Value = c.End(xlToRight).Value
            
                     
            
                'Formatting
                Range("H" & Qone).NumberFormat = "0%"
                Range("H" & Qone).HorizontalAlignment = xlCenter
                Range("H" & Qone).VerticalAlignment = xlBottom
             
             
                'Set left border on students' names
                With Range("G" & Qone).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                'Set right border on students' scores
                With Range("H" & Qone).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With

                'Aggregate the number of students count in the header
                If ((Qone - 4) = 1) Then
                    Range("G4").Value = (Qone - 4) & " Student"
                Else
                    Range("G4").Value = (Qone - 4) & " Students"
                End If
            
                'Set counter to move down one cell
                Qone = Qone + 1
            End If

            'Quartile 2
            If (c.End(xlToRight).Value < 0.495) Then
        
                If (c.End(xlToRight).Value >= 0.255) Then
                
                    'Name of student
                    Range("I" & Qtwo).Value = c.Value
                
                    'Student's score
                    Range("J" & Qtwo).Value = c.End(xlToRight).Value
            
                'Formatting
                Range("J" & Qtwo).NumberFormat = "0%"
                Range("J" & Qtwo).HorizontalAlignment = xlCenter
                Range("J" & Qtwo).VerticalAlignment = xlBottom
             
                'Set left border on students' names
                With Range("I" & Qtwo).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                'Set right border on students' scores
                With Range("J" & Qtwo).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
            
                'Aggregate the number of students count in the header
                If ((Qtwo - 4) = 1) Then
                    Range("I4").Value = (Qtwo - 4) & " Student"
                Else
                    Range("I4").Value = (Qtwo - 4) & " Students"
                End If
            
            
                'Set counter to move down one cell
                Qtwo = Qtwo + 1
               
                End If
             
             
            End If
        
        
            'Quartile 3
            If (c.End(xlToRight).Value < 0.755) Then
        
                If (c.End(xlToRight).Value >= 0.495) Then
                
                    'Name of student
                    Range("K" & Qthree).Value = c.Value

                    'Student's score
                    Range("L" & Qthree).Value = c.End(xlToRight).Value
            
                'Formatting
                Range("L" & Qthree).NumberFormat = "0%"
                Range("L" & Qthree).HorizontalAlignment = xlCenter
                Range("L" & Qthree).VerticalAlignment = xlBottom
            
                'Set left border on students' names
                With Range("K" & Qthree).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                'Set right border on students' scores
                With Range("L" & Qthree).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With


                'Aggregate the number of students count in the header
                If ((Qthree - 4) = 1) Then
                    Range("K4").Value = (Qthree - 4) & " Student"
                Else
                    Range("K4").Value = (Qthree - 4) & " Students"
                End If
            
            
                'Set counter to move down one cell
                Qthree = Qthree + 1
               
                End If

            End If

            'Quartile 4
            If c.End(xlToRight).Value >= 0.755 Then
            
                'Name of student
                Range("M" & Qfour).Value = c.Value
            
                'Student's score
                Range("N" & Qfour).Value = c.End(xlToRight).Value
            
                'Formatting
                Range("N" & Qfour).NumberFormat = "0%"
                Range("N" & Qfour).HorizontalAlignment = xlCenter
                Range("N" & Qfour).VerticalAlignment = xlBottom

            'Set left border on students' names
                With Range("M" & Qfour).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
            
                'Set right border on students' scores
                With Range("N" & Qfour).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With

                'Aggregate the number of students count in the header
                If ((Qfour - 4) = 1) Then
                    Range("M4").Value = (Qfour - 4) & " Student"
                Else
                    Range("M4").Value = (Qfour - 4) & " Students"
                End If
            
            
                'Set counter to move down one cell
                Qfour = Qfour + 1

            End If
        
        'Set the number of proficient students (85% or better) in cell C5
        Range("C5").Value = WorksheetFunction.CountIf(Range("N5:N36"), ">= 0.845")
    End If
    End If
    
    Next c

    'Set bottom border for the last name and score in each quartile
  
                With Range("G" & Qone - 1).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("H" & Qone - 1).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("I" & Qtwo - 1).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("J" & Qtwo - 1).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("K" & Qthree - 1).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("L" & Qthree - 1).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("M" & Qfour - 1).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("N" & Qfour - 1).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
            
    'Reset the bottom border for the header cells
                With Range("G4").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("H4").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("I4").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("J4").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("K4").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("L4").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("M4").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With Range("N4").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
    
    Else
    Range("C5").Value = ""
    End If

End If
   
   Application.EnableEvents = True

'Eliminate wasted space by autofitting all columns
Call AutoFitSheet

'Sort the columns
Call SortSummaryData


'Reset the view to the top-leftmost portion of the sheet
Range("A1").Activate

    
End Sub

