VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Call ZoomerAnnouncements
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


Sub PrintAnnouncements(PrintKey As String)
Dim RF As Long, RV As Long
Dim Rw As Long




With ActiveSheet

RV = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

Rw = Application.WorksheetFunction.Max(RF, RV)

.PageSetup.PrintArea = "$A$3:" & Cells(Rw, 5).Address


' This was to make the document as big as the page.  Couldn't get it to work.
'     PageSetup.Zoom = False
'     PageSetup.PaperSize = xlPaperLegal
'     PageSetup.FitToPagesWide = 1
'     PageSetup.FitToPagesTall = 1


'Make button invisible for printing
Range(PrintKey).Interior.Color = RGB(255, 255, 255)



Application.Dialogs(xlDialogPrint).Show

Range(PrintKey).Interior.Color = RGB(0, 112, 196)

Range("A1").Activate

End With




End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True   'turn off Edit mode when using “Edit directly in a cell”
      If Not Intersect(Range("B1"), Target) Is Nothing Then
          Call PrintAnnouncements("B1")
      End If
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
   If Not Application.Intersect(Target, Range("a1:b1")) Is Nothing And Range("a1").Value <> "Select Unit" Then
         Application.EnableEvents = False
      
'Start by shrinking the columns as much as possible
Columns("A:AC").ColumnWidth = 6.71


Dim c As Range
Dim sheetname As String
Dim theSheet As Worksheet

Dim EightyFiveCounter As Integer
Dim HundoCounter As Integer

EightyFiveCounter = 5
HundoCounter = 5

sheetname = Range("A1").Value
'theSheet.Name = sheetname

'Clear out old data
Range("A5: D46").Delete

  
For Each c In Worksheets(sheetname).Range("A6:A34")


        'Students who scored 85 percent:
          If c.End(xlToRight).Value >= 0.795 Then
            If c.End(xlToRight).Value < 1 Then
                'Name of student
                Range("A" & EightyFiveCounter).Value = c.Value


                'Set counter to move down one cell
                 EightyFiveCounter = EightyFiveCounter + 1
                 End If
                 

        End If


        'Students who scored 100 percent:
        If (c.End(xlToRight).Value = 1) Then
        
                
                'Name of student
                Range("C" & HundoCounter).Value = c.Value
                
                'Student's score
             '   Range("E" & HundoCounter).Value = c.End(xlToRight).Value
            
            'Formatting
         '   Range("E" & HundoCounter).NumberFormat = "0%"
         '   Range("E" & HundoCounter).HorizontalAlignment = xlCenter
         '   Range("E" & HundoCounter).VerticalAlignment = xlBottom
            
            'Set counter to move down one cell
            HundoCounter = HundoCounter + 1
             
        End If
Next c


 Range("A5:C35").Select
         With Selection.Font
             .Size = 14
         End With

   End If
   Application.EnableEvents = True

Range("A1").Activate
Call AutoFitSheet
End Sub






