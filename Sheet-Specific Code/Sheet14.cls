VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Call ZoomerLetters
End Sub

'If you change A1, it recreates the letters by calling CreateLetters
Private Sub Worksheet_Change(ByVal Target As Range)

   If Not Application.Intersect(Target, Range("A1:B1")) Is Nothing And (Range("A1").Value <> "Select Unit") Then
        Call ShowDialog
        Call ShowFinishedDialog
   End If
   
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True   'turn off Edit mode when using “Edit directly in a cell”
      'Print all with default printer
      If Not Intersect(Range("B1"), Target) Is Nothing Then
          Call PrintLetters("B1")
      End If
      
      If Not Intersect(Range("M1"), Target) Is Nothing Then
            
            'Hide print Keys
            Range("B1").Interior.Color = RGB(255, 255, 255)
            Range("M1").Interior.Color = RGB(255, 255, 255)
            
            'Bring up print dialogue
            Application.Dialogs(xlDialogPrint).Show
            
            'Reset button colors to blue
            Range("B1").Interior.Color = RGB(0, 112, 196)
            Range("M1").Interior.Color = RGB(0, 112, 196)
            
            'Point away from the print key in case the user wants to click it again.
            Range("AC1").Activate
      End If
End Sub

