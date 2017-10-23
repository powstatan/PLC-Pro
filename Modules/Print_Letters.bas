Attribute VB_Name = "Print_Letters"
Option Explicit

Sub PrintLetters(PrintKey As String)
Dim CF As Long, CV As Long, RF As Long, RV As Long
Dim Col As Long, Rw As Long

With ActiveSheet

'Make buttons invisible for printing
'This is currently unnecessary, but left here
'in case I want to put the print key(s) somewhere
'else on the page.  Plus, it looks cool.

Range(PrintKey).Interior.Color = RGB(255, 255, 255)
Range("M1").Interior.Color = RGB(255, 255, 255)

.PageSetup.Orientation = xlLandscape

'Print all pages, skipping the page with the
'unit selector and print keys.
.PrintOut From:=2, To:=50, Copies:=1

'Reset button colors to blue
Range(PrintKey).Interior.Color = RGB(0, 112, 196)
Range("M1").Interior.Color = RGB(0, 112, 196)

'Point away from the print key in case the user wants to click it again.
Range("AC1").Activate

End With

End Sub

