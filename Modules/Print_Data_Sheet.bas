Attribute VB_Name = "Print_Data_Sheet"
Sub PrintDataSheet(PrintKeyLocation As String)
Dim CF As Long, CV As Long, RF As Long, RV As Long
Dim Col As Long, Rw As Long




With ActiveSheet

DisplayPageBreaks = False

CV = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

RV = .Cells.Find(What:="*", After:=Range("A1"), LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

Col = Application.WorksheetFunction.Max(CF, CV)
Rw = Application.WorksheetFunction.Max(RF, RV)

.PageSetup.PrintArea = "$A$1:" & Cells(Rw, Col).Address


'Make button invisible for printing
Range(PrintKeyLocation).Interior.Color = RGB(255, 255, 255)

Application.Dialogs(xlDialogPrint).Show

Range(PrintKeyLocation).Interior.Color = RGB(0, 112, 196)
'Activate another cell so that you can click on the button again
Range("A1").Activate

End With




End Sub


