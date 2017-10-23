VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Unlocker()
Dim wSheet As Worksheet
    Application.ScreenUpdating = False
    For Each wSheet In Worksheets
        'Unlock all pages
        wSheet.Unprotect Password:="doncarlos"
    Next wSheet
    Application.ScreenUpdating = True
    Call EnableCut
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call EnableCut
End Sub


Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, _
ByVal Target As Excel.Range)

Select Case Application.CutCopyMode
Case Is = False 'do nothing
Case Is = xlCopy 'do nothing
Case Is = xlCut
MsgBox "Please DO NOT Cut and Paste. Use Copy and Paste; then delete souce."
Application.CutCopyMode = False 'clear clipboard and cancel cut
End Select

End Sub


Private Sub Workbook_Open()
Dim wSheet As Worksheet
Application.ScreenUpdating = False
    For Each wSheet In Worksheets
        'Set the editing password for all sheets
        wSheet.Protect Password:="doncarlos", UserInterFaceOnly:=True
        wSheet.EnableSelection = xlUnlockedCells
    Next wSheet
'This will disable the cut feature, which can screw things up.
Call DisableCut
Application.ScreenUpdating = True
End Sub

Private Sub createFootersHeaders()
Dim wSheet As Worksheet
Application.ScreenUpdating = False
        'Change the headers and footers.
        'This really lags and should only be used when making a change!!
'  For Each wSheet In Worksheets

'            If InStr(wSheet.Name, "Unit") Then
'
'                With wSheet.PageSetup
'                    .LeftHeader = ""
'                    .CenterHeader = "&""Times New Roman,bold""&12" & "Mathematics" & vbLf & "Skills Analysis Sheet for " & wSheet.Range("E1").Value & vbLf & wSheet.Range("B1").Value & ", &A"
'                    .RightHeader = ""
'                   .LeftFooter = "&""Times New Roman""&8" & wSheet.Range("B1").Value & ", &A" & vbLf & wSheet.Range("E1").Value & vbLf & "FOR OPTIONAL USE"
'                   .CenterFooter = "&""Times New Roman""&8" & "Created by Anthony Schultz"
'                   .RightFooter = "&""Times New Roman""&8" & "Baltimore County Public Schools" & vbLf & "Office of Mathematics PreK-12" & vbLf & "October 2011"
'               End With
'            End If

'  Next wSheet
Application.ScreenUpdating = True
End Sub


