Attribute VB_Name = "CutPaste"
Sub DisableCut()

On Error Resume Next
'Disable RigthClick on SheetCells which also gives you the option to cut
Application.CommandBars("Cell").Enabled = False

'Disable Cut button (old excel)
Application.CommandBars("Standard").Controls.Item("Cut").Enabled = False

'Disable Cut button (old excel)
Select Case Application.CutCopyMode
Case Is = False 'do nothing
Case Is = xlCopy 'do nothing
Case Is = xlCut
Application.CutCopyMode = False 'clear clipboard and cancel cut
End Select

'Disable Cut button
Application.CommandBars("Edit").Controls.Item("Cut").Enabled = False

'Divert Ctrl + X = Cut
Application.OnKey "^x", "NoNo"

'Disable Cell drag & Drop
Application.CellDragAndDrop = False
End Sub


Sub EnableCut()
Application.CommandBars("Edit").Controls.Item("Cut").Enabled = True
Application.CommandBars("Standard").Controls.Item("Cut").Enabled = True
Application.CommandBars("Cell").Enabled = True
Application.OnKey "^x"
Application.OnKey "{Delete}"
Application.CellDragAndDrop = True
End Sub

Sub NoNo()
Dim MyMsg As String
Dim Lf As String
Lf = Chr(13)
MyMsg = "Sorry! You cannot use the Cut feature." & Lf
MyMsg = MyMsg & "The integrety of the data may be" & Lf
MyMsg = MyMsg & "corrupted by this action.  Please" & Lf
MyMsg = MyMsg & "use Copy & Paste." & Lf & Lf
MyMsg = MyMsg & "Thanks!" & Lf
MyMsg = MyMsg & "-Anthony"
MsgBox MyMsg, vbInformation, "Data Integrity"
End Sub

