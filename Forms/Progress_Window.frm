VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Progress_Window 
   Caption         =   "Progress"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8685
   OleObjectBlob   =   "Progress_Window.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Progress_Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
Unload Me
End
End Sub

Private Sub UserForm_activate()
    Call createLetters
End Sub
