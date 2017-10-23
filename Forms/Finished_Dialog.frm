VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Finished_Dialog 
   Caption         =   "Letter Creation Complete"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4695
   OleObjectBlob   =   "Finished_Dialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Finished_Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Call PrintLetters("B1")
Me.Hide
End Sub
Private Sub CommandButton2_Click()
Me.Hide
End Sub
