Attribute VB_Name = "Zoom"
Sub Zoomer()
    Range("A1").Select
    Range("B6").Select
    ActiveWindow.Zoom = 100
End Sub
Sub ZoomerInfo()
    Range("B1").Select
    ActiveWindow.Zoom = 100
End Sub
Sub ZoomerLetters()
    Range("A1").Select
    ActiveWindow.Zoom = 115
End Sub
Sub ZoomerSummary()
    Range("A1").Select
    ActiveWindow.Zoom = 130
End Sub
Sub ZoomerAnnouncements()
    Range("A1").Select
    ActiveWindow.Zoom = 130
End Sub
