Attribute VB_Name = "Sort_Summary_Data"
Sub SortSummaryData()
'
' SortSummaryData Macro
' Macro written 4/8/2011 by Anthony Schultz
'

'




    Range("G5:H35").Select
    Selection.Sort Key1:=Range("H5"), Order1:=xlDescending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("I5:J35").Select
    Range("I35").Activate
    Selection.Sort Key1:=Range("J5"), Order1:=xlDescending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("K5:L35").Select
    Range("K35").Activate
    Selection.Sort Key1:=Range("K5"), Order1:=xlDescending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("M5:N35").Select
    Range("M35").Activate
    Selection.Sort Key1:=Range("N5"), Order1:=xlDescending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal




End Sub


