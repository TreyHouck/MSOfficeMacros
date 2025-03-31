Attribute VB_Name = "ChangeOrderReviewPrep"
Sub CheckAndPlaceZero()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheet you want to work with
    Set ws = ThisWorkbook.Sheets("1. Contract-Hours Comparison T")

    ' Find the last row with data in column J or K
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "K").End(xlUp).Row > lastRow Then
        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    End If

    ' Loop through each row and check the values in columns J and K
    For i = 1 To lastRow
        If ws.Cells(i, 10).Value <> "" And ws.Cells(i, 11).Value = "" Then
            ws.Cells(i, 11).Value = 0
        ElseIf ws.Cells(i, 11).Value <> "" And ws.Cells(i, 10).Value = "" Then
            ws.Cells(i, 10).Value = 0
        End If
    Next i

    ' Loop through each row again to check if both columns J and K are 0
    For i = 1 To lastRow
        If ws.Cells(i, 10).Value = 0 And ws.Cells(i, 11).Value = 0 Then
            ws.Cells(i, 10).Value = ""
            ws.Cells(i, 11).Value = ""
        End If
    Next i
End Sub
