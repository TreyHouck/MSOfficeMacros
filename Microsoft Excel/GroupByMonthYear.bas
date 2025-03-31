Attribute VB_Name = "GroupbyMonth"
Sub GroupbyMonth()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dateCell As Range
    Dim dateValue As Date
    Dim prevGroupKey As String
    Dim groupKey As String
    Dim startRow As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet2") ' Change to your sheet name
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    startRow = 5 ' Assuming the first row is a header
 
    ' Loop through each date in column B
    For i = 2 To lastRow + 1 ' Adjusted to account for inserted rows
        Set dateCell = ws.Cells(i, 2)
        If IsDate(dateCell.Value) Or i = lastRow + 1 Then
            If i <= lastRow Then
                dateValue = dateCell.Value
                groupKey = Format(dateValue, "mmmm yyyy")
            End If
            
            ' Check for change in month or year or end of data
            If i > 2 And (groupKey <> prevGroupKey Or i = lastRow + 1) Then
                ws.Rows(i).Insert Shift:=xlDown
                ws.Cells(i, 8).Formula = "=SUM(H" & startRow & ":H" & i - 1 & ")"
                ws.Cells(i, 8).Font.Bold = True ' Make the sum bold
                lastRow = lastRow + 1
                i = i + 1
                startRow = i ' Adjusted to include the first date in the next set
            End If
            
            prevGroupKey = groupKey

        End If
    Next i
End Sub

