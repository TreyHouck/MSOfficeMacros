Attribute VB_Name = "DeleteEmptyRows"
Sub DeleteEmptyRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    Dim emptyRowCount As Long
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize the empty row counter
    emptyRowCount = 0
    
    ' Loop through the rows from the bottom up
    For rowCount = lastRow To 1 Step -1
        ' Check if the entire row is empty
        If Application.WorksheetFunction.CountA(ws.Rows(rowCount)) = 0 Then
            emptyRowCount = emptyRowCount + 1
            ' If two or more consecutive empty rows are found, exit the loop
            If emptyRowCount >= 2 Then Exit For
            ' Delete the empty row
            ws.Rows(rowCount).Delete
        Else
            ' Reset the empty row counter if a non-empty row is found
            emptyRowCount = 0
        End If
    Next rowCount
End Sub

