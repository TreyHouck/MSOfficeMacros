Attribute VB_Name = "AdamSchedulePrep"
Sub AdamSchedulePrep()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    Dim emptyRowCount As Long
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Unmerge all merged cells in the worksheet
    For Each cell In ws.UsedRange
        If cell.MergeCells Then
            cell.MergeArea.UnMerge
        End If
    Next cell

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
    
    ' Search for the first row with an underlined cell
    foundUnderlined = False
    For rowCount = 1 To lastRow
        For Each cell In ws.Rows(rowCount).Cells
            If cell.Font.Underline <> xlUnderlineStyleNone Then
                firstUnderlinedRow = rowCount
                foundUnderlined = True
                Exit For
            End If
        Next cell
        If foundUnderlined Then Exit For
    Next rowCount
    
    ' Move the first row with an underlined cell to the top
    If foundUnderlined Then
        ws.Rows(firstUnderlinedRow).Cut
        ws.Rows(1).Insert Shift:=xlDown
    End If
    
    ' Delete all other rows with underlined cells
    For rowCount = lastRow To 2 Step -1
        For Each cell In ws.Rows(rowCount).Cells
            If cell.Font.Underline <> xlUnderlineStyleNone Then
                ws.Rows(rowCount).Delete
                Exit For
            End If
        Next cell
    Next rowCount
    
End Sub

