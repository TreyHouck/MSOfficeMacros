Attribute VB_Name = "Module1"
Sub AddApostrophe()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    
    'Set the worksheet and table
    Set ws = ThisWorkbook.Sheets("Tracking")
    Set tbl = ws.ListObjects("Tracking")
    
    'Loop through each cell in column "Date Last Contacted" of the table
    For Each cell In tbl.ListColumns("Date Last Contacted").DataBodyRange
        If cell.Value <> "" Then
            If Left(cell.Value, 1) <> "'" Then
                cell.Value = "'" & cell.Value
            End If
        End If
    Next cell
End Sub
