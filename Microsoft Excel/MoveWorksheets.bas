Attribute VB_Name = "MoveWorksheets"
Sub MoveWorksheets()
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim ws As Worksheet
    Dim destinationPath As String
    Dim fd As FileDialog
    
    'Set the source workbook
    Set sourceWorkbook = ActiveWorkbook
    
    'Create a FileDialog object as a File Picker dialog box
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'Set the dialog box title and filter
    With fd
        .Title = "Select the destination Workbook"
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xlsb; *.xls"
        .AllowMultiSelect = False
        If .Show = -1 Then 'If the user selects a file
            destinationPath = .SelectedItems(1)
        Else
            MsgBox "No file selected. Exiting macro."
            Exit Sub
        End If
    End With
    
    'Extract the workbook name from the path
    wbName = Dir(destinationPath)
    
    'Check if the destination workbook is already open
    On Error Resume Next
    Set destinationWorkboook = Workbooks(wbName)
    On Error GoTo 0
    
    'If the workbook is not open, open it
    If destinationWorkbook Is Nothing Then
        Set destinationWorkbook = Workbooks.Open(destinationPath)
    End If
    
    'Loop through each worksheet in the source workbook
    For Each ws In sourceWorkbook.Worksheets
        'Move the worksheet to the destination workbook
        ws.Move After:=destinationWorkbook.Sheets(destinationWorkbook.Sheets.Count)
    Next ws
    
    'Save and close the workbooks
    destinationWorkbook.Save
    sourceWorkbook.Close (SaveChanges = False)
    
    
End Sub
