Attribute VB_Name = "ActiveOrNot"
Sub ActiveOrNot()
    Dim Ws1 As Worksheet
    Dim Tracking_Workbook As Workbook
    Dim Ws2 As Worksheet
    Dim Last_Row_Ws1 As Long
    Dim Last_Row_Ws2 As Long
    Dim i As Long
    Dim Prophix_Range As Range
    Dim Current_Range As Range
    Dim Current_Value As Long
    Dim Ws1_Location As Range
    
    'Set the Project updates tracking worksheet and the current one downloaded from Prophix'
    Set Ws1 = ThisWorkbook.Sheets(2)
    
    Set Tracking_Workbook = Workbooks.Open("https://wilsonconst.sharepoint.com/sites/pwa/Shared Documents/Controls/Project Updates Tracking.xlsx")
    Set Ws2 = Tracking_Workbook.Sheets(1)

    'Find the lastrows in worksheet one and two'
    Last_Row_Ws1 = Ws1.Cells(Ws1.Rows.Count, "A").End(xlUp).Row
    Last_Row_Ws2 = Ws2.Cells(Ws2.Rows.Count, "C").End(xlUp).Row
    
    'Ensure the cells in prophix are ready for a 1:1 match in the project updates tracking worksheet'
    
    Ws1.Activate
    
    For i = 7 To Last_Row_Ws1
        Ws1.Cells(i, 1).Value = Left(Cells(i, 1), 4)
    Next i
    
    Set Tracking_Range = Ws2.Range("A2", Ws2.Cells(Last_Row_Ws2, 3))
    
    'If any of the values in Prophix_Range match the value from Column C of Project Updates Tracking, bold them'
    For i = 7 To Last_Row_Ws1
        Set Current_Range = Ws1.Cells(i, 1)
        Current_Value = Current_Range.Value
        Set Ws2_Location = Tracking_Range.Find(Current_Value)
        
        If Ws2_Location Is Nothing Then
          Ws1.Cells(i, 1).Interior.Color = RGB(255, 0, 0)
        End If
    Next i
    
    
End Sub
