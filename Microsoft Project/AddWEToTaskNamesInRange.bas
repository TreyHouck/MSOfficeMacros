Attribute VB_Name = "Module1"
Sub AddWEToTaskNamesInRange()
    Dim startID As Long
    Dim endID As Long
    Dim t As Task
    
    ' Prompt user for the range of task IDs
    startID = InputBox("Enter the starting task ID:")
    endID = InputBox("Enter the ending task ID:")
    
    ' Loop through the specified range of tasks
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.ID >= startID And t.ID <= endID Then
                t.Name = "WE " & t.Name
            End If
        End If
    Next t
End Sub
