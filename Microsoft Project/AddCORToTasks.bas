Attribute VB_Name = "AddCORToTasks"
Sub AddCORToTasks()
    Dim t As Task
    Dim startID As Integer
    Dim endID As Integer
    Dim taskID As Integer
    
    ' Prompt user for the range of task IDs
    startID = InputBox("Enter the starting task ID:")
    endID = InputBox("Enter the ending task ID:")
    
    ' Loop through the specified range of tasks
    For taskID = startID To endID
        Set t = ActiveProject.Tasks(taskID)
        If Not t Is Nothing Then
            t.Name = "COR " & t.Name
        End If
    Next taskID
End Sub
