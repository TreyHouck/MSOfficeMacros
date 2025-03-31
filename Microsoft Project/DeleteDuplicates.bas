Attribute VB_Name = "DeleteDuplicates"
Sub DeleteDuplicateTasksInRange()
    Dim t As Task
    Dim proj As Project
    Dim taskNames As Collection
    Dim taskName As String
    Dim startID As Integer
    Dim endID As Integer
    
    ' Prompt the user to input the range of task IDs to process
    startID = InputBox("Enter the start ID of the range:", "Start ID")
    endID = InputBox("Enter the end ID of the range:", "End ID")
    
    Set proj = ActiveProject
    Set taskNames = New Collection
    
    On Error Resume Next
    
    ' Iterate through all tasks in the project
    For Each t In proj.Tasks
        If Not t Is Nothing Then
            ' Check if the task is within the specified ID range
            If t.ID >= startID And t.ID <= endID Then
                taskName = t.Name
                
                ' Check if the task name already exists in the collection
                If taskName <> "" Then
                    taskNames.Add taskName, taskName
                    If Err.Number <> 0 Then
                        ' If the task name already exists, delete the duplicate task
                        t.Delete
                        Err.Clear
                    End If
                End If
            End If
        End If
    Next t
    
    On Error GoTo 0
End Sub
