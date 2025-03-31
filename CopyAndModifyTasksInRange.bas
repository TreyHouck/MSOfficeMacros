Attribute VB_Name = "CopyAndModifyTasksInRange"
Sub CopyAndModifyTasksInRange()
    Dim t As Task
    Dim newTask As Task
    Dim proj As Project
    Dim taskName As String
    Dim startID As Integer
    Dim endID As Integer
    
    'Define the range of task IDs to process
    startID = InputBox("Enter the start ID of the range:", "Start ID")
    endID = InputBox("Enter the end ID of the range:", "End ID")
    
    Set proj = ActiveProject
    
    For Each t In proj.Tasks
        If Not t Is Nothing Then
            ' Check if the task is within the specified ID range and is not a summary task
            If t.ID >= startID And t.ID <= endID And t.Summary = False Then
                ' Create a new task below the original task
                Set newTask = proj.Tasks.Add(t.Name, t.ID + 1)
                ' Copy task details
                newTask.Start = t.Start
                newTask.Finish = t.Finish
                newTask.Duration = t.Duration
                newTask.ResourceNames = t.ResourceNames
                ' Modify the task name to omit "Load, Haul, "
                taskName = Replace(t.Name, "Load, Haul, ", "")
                newTask.Name = taskName
            End If
        End If
    Next t
End Sub
