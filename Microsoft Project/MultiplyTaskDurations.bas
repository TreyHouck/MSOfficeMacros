Attribute VB_Name = "Module1"
Sub MultiplyTaskDurations()
    Dim t As Task
    Dim multiplier As Double
    Dim startID As Long
    Dim endID As Long
    Dim taskID As Long
    
    ' Prompt user for the multiplier value
    multiplier = InputBox("Enter the multiplier value:", "Multiply Task Durations")
    
    ' Prompt user for the start and end task IDs
    startID = InputBox("Enter the start task ID:", "Multiply Task Durations")
    endID = InputBox("Enter the end task ID:", "Multiply Task Durations")
    
    ' Loop through the specified range of tasks
    For taskID = startID To endID
        Set t = ActiveProject.Tasks(taskID)
        If Not t Is Nothing Then
            t.Duration = t.Duration * multiplier
        End If
    Next taskID
    
    MsgBox "Task durations have been multiplied by " & multiplier, vbInformation
End Sub
