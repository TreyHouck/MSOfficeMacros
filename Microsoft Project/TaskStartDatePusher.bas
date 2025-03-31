Attribute VB_Name = "TaskStartDatePusher"
Sub SetStartDateForZeroPercentComplete()
    Dim t As Task
    Dim today As Date
    today = Date
    
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.PercentComplete = 0 And t.Start <= today Then
            t.Start = today
            End If
        End If
    Next t
End Sub
