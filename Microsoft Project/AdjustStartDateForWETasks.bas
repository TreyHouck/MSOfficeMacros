Attribute VB_Name = "Module2"
Sub AdjustStartDateForWETasks()
    Dim t As Task
    Dim taskDate As Date
    Dim taskName As String
    Dim datePosition As Integer
    
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            taskName = t.Name
            If InStr(taskName, "WE") > 0 Then
                ' Extract the date from the task name
                datePosition = InStr(taskName, "WE") + 3
                On Error Resume Next
                taskDate = CDate(Mid(taskName, datePosition, 10))
                On Error GoTo 0
                
                ' If a valid date is found, adjust the start date
                If IsDate(taskDate) Then
                    t.Start = taskDate - 6
                End If
            End If
        End If
    Next t
End Sub
