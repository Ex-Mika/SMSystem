Class StudentScheduleView
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedId As String = If(studentId, String.Empty).Trim()
        ScheduleStudentIdTextBlock.Text = normalizedId
        WeeklyScheduleTabView.SetStudentContext(normalizedId, studentName)
    End Sub
End Class
