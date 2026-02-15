Class StudentWeeklyScheduleTab
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        WeeklyScheduleStudentIdTextBlock.Text = If(studentId, String.Empty).Trim()
    End Sub
End Class
