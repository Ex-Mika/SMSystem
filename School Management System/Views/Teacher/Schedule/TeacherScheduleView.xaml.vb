Class TeacherScheduleView
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetTeacherContext(teacherId As String, teacherName As String)
        ScheduleTeacherIdTextBlock.Text = If(teacherId, String.Empty).Trim()
    End Sub
End Class
