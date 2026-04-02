Class TeacherClassesView
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetTeacherContext(teacherId As String, teacherName As String)
        ClassesTeacherIdTextBlock.Text = If(teacherId, String.Empty).Trim()
    End Sub
End Class
