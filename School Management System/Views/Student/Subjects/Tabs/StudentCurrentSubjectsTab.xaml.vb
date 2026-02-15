Class StudentCurrentSubjectsTab
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        CurrentStudentIdTextBlock.Text = If(studentId, String.Empty).Trim()
    End Sub
End Class
