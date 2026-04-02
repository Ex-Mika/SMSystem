Class StudentGradeHistoryTab
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        GradeHistoryStudentIdTextBlock.Text = If(studentId, String.Empty).Trim()
    End Sub
End Class
