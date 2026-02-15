Class StudentSubjectsView
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedId As String = If(studentId, String.Empty).Trim()
        Dim normalizedName As String = If(studentName, String.Empty).Trim()

        SubjectsHeaderNameTextBlock.Text = If(String.IsNullOrWhiteSpace(normalizedName), "Student", normalizedName)
        SubjectsHeaderIdTextBlock.Text = normalizedId
        CurrentSubjectsTabView.SetStudentContext(normalizedId, normalizedName)
    End Sub
End Class
