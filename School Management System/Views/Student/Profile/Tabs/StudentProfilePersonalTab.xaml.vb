Class StudentProfilePersonalTab
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedId As String = If(studentId, String.Empty).Trim()
        Dim normalizedName As String = If(studentName, String.Empty).Trim()

        StudentIdValueTextBlock.Text = If(String.IsNullOrWhiteSpace(normalizedId), "No Student ID", normalizedId)
        StudentNameValueTextBlock.Text = If(String.IsNullOrWhiteSpace(normalizedName), "Student", normalizedName)
    End Sub
End Class
