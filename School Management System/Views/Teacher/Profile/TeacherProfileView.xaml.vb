Class TeacherProfileView
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetTeacherContext(teacherId As String, teacherName As String)
        Dim normalizedId As String = If(teacherId, String.Empty).Trim()
        Dim normalizedName As String = If(teacherName, String.Empty).Trim()

        ProfileTeacherNameTextBlock.Text = If(String.IsNullOrWhiteSpace(normalizedName), "Teacher", normalizedName)
        ProfileTeacherIdTextBlock.Text = normalizedId
        ProfileTeacherEmployeeIdValueTextBlock.Text = If(String.IsNullOrWhiteSpace(normalizedId), "Not set", normalizedId)
        ProfileTeacherNameValueTextBlock.Text = If(String.IsNullOrWhiteSpace(normalizedName), "Not set", normalizedName)

        If String.IsNullOrWhiteSpace(normalizedName) Then
            ProfileTeacherInitialTextBlock.Text = "T"
        Else
            ProfileTeacherInitialTextBlock.Text = normalizedName.Substring(0, 1).ToUpperInvariant()
        End If
    End Sub
End Class
