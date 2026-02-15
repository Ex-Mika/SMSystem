Class StudentProfileView
    Private Enum StudentProfileSection
        Personal
        Academic
        Account
    End Enum

    Public Sub New()
        InitializeComponent()
        SetActiveSection(StudentProfileSection.Personal)
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedId As String = If(studentId, String.Empty).Trim()
        Dim normalizedName As String = If(studentName, String.Empty).Trim()
        Dim displayName As String = If(String.IsNullOrWhiteSpace(normalizedName), "Student", normalizedName)
        Dim displayId As String = If(String.IsNullOrWhiteSpace(normalizedId), "No Student ID", normalizedId)

        ProfileHeaderNameTextBlock.Text = displayName
        ProfileHeaderIdTextBlock.Text = displayId
        ProfileHeaderInitialTextBlock.Text = GetProfileInitial(displayName)
        PersonalTabView.SetStudentContext(normalizedId, normalizedName)
    End Sub

    Private Sub PersonalSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentProfileSection.Personal)
    End Sub

    Private Sub AcademicSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentProfileSection.Academic)
    End Sub

    Private Sub AccountSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentProfileSection.Account)
    End Sub

    Private Sub SetActiveSection(section As StudentProfileSection)
        PersonalTabView.Visibility = If(section = StudentProfileSection.Personal, Visibility.Visible, Visibility.Collapsed)
        AcademicTabView.Visibility = If(section = StudentProfileSection.Academic, Visibility.Visible, Visibility.Collapsed)
        AccountTabView.Visibility = If(section = StudentProfileSection.Account, Visibility.Visible, Visibility.Collapsed)

        ApplySectionButtonState(PersonalSectionButton, section = StudentProfileSection.Personal)
        ApplySectionButtonState(AcademicSectionButton, section = StudentProfileSection.Academic)
        ApplySectionButtonState(AccountSectionButton, section = StudentProfileSection.Account)
    End Sub

    Private Sub ApplySectionButtonState(sectionButton As Button, isSelected As Boolean)
        sectionButton.Style = CType(FindResource(If(isSelected,
                                                    "DashboardProfileSegmentSelectedButtonStyle",
                                                    "DashboardProfileSegmentButtonStyle")), Style)
    End Sub

    Private Function GetProfileInitial(fullName As String) As String
        Dim normalizedName As String = If(fullName, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedName) Then
            Return "S"
        End If

        Return normalizedName.Substring(0, 1).ToUpperInvariant()
    End Function
End Class
