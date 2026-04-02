Class StudentSubjectsView
    Private Enum StudentSubjectsSection
        CurrentLoad
        Completed
        Planned
    End Enum

    Public Sub New()
        InitializeComponent()
        SetActiveSection(StudentSubjectsSection.CurrentLoad)
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedId As String = If(studentId, String.Empty).Trim()
        Dim normalizedName As String = If(studentName, String.Empty).Trim()

        SubjectsHeaderNameTextBlock.Text = If(String.IsNullOrWhiteSpace(normalizedName), "Student", normalizedName)
        SubjectsHeaderIdTextBlock.Text = normalizedId
        CurrentSubjectsTabView.SetStudentContext(normalizedId, normalizedName)
    End Sub

    Private Sub CurrentLoadSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentSubjectsSection.CurrentLoad)
    End Sub

    Private Sub CompletedSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentSubjectsSection.Completed)
    End Sub

    Private Sub PlannedSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentSubjectsSection.Planned)
    End Sub

    Private Sub SetActiveSection(section As StudentSubjectsSection)
        CurrentSubjectsTabView.Visibility = If(section = StudentSubjectsSection.CurrentLoad, Visibility.Visible, Visibility.Collapsed)
        CompletedSubjectsTabView.Visibility = If(section = StudentSubjectsSection.Completed, Visibility.Visible, Visibility.Collapsed)
        PlannedSubjectsTabView.Visibility = If(section = StudentSubjectsSection.Planned, Visibility.Visible, Visibility.Collapsed)

        ApplySectionButtonState(CurrentLoadSectionButton, section = StudentSubjectsSection.CurrentLoad)
        ApplySectionButtonState(CompletedSectionButton, section = StudentSubjectsSection.Completed)
        ApplySectionButtonState(PlannedSectionButton, section = StudentSubjectsSection.Planned)
    End Sub

    Private Sub ApplySectionButtonState(sectionButton As Button, isSelected As Boolean)
        sectionButton.Style = CType(FindResource(If(isSelected,
                                                    "DashboardProfileSegmentSelectedButtonStyle",
                                                    "DashboardProfileSegmentButtonStyle")), Style)
    End Sub
End Class
