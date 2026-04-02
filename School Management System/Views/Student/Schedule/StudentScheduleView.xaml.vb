Class StudentScheduleView
    Private Enum StudentScheduleSection
        AvailableSubjects
        SelectedLoad
        Submission
    End Enum

    Public Sub New()
        InitializeComponent()
        SetActiveSection(StudentScheduleSection.AvailableSubjects)
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedId As String = If(studentId, String.Empty).Trim()
        ScheduleStudentIdTextBlock.Text = normalizedId
        WeeklyScheduleTabView.SetStudentContext(normalizedId, studentName)
    End Sub

    Private Sub AvailableSubjectsSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentScheduleSection.AvailableSubjects)
    End Sub

    Private Sub SelectedLoadSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentScheduleSection.SelectedLoad)
    End Sub

    Private Sub SubmissionSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentScheduleSection.Submission)
    End Sub

    Private Sub SetActiveSection(section As StudentScheduleSection)
        WeeklyScheduleTabView.Visibility = If(section = StudentScheduleSection.AvailableSubjects, Visibility.Visible, Visibility.Collapsed)
        SelectedLoadTabView.Visibility = If(section = StudentScheduleSection.SelectedLoad, Visibility.Visible, Visibility.Collapsed)
        SubmissionTabView.Visibility = If(section = StudentScheduleSection.Submission, Visibility.Visible, Visibility.Collapsed)

        ApplySectionButtonState(AvailableSubjectsSectionButton, section = StudentScheduleSection.AvailableSubjects)
        ApplySectionButtonState(SelectedLoadSectionButton, section = StudentScheduleSection.SelectedLoad)
        ApplySectionButtonState(SubmissionSectionButton, section = StudentScheduleSection.Submission)
    End Sub

    Private Sub ApplySectionButtonState(sectionButton As Button, isSelected As Boolean)
        sectionButton.Style = CType(FindResource(If(isSelected,
                                                    "DashboardProfileSegmentSelectedButtonStyle",
                                                    "DashboardProfileSegmentButtonStyle")), Style)
    End Sub
End Class
