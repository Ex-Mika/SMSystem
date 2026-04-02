Class StudentGradesView
    Private Enum StudentGradesSection
        CurrentTerm
        Breakdown
        History
    End Enum

    Public Sub New()
        InitializeComponent()
        SetActiveSection(StudentGradesSection.CurrentTerm)
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedId As String = If(studentId, String.Empty).Trim()
        Dim normalizedName As String = If(studentName, String.Empty).Trim()
        GradesHeaderNameTextBlock.Text = If(String.IsNullOrWhiteSpace(normalizedName), "Student", normalizedName)
        GradesStudentIdTextBlock.Text = normalizedId
        CurrentGradesTabView.SetStudentContext(normalizedId, studentName)
        GradeBreakdownTabView.SetStudentContext(normalizedId, studentName)
        GradeHistoryTabView.SetStudentContext(normalizedId, studentName)
    End Sub

    Private Sub CurrentTermSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentGradesSection.CurrentTerm)
    End Sub

    Private Sub BreakdownSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentGradesSection.Breakdown)
    End Sub

    Private Sub HistorySectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentGradesSection.History)
    End Sub

    Private Sub SetActiveSection(section As StudentGradesSection)
        CurrentGradesTabView.Visibility = If(section = StudentGradesSection.CurrentTerm, Visibility.Visible, Visibility.Collapsed)
        GradeBreakdownTabView.Visibility = If(section = StudentGradesSection.Breakdown, Visibility.Visible, Visibility.Collapsed)
        GradeHistoryTabView.Visibility = If(section = StudentGradesSection.History, Visibility.Visible, Visibility.Collapsed)

        ApplySectionButtonState(CurrentTermSectionButton, section = StudentGradesSection.CurrentTerm)
        ApplySectionButtonState(BreakdownSectionButton, section = StudentGradesSection.Breakdown)
        ApplySectionButtonState(HistorySectionButton, section = StudentGradesSection.History)
    End Sub

    Private Sub ApplySectionButtonState(sectionButton As Button, isSelected As Boolean)
        sectionButton.Style = CType(FindResource(If(isSelected,
                                                    "DashboardProfileSegmentSelectedButtonStyle",
                                                    "DashboardProfileSegmentButtonStyle")), Style)
    End Sub
End Class
