Class TeacherGradingView
    Private Enum TeacherGradingSection
        Pending
        Published
        History
    End Enum

    Public Sub New()
        InitializeComponent()
        SetActiveSection(TeacherGradingSection.Pending)
    End Sub

    Public Sub SetTeacherContext(teacherId As String, teacherName As String)
        GradingTeacherIdTextBlock.Text = If(teacherId, String.Empty).Trim()
    End Sub

    Private Sub PendingSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(TeacherGradingSection.Pending)
    End Sub

    Private Sub PublishedSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(TeacherGradingSection.Published)
    End Sub

    Private Sub HistorySectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(TeacherGradingSection.History)
    End Sub

    Private Sub SetActiveSection(section As TeacherGradingSection)
        PendingTabView.Visibility = If(section = TeacherGradingSection.Pending, Visibility.Visible, Visibility.Collapsed)
        PublishedTabView.Visibility = If(section = TeacherGradingSection.Published, Visibility.Visible, Visibility.Collapsed)
        HistoryTabView.Visibility = If(section = TeacherGradingSection.History, Visibility.Visible, Visibility.Collapsed)

        ApplySectionButtonState(PendingSectionButton, section = TeacherGradingSection.Pending)
        ApplySectionButtonState(PublishedSectionButton, section = TeacherGradingSection.Published)
        ApplySectionButtonState(HistorySectionButton, section = TeacherGradingSection.History)
    End Sub

    Private Sub ApplySectionButtonState(sectionButton As Button, isSelected As Boolean)
        sectionButton.Style = CType(FindResource(If(isSelected,
                                                    "DashboardProfileSegmentSelectedButtonStyle",
                                                    "DashboardProfileSegmentButtonStyle")), Style)
    End Sub
End Class
