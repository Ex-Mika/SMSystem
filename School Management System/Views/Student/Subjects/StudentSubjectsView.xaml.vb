Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class StudentSubjectsView
    Private Enum StudentSubjectsSection
        CurrentLoad
        Completed
        Planned
    End Enum

    Private ReadOnly _studentEnrollmentManagementService As New StudentEnrollmentManagementService()
    Private _currentStudentId As String = String.Empty

    Public Sub New()
        InitializeComponent()
        SetActiveSection(StudentSubjectsSection.CurrentLoad)
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedId As String = If(studentId, String.Empty).Trim()
        Dim normalizedName As String = If(studentName, String.Empty).Trim()

        _currentStudentId = normalizedId
        SubjectsHeaderNameTextBlock.Text = If(String.IsNullOrWhiteSpace(normalizedName), "Student", normalizedName)
        SubjectsHeaderIdTextBlock.Text = normalizedId
        CurrentSubjectsTabView.SetStudentContext(normalizedId, normalizedName)
        LoadEnrollmentSnapshot()
    End Sub

    Private Sub LoadEnrollmentSnapshot()
        If String.IsNullOrWhiteSpace(_currentStudentId) Then
            ApplyEnrollmentSnapshot(New StudentEnrollmentSnapshot(),
                                    "No student enrollment loaded.")
            Return
        End If

        Dim result = _studentEnrollmentManagementService.GetEnrollmentSnapshot(_currentStudentId)
        If result Is Nothing OrElse Not result.IsSuccess Then
            ApplyEnrollmentSnapshot(New StudentEnrollmentSnapshot(),
                                    If(result Is Nothing,
                                       "Unable to load current subjects.",
                                       result.Message))
            Return
        End If

        ApplyEnrollmentSnapshot(result.Data)
    End Sub

    Private Sub ApplyEnrollmentSnapshot(snapshot As StudentEnrollmentSnapshot,
                                        Optional statusMessage As String = "")
        Dim resolvedSnapshot As StudentEnrollmentSnapshot =
            If(snapshot, New StudentEnrollmentSnapshot())

        CurrentSubjectsCountTextBlock.Text = resolvedSnapshot.SelectedSubjectCount.ToString()
        CurrentSubjectsUnitsTextBlock.Text = resolvedSnapshot.SelectedTotalUnitsLabel
        CompletedSubjectsCountTextBlock.Text = "0"
        SubjectsEnrollmentTagTextBlock.Text = ResolveEnrollmentBadgeText(resolvedSnapshot)
        SubjectsSectionTagTextBlock.Text = ResolveSectionBadgeText(resolvedSnapshot)
        CurrentSubjectsTabView.SetEnrollmentSnapshot(resolvedSnapshot, statusMessage)
    End Sub

    Private Function ResolveEnrollmentBadgeText(snapshot As StudentEnrollmentSnapshot) As String
        If snapshot Is Nothing OrElse Not snapshot.HasSelectedSubjects Then
            Return "Enrollment Open"
        End If

        If snapshot.IsFullyEnrolled Then
            Return "Enrolled"
        End If

        If snapshot.HasRemainingSubjects Then
            Return "Load In Progress"
        End If

        Return "Current Load Saved"
    End Function

    Private Function ResolveSectionBadgeText(snapshot As StudentEnrollmentSnapshot) As String
        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "Section Pending"
        End If

        Return StudentScheduleHelper.BuildStudentSectionLabel(snapshot.Student,
                                                              "Section Pending")
    End Function

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
