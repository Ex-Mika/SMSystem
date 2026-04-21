Imports System.Data
Imports Microsoft.Win32
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class StudentScheduleView
    Private Enum StudentScheduleSection
        AvailableSubjects
        SelectedLoad
        Submission
    End Enum

    Public Event TimetableChanged As EventHandler

    Private ReadOnly _studentEnrollmentManagementService As New StudentEnrollmentManagementService()
    Private ReadOnly _enrollmentCertificateExportService As New EnrollmentCertificateExportService()
    Private _currentStudentId As String = String.Empty
    Private _currentEnrollmentSnapshot As New StudentEnrollmentSnapshot()

    Public Sub New()
        InitializeComponent()
        AddHandler SelectedLoadTabView.EnrollRequested, AddressOf SelectedLoadTabView_EnrollRequested
        AddHandler SelectedLoadTabView.RemoveRequested, AddressOf SelectedLoadTabView_RemoveRequested
        AddHandler SubmissionTabView.SectionSelectionRequested, AddressOf SubmissionTabView_SectionSelectionRequested
        AddHandler SubmissionTabView.CertificateExportRequested, AddressOf SubmissionTabView_CertificateExportRequested
        ApplyEnrollmentSnapshot(New StudentEnrollmentSnapshot())
        SetActiveSection(StudentScheduleSection.AvailableSubjects)
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedId As String = If(studentId, String.Empty).Trim()
        _currentStudentId = normalizedId
        ScheduleStudentIdTextBlock.Text = normalizedId
        WeeklyScheduleTabView.SetStudentContext(normalizedId, studentName)
        LoadEnrollmentSnapshot()
    End Sub

    Public Function GetCurrentTimetableSnapshot() As DataTable
        Return WeeklyScheduleTabView.GetTimetableSnapshot()
    End Function

    Private Sub LoadEnrollmentSnapshot(Optional statusMessage As String = "")
        If String.IsNullOrWhiteSpace(_currentStudentId) Then
            ApplyEnrollmentSnapshot(New StudentEnrollmentSnapshot(),
                                    "No student enrollment loaded.")
            Return
        End If

        Dim result = _studentEnrollmentManagementService.GetEnrollmentSnapshot(_currentStudentId)
        If Not result.IsSuccess Then
            ApplyEnrollmentSnapshot(New StudentEnrollmentSnapshot(), result.Message)
            Return
        End If

        ApplyEnrollmentSnapshot(result.Data, statusMessage)
    End Sub

    Private Sub ApplyEnrollmentSnapshot(snapshot As StudentEnrollmentSnapshot,
                                        Optional statusMessage As String = "")
        _currentEnrollmentSnapshot = If(snapshot, New StudentEnrollmentSnapshot())

        WeeklyScheduleTabView.SetEnrollmentSnapshot(_currentEnrollmentSnapshot)
        SelectedLoadTabView.SetEnrollmentSnapshot(_currentEnrollmentSnapshot, statusMessage)
        SubmissionTabView.SetEnrollmentSnapshot(_currentEnrollmentSnapshot, statusMessage)

        ScheduleSelectedSubjectsTextBlock.Text =
            _currentEnrollmentSnapshot.SelectedSubjectCount.ToString()
        ScheduleTotalUnitsTextBlock.Text = _currentEnrollmentSnapshot.SelectedTotalUnitsLabel
        ScheduleEnrollmentStatusTextBlock.Text =
            ResolveEnrollmentStatusLabel(_currentEnrollmentSnapshot)
        RaiseEvent TimetableChanged(Me, EventArgs.Empty)
    End Sub

    Private Function ResolveEnrollmentStatusLabel(snapshot As StudentEnrollmentSnapshot) As String
        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "Open"
        End If

        If Not snapshot.HasAssignedYearLevel Then
            Return "Profile Needed"
        End If

        If Not snapshot.HasSelectedSubjects Then
            Return "Open"
        End If

        If Not snapshot.HasAssignedSection Then
            Return "Draft"
        End If

        If snapshot.HasRemainingSubjects Then
            Return "In Progress"
        End If

        Return "Enrolled"
    End Function

    Private Sub SelectedLoadTabView_EnrollRequested(sender As Object,
                                                    e As StudentExamScheduleTab.StudentSubjectActionEventArgs)
        ProcessEnrollmentChange(e, shouldEnroll:=True)
    End Sub

    Private Sub SelectedLoadTabView_RemoveRequested(sender As Object,
                                                    e As StudentExamScheduleTab.StudentSubjectActionEventArgs)
        ProcessEnrollmentChange(e, shouldEnroll:=False)
    End Sub

    Private Sub ProcessEnrollmentChange(e As StudentExamScheduleTab.StudentSubjectActionEventArgs,
                                        shouldEnroll As Boolean)
        If e Is Nothing OrElse
           String.IsNullOrWhiteSpace(e.SubjectCode) OrElse
           String.IsNullOrWhiteSpace(_currentStudentId) Then
            Return
        End If

        Dim result =
            If(shouldEnroll,
               _studentEnrollmentManagementService.EnrollSubject(_currentStudentId, e.SubjectCode),
               _studentEnrollmentManagementService.RemoveSubject(_currentStudentId, e.SubjectCode))

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Enrollment",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        LoadEnrollmentSnapshot(result.Message)
    End Sub

    Private Sub SubmissionTabView_SectionSelectionRequested(sender As Object,
                                                            e As StudentConsultationScheduleTab.StudentSectionSelectionEventArgs)
        If e Is Nothing OrElse
           String.IsNullOrWhiteSpace(e.SectionName) OrElse
           String.IsNullOrWhiteSpace(_currentStudentId) Then
            Return
        End If

        Dim result =
            _studentEnrollmentManagementService.UpdateStudentSection(_currentStudentId,
                                                                    e.SectionName)
        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Enrollment",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        LoadEnrollmentSnapshot(result.Message)
    End Sub

    Private Sub SubmissionTabView_CertificateExportRequested(sender As Object,
                                                             e As EventArgs)
        If _currentEnrollmentSnapshot Is Nothing OrElse
           _currentEnrollmentSnapshot.Student Is Nothing Then
            SubmissionTabView.ShowCertificateExportStatus(
                "No enrollment record is loaded for certificate export.",
                isError:=True)
            Return
        End If

        Dim student As StudentRecord = _currentEnrollmentSnapshot.Student
        Dim saveDialog As New SaveFileDialog() With {
            .Title = "Export Certificate of Enrollment",
            .Filter = "PDF Document (*.pdf)|*.pdf",
            .DefaultExt = ".pdf",
            .AddExtension = True,
            .OverwritePrompt = True,
            .FileName = BuildDefaultCertificateFileName(student)
        }

        Dim accepted As Boolean? = saveDialog.ShowDialog()
        If accepted <> True Then
            SubmissionTabView.ShowCertificateExportStatus("Certificate export canceled.",
                                                          isError:=False)
            Return
        End If

        Dim exportResult As Backend.Common.ServiceResult(Of String) =
            _enrollmentCertificateExportService.ExportCertificate(
                _currentEnrollmentSnapshot,
                saveDialog.FileName)
        If Not exportResult.IsSuccess Then
            SubmissionTabView.ShowCertificateExportStatus(exportResult.Message,
                                                          isError:=True)
            MessageBox.Show(exportResult.Message,
                            "Certificate of Enrollment",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return
        End If

        Dim successMessage As String =
            "Certificate exported to " & exportResult.Data & "."
        SubmissionTabView.ShowCertificateExportStatus(successMessage,
                                                      isError:=False)
        MessageBox.Show(successMessage,
                        "Certificate of Enrollment",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information)
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

    Private Function BuildDefaultCertificateFileName(student As StudentRecord) As String
        Dim studentId As String =
            If(student Is Nothing, String.Empty, student.StudentNumber)
        Dim safeStudentId As String = If(studentId, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(safeStudentId) Then
            safeStudentId = "student"
        End If

        For Each invalidCharacter As Char In IO.Path.GetInvalidFileNameChars()
            safeStudentId = safeStudentId.Replace(invalidCharacter, "_"c)
        Next

        Return "CertificateOfEnrollment-" &
            safeStudentId &
            "-" &
            DateTime.Now.ToString("yyyyMMdd-HHmm") &
            ".pdf"
    End Function
End Class
