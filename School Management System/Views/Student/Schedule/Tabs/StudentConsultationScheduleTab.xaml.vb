Imports System.Collections.Generic
Imports System.Windows.Media
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models

Class StudentConsultationScheduleTab
    Public Class StudentSectionSelectionEventArgs
        Inherits EventArgs

        Public Sub New(sectionName As String)
            Me.SectionName = If(sectionName, String.Empty).Trim()
        End Sub

        Public ReadOnly Property SectionName As String
    End Class

    Public Event SectionSelectionRequested As EventHandler(Of StudentSectionSelectionEventArgs)
    Public Event CertificateExportRequested As EventHandler

    Public Sub New()
        InitializeComponent()
        SetEnrollmentSnapshot(Nothing)
    End Sub

    Public Sub SetEnrollmentSnapshot(snapshot As StudentEnrollmentSnapshot,
                                     Optional statusMessage As String = "")
        Dim resolvedSnapshot As StudentEnrollmentSnapshot =
            If(snapshot, New StudentEnrollmentSnapshot())
        Dim hasYearLevel As Boolean = resolvedSnapshot.HasAssignedYearLevel
        Dim hasSection As Boolean = resolvedSnapshot.HasAssignedSection
        Dim hasFinalizedLoad As Boolean =
            resolvedSnapshot.HasSelectedSubjects AndAlso
            Not resolvedSnapshot.HasRemainingSubjects

        Dim completedRequirements As Integer = 0
        If hasYearLevel Then
            completedRequirements += 1
        End If

        If hasSection Then
            completedRequirements += 1
        End If

        If hasFinalizedLoad Then
            completedRequirements += 1
        End If

        Dim percentageComplete As Integer =
            CInt(Math.Round((completedRequirements / 3D) * 100D,
                            MidpointRounding.AwayFromZero))

        RequirementsCompleteValueTextBlock.Text = percentageComplete.ToString() & "%"
        AdviserApprovalValueTextBlock.Text =
            If(resolvedSnapshot.IsFullyEnrolled, "Ready", "Pending")

        SetRequirementState(YearLevelStatusTextBlock,
                            YearLevelNoteTextBlock,
                            hasYearLevel,
                            If(hasYearLevel,
                               resolvedSnapshot.Student.YearLevelLabel,
                               "Ask the registrar to assign your year level."))
        SetRequirementState(SectionStatusTextBlock,
                            SectionNoteTextBlock,
                            hasSection,
                            If(hasSection,
                               BuildSectionLabel(resolvedSnapshot.Student.SectionName,
                                                 resolvedSnapshot.Student),
                               "Choose a section from the schedules below."))
        SetRequirementState(SelectedLoadStatusTextBlock,
                            SelectedLoadNoteTextBlock,
                            hasFinalizedLoad,
                            BuildSelectedLoadRequirementNote(resolvedSnapshot))
        ApplySectionSelectionState(resolvedSnapshot)

        SubmissionStatusTextBlock.Text =
            ResolveSubmissionStatusMessage(resolvedSnapshot,
                                           completedRequirements,
                                           statusMessage)
        SubmissionHelperTextBlock.Text =
            "Subjects are saved immediately when you enroll or remove them. " &
            "This screen now tracks readiness only."
        ExportCertificateButton.IsEnabled =
            resolvedSnapshot IsNot Nothing AndAlso resolvedSnapshot.IsFullyEnrolled
        CertificateExportHintTextBlock.Text =
            BuildCertificateExportHint(resolvedSnapshot)
        ShowCertificateExportStatus(BuildCertificateExportStatus(resolvedSnapshot),
                                    isError:=False)
    End Sub

    Private Sub ApplySectionSelectionState(snapshot As StudentEnrollmentSnapshot)
        Dim resolvedSnapshot As StudentEnrollmentSnapshot =
            If(snapshot, New StudentEnrollmentSnapshot())
        Dim sectionOptions As List(Of StudentSectionOption) =
            If(resolvedSnapshot.AvailableSections, New List(Of StudentSectionOption)())
        Dim hasSectionOptions As Boolean = sectionOptions.Count > 0

        SectionSelectionComboBox.ItemsSource = sectionOptions
        SectionSelectionComboBox.IsEnabled = hasSectionOptions
        SaveSectionButton.IsEnabled = hasSectionOptions

        If hasSectionOptions Then
            SelectCurrentSectionOption(sectionOptions, resolvedSnapshot.Student)

            If SectionSelectionComboBox.SelectedItem Is Nothing Then
                SectionSelectionComboBox.SelectedIndex = 0
            End If
        Else
            SectionSelectionComboBox.SelectedItem = Nothing
        End If

        SectionSelectionStatusTextBlock.Text =
            BuildSectionSelectionStatus(resolvedSnapshot, hasSectionOptions)
    End Sub

    Private Function ResolveSubmissionStatusMessage(snapshot As StudentEnrollmentSnapshot,
                                                    completedRequirements As Integer,
                                                    statusMessage As String) As String
        Dim normalizedStatusMessage As String = If(statusMessage, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(normalizedStatusMessage) Then
            Return normalizedStatusMessage
        End If

        If completedRequirements = 3 Then
            Return "Your enrollment is complete. Your timetable is now ready to load."
        End If

        If snapshot IsNot Nothing AndAlso
           snapshot.HasSelectedSubjects AndAlso
           snapshot.HasRemainingSubjects Then
            Return "Add the remaining subjects in Enrollment before your timetable is unlocked."
        End If

        If snapshot IsNot Nothing AndAlso
           Not String.IsNullOrWhiteSpace(snapshot.NoticeMessage) Then
            Return snapshot.NoticeMessage
        End If

        Return "Complete the checklist below to prepare your enrollment."
    End Function

    Private Function BuildSelectedLoadRequirementNote(snapshot As StudentEnrollmentSnapshot) As String
        If snapshot Is Nothing OrElse Not snapshot.HasSelectedSubjects Then
            Return "Add all subjects assigned to your year level."
        End If

        Dim loadSummary As String =
            snapshot.SelectedSubjectCount.ToString() &
            " subjects | " &
            snapshot.SelectedTotalUnitsLabel &
            " units"
        If Not snapshot.HasRemainingSubjects Then
            Return loadSummary
        End If

        Dim remainingLabel As String = " subject remaining"
        If snapshot.AvailableSubjectCount <> 1 Then
            remainingLabel = " subjects remaining"
        End If

        Return loadSummary &
            " | " &
            snapshot.AvailableSubjectCount.ToString() &
            remainingLabel
    End Function

    Private Sub SetRequirementState(statusTextBlock As TextBlock,
                                    noteTextBlock As TextBlock,
                                    isComplete As Boolean,
                                    noteText As String)
        If statusTextBlock Is Nothing OrElse noteTextBlock Is Nothing Then
            Return
        End If

        statusTextBlock.Text = If(isComplete, "Complete", "Needed")
        statusTextBlock.Foreground =
            CType(FindResource(If(isComplete,
                                  "DashboardPrimaryBrush",
                                  "DashboardTextMutedBrush")), Brush)
        noteTextBlock.Text = If(noteText, String.Empty).Trim()
    End Sub

    Private Sub SaveSectionButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedOption As StudentSectionOption =
            TryCast(SectionSelectionComboBox.SelectedItem, StudentSectionOption)
        If selectedOption Is Nothing OrElse
           String.IsNullOrWhiteSpace(selectedOption.SectionValue) Then
            Return
        End If

        RaiseEvent SectionSelectionRequested(Me,
                                             New StudentSectionSelectionEventArgs(
                                                 selectedOption.SectionValue))
    End Sub

    Private Sub ExportCertificateButton_Click(sender As Object, e As RoutedEventArgs)
        If Not ExportCertificateButton.IsEnabled Then
            Return
        End If

        RaiseEvent CertificateExportRequested(Me, EventArgs.Empty)
    End Sub

    Public Sub ShowCertificateExportStatus(statusMessage As String,
                                           isError As Boolean)
        CertificateExportStatusTextBlock.Text =
            If(statusMessage, String.Empty).Trim()
        CertificateExportStatusTextBlock.Foreground =
            CType(FindResource(If(isError,
                                  "DashboardDangerBrush",
                                  "DashboardTextMutedBrush")), Brush)
    End Sub

    Private Sub SelectCurrentSectionOption(options As IEnumerable(Of StudentSectionOption),
                                           student As StudentRecord)
        SectionSelectionComboBox.SelectedItem = Nothing

        Dim currentToken As String = BuildSectionComparisonToken(If(student Is Nothing,
                                                                    String.Empty,
                                                                    student.SectionName),
                                                                 student)
        If options Is Nothing OrElse String.IsNullOrWhiteSpace(currentToken) Then
            Return
        End If

        For Each sectionOption As StudentSectionOption In options
            If sectionOption Is Nothing Then
                Continue For
            End If

            If String.Equals(If(sectionOption.ComparisonToken, String.Empty).Trim(),
                             currentToken,
                             StringComparison.OrdinalIgnoreCase) Then
                SectionSelectionComboBox.SelectedItem = sectionOption
                Return
            End If
        Next
    End Sub

    Private Function BuildSectionSelectionStatus(snapshot As StudentEnrollmentSnapshot,
                                                 hasSectionOptions As Boolean) As String
        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "Load your student record to choose a section."
        End If

        Dim currentSectionLabel As String =
            BuildSectionLabel(snapshot.Student.SectionName, snapshot.Student)
        If Not String.IsNullOrWhiteSpace(currentSectionLabel) Then
            If hasSectionOptions Then
                Return "Current section: " & currentSectionLabel &
                    ". You can switch to any section listed above."
            End If

            Return "Current section: " & currentSectionLabel & "."
        End If

        If hasSectionOptions Then
            Return "Choose a section from the schedules already assigned to teachers."
        End If

        If snapshot.SelectedSubjectCount > 0 Then
            Return "No scheduled sections are available for your selected subjects yet."
        End If

        If snapshot.AvailableSubjectCount > 0 Then
            Return "No scheduled sections are available for your year level yet."
        End If

        Return "Sections appear here once matching teacher schedules are available."
    End Function

    Private Function BuildCertificateExportHint(snapshot As StudentEnrollmentSnapshot) As String
        If snapshot IsNot Nothing AndAlso snapshot.IsFullyEnrolled Then
            Return "Export a formatted PDF copy of your certificate with your " &
                "current enrolled schedule."
        End If

        Return "Finish enrollment, choose a section, and complete your load " &
            "to unlock PDF export."
    End Function

    Private Function BuildCertificateExportStatus(snapshot As StudentEnrollmentSnapshot) As String
        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "Load your student enrollment record to prepare the certificate."
        End If

        If Not snapshot.HasAssignedYearLevel Then
            Return "Year level assignment is required before the certificate can be generated."
        End If

        If Not snapshot.HasSelectedSubjects Then
            Return "Select your subjects first before exporting the certificate."
        End If

        If snapshot.HasRemainingSubjects Then
            Return "Complete the remaining subject selections before exporting the certificate."
        End If

        If Not snapshot.HasAssignedSection Then
            Return "Choose a section before exporting the certificate."
        End If

        Return "Your enrollment is complete. The certificate is ready to export as PDF."
    End Function

    Private Function BuildSectionLabel(sectionValue As String,
                                       student As StudentRecord) As String
        Return StudentScheduleHelper.BuildSectionValue(sectionValue,
                                                       ResolveStudentYearToken(student),
                                                       String.Empty)
    End Function

    Private Function BuildSectionComparisonToken(sectionValue As String,
                                                 student As StudentRecord) As String
        Dim normalizedSection As String = NormalizeSectionToken(sectionValue)
        If String.IsNullOrWhiteSpace(normalizedSection) Then
            Return String.Empty
        End If

        Dim yearToken As String = ResolveStudentYearToken(student)
        If String.IsNullOrWhiteSpace(yearToken) OrElse
           normalizedSection.StartsWith(yearToken,
                                        StringComparison.OrdinalIgnoreCase) Then
            Return normalizedSection.ToUpperInvariant()
        End If

        Return (yearToken & normalizedSection).ToUpperInvariant()
    End Function

    Private Function ResolveStudentYearToken(student As StudentRecord) As String
        If student Is Nothing OrElse Not student.YearLevel.HasValue Then
            Return String.Empty
        End If

        Return student.YearLevel.Value.ToString()
    End Function

    Private Function NormalizeSectionToken(sectionValue As String) As String
        Dim normalizedSection As String = If(sectionValue, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedSection) Then
            Return String.Empty
        End If

        If normalizedSection.StartsWith("Section:",
                                        StringComparison.OrdinalIgnoreCase) Then
            normalizedSection = normalizedSection.Substring("Section:".Length).Trim()
        ElseIf normalizedSection.StartsWith("Section ",
                                            StringComparison.OrdinalIgnoreCase) Then
            normalizedSection = normalizedSection.Substring("Section ".Length).Trim()
        End If

        Return normalizedSection.Replace(" ", String.Empty)
    End Function
End Class
