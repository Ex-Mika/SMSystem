Imports System.Collections.Generic
Imports System.Globalization
Imports Microsoft.Win32
Imports System.Windows.Media
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class TeacherGradingView
    Private Enum TeacherGradingSection
        Pending
        Published
        History
    End Enum

    Private Enum ExportStatusTone
        Neutral
        Success
        [Error]
    End Enum

    Private Class StudentListExportSelectionOption
        Public Property DisplayLabel As String = String.Empty
        Public Property Target As TeacherStudentListExportTarget
        Public Property IsAllOption As Boolean

        Public Overrides Function ToString() As String
            Return If(DisplayLabel, String.Empty)
        End Function
    End Class

    Private ReadOnly _studentGradeManagementService As New StudentGradeManagementService()
    Private ReadOnly _teacherStudentListExportService As New TeacherStudentListExportService()

    Private _activeSection As TeacherGradingSection = TeacherGradingSection.Pending
    Private _allGradeRecords As New List(Of StudentSubjectGradeRecord)()
    Private _currentTeacherId As String = String.Empty
    Private _currentTeacherName As String = String.Empty
    Private _selectedGradeRecord As StudentSubjectGradeRecord
    Private _isRefreshingEditor As Boolean
    Private _studentListExportTargetLoadError As String = String.Empty

    Public Sub New()
        InitializeComponent()
        ClearEditor()
        RefreshExportActionState()
        SetActiveSection(TeacherGradingSection.Pending)
    End Sub

    Public Sub SetTeacherContext(teacherId As String, teacherName As String)
        _currentTeacherId = If(teacherId, String.Empty).Trim()
        _currentTeacherName = If(teacherName, String.Empty).Trim()
        GradingTeacherIdTextBlock.Text = ResolveTeacherContextLabel()
        LoadStudentListExportTargets()
        LoadGradeRecords()
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

    Private Sub GradingRosterDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ApplySelectedRecord(TryCast(GradingRosterDataGrid.SelectedItem,
                                    StudentSubjectGradeRecord))
    End Sub

    Private Sub ScoreInputTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        If _isRefreshingEditor Then
            Return
        End If

        UpdateGradePreview()
    End Sub

    Private Sub SaveGradeButton_Click(sender As Object, e As RoutedEventArgs)
        If _selectedGradeRecord Is Nothing Then
            MessageBox.Show("Select a student from the roster before saving a grade.",
                            "Grading",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        Dim quizScore As Decimal
        Dim projectScore As Decimal
        Dim midtermScore As Decimal
        Dim finalExamScore As Decimal
        Dim validationMessage As String = String.Empty

        If Not TryReadScoreInputs(quizScore,
                                  projectScore,
                                  midtermScore,
                                  finalExamScore,
                                  validationMessage) Then
            GradeEditorStatusTextBlock.Text = validationMessage
            MessageBox.Show(validationMessage,
                            "Grading",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        Dim request As New StudentGradeSaveRequest() With {
            .TeacherId = _currentTeacherId,
            .StudentNumber = _selectedGradeRecord.StudentNumber,
            .SubjectCode = _selectedGradeRecord.SubjectCode,
            .QuizScore = quizScore,
            .ProjectScore = projectScore,
            .MidtermScore = midtermScore,
            .FinalExamScore = finalExamScore
        }

        Dim result = _studentGradeManagementService.SaveTeacherGrade(request)
        If result Is Nothing OrElse Not result.IsSuccess Then
            Dim failureMessage As String =
                If(result Is Nothing,
                   "Unable to save the grade right now.",
                   result.Message)
            GradeEditorStatusTextBlock.Text = failureMessage
            MessageBox.Show(failureMessage,
                            "Grading",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        LoadGradeRecords(_selectedGradeRecord.StudentNumber,
                         _selectedGradeRecord.SubjectCode,
                         result.Message)
    End Sub

    Private Sub ResetGradeEditorButton_Click(sender As Object, e As RoutedEventArgs)
        ApplySelectedRecord(_selectedGradeRecord)
    End Sub

    Private Sub ExportStudentListTargetComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        RefreshExportActionState()
    End Sub

    Private Sub ExportStudentListPdfButton_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrWhiteSpace(_currentTeacherId) Then
            ShowStudentListExportStatus("Load the teacher grading roster before exporting the student list.",
                                        ExportStatusTone.Error)
            Return
        End If

        Dim selectedExportOption As StudentListExportSelectionOption =
            GetSelectedStudentListExportOption()

        Dim saveDialog As New SaveFileDialog() With {
            .Title = "Export Teacher Student List",
            .Filter = "PDF Document (*.pdf)|*.pdf",
            .DefaultExt = ".pdf",
            .AddExtension = True,
            .OverwritePrompt = True,
            .FileName = BuildDefaultStudentListFileName(selectedExportOption)
        }

        Dim accepted As Boolean? = saveDialog.ShowDialog()
        If accepted <> True Then
            ShowStudentListExportStatus("Student list export canceled.",
                                        ExportStatusTone.Neutral)
            Return
        End If

        Dim exportResult As ServiceResult(Of String) =
            _teacherStudentListExportService.ExportTeacherStudentList(_currentTeacherId,
                                                                     _currentTeacherName,
                                                                     saveDialog.FileName,
                                                                     ResolveSelectedStudentListExportTarget(selectedExportOption))
        If exportResult Is Nothing OrElse Not exportResult.IsSuccess Then
            Dim failureMessage As String =
                If(exportResult Is Nothing,
                   "Unable to export the teacher student list right now.",
                   exportResult.Message)
            ShowStudentListExportStatus(failureMessage,
                                        ExportStatusTone.Error)
            MessageBox.Show(failureMessage,
                            "Teacher Student List",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return
        End If

        Dim successMessage As String =
            BuildStudentListExportSuccessMessage(exportResult.Data,
                                                 selectedExportOption)
        ShowStudentListExportStatus(successMessage,
                                    ExportStatusTone.Success)
        MessageBox.Show(successMessage,
                        "Teacher Student List",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information)
    End Sub

    Private Sub LoadGradeRecords(Optional preferredStudentNumber As String = "",
                                 Optional preferredSubjectCode As String = "",
                                 Optional statusOverride As String = "")
        If String.IsNullOrWhiteSpace(_currentTeacherId) Then
            _allGradeRecords = New List(Of StudentSubjectGradeRecord)()
            UpdateSummaryCards()
            RefreshActiveSection("No teacher grading roster loaded.",
                                 preferredStudentNumber,
                                 preferredSubjectCode)
            Return
        End If

        Dim result = _studentGradeManagementService.GetTeacherGradeRecords(_currentTeacherId)
        If result Is Nothing OrElse Not result.IsSuccess Then
            _allGradeRecords = New List(Of StudentSubjectGradeRecord)()
            UpdateSummaryCards()
            RefreshActiveSection(If(result Is Nothing,
                                    "Unable to load the teacher grading roster.",
                                    result.Message),
                                 preferredStudentNumber,
                                 preferredSubjectCode)
            Return
        End If

        _allGradeRecords = If(result.Data, New List(Of StudentSubjectGradeRecord)())
        UpdateSummaryCards()
        RefreshActiveSection(statusOverride,
                             preferredStudentNumber,
                             preferredSubjectCode)
    End Sub

    Private Sub UpdateSummaryCards()
        Dim pendingCount As Integer = 0
        Dim publishedCount As Integer = 0
        Dim totalFinalGrade As Decimal = 0D

        For Each gradeRecord As StudentSubjectGradeRecord In _allGradeRecords
            If gradeRecord Is Nothing Then
                Continue For
            End If

            If gradeRecord.HasPublishedGrade Then
                publishedCount += 1

                If gradeRecord.FinalGrade.HasValue Then
                    totalFinalGrade += gradeRecord.FinalGrade.Value
                End If
            Else
                pendingCount += 1
            End If
        Next

        PendingChecksTextBlock.Text = pendingCount.ToString()
        PublishedGradesTextBlock.Text = publishedCount.ToString()

        If publishedCount = 0 Then
            ClassAverageTextBlock.Text = "0.00"
        Else
            ClassAverageTextBlock.Text =
                Decimal.Round(totalFinalGrade / publishedCount,
                              2,
                              MidpointRounding.AwayFromZero).
                    ToString("0.00", CultureInfo.InvariantCulture)
        End If
    End Sub

    Private Sub SetActiveSection(section As TeacherGradingSection)
        _activeSection = section

        RosterContentPanel.Visibility =
            If(section = TeacherGradingSection.History,
               Visibility.Collapsed,
               Visibility.Visible)
        HistoryPanel.Visibility =
            If(section = TeacherGradingSection.History,
               Visibility.Visible,
               Visibility.Collapsed)

        ApplySectionButtonState(PendingSectionButton,
                                section = TeacherGradingSection.Pending)
        ApplySectionButtonState(PublishedSectionButton,
                                section = TeacherGradingSection.Published)
        ApplySectionButtonState(HistorySectionButton,
                                section = TeacherGradingSection.History)

        RefreshActiveSection()
    End Sub

    Private Sub RefreshActiveSection(Optional statusOverride As String = "",
                                     Optional preferredStudentNumber As String = "",
                                     Optional preferredSubjectCode As String = "")
        If _activeSection = TeacherGradingSection.History Then
            BindHistoryRecords(statusOverride)
            Return
        End If

        Dim filteredRecords As List(Of StudentSubjectGradeRecord) =
            GetFilteredRosterRecords()
        Dim defaultStatusMessage As String = BuildRosterStatusMessage(filteredRecords.Count)
        Dim statusMessage As String = If(String.IsNullOrWhiteSpace(statusOverride),
                                         defaultStatusMessage,
                                         statusOverride)

        RosterPanelTitleTextBlock.Text = ResolveRosterTitle()
        RosterPanelSubtitleTextBlock.Text = ResolveRosterSubtitle()
        TeacherGradingStatusTextBlock.Text = statusMessage
        GradingRosterDataGrid.ItemsSource = filteredRecords

        Dim preferredRecord As StudentSubjectGradeRecord =
            FindRecord(filteredRecords,
                       preferredStudentNumber,
                       preferredSubjectCode)
        If preferredRecord Is Nothing AndAlso filteredRecords.Count > 0 Then
            preferredRecord = filteredRecords(0)
        End If

        _isRefreshingEditor = True
        GradingRosterDataGrid.SelectedItem = preferredRecord
        _isRefreshingEditor = False

        If preferredRecord Is Nothing Then
            ClearEditor()
            Return
        End If

        ApplySelectedRecord(preferredRecord)
    End Sub

    Private Function GetFilteredRosterRecords() As List(Of StudentSubjectGradeRecord)
        Dim filteredRecords As New List(Of StudentSubjectGradeRecord)()

        For Each gradeRecord As StudentSubjectGradeRecord In _allGradeRecords
            If gradeRecord Is Nothing Then
                Continue For
            End If

            If _activeSection = TeacherGradingSection.Pending AndAlso
               gradeRecord.HasPublishedGrade Then
                Continue For
            End If

            If _activeSection = TeacherGradingSection.Published AndAlso
               Not gradeRecord.HasPublishedGrade Then
                Continue For
            End If

            filteredRecords.Add(gradeRecord)
        Next

        Return filteredRecords
    End Function

    Private Sub BindHistoryRecords(Optional statusOverride As String = "")
        Dim publishedRecords As New List(Of StudentSubjectGradeRecord)()

        For Each gradeRecord As StudentSubjectGradeRecord In _allGradeRecords
            If gradeRecord IsNot Nothing AndAlso gradeRecord.HasPublishedGrade Then
                publishedRecords.Add(gradeRecord)
            End If
        Next

        GradingHistoryDataGrid.ItemsSource = publishedRecords

        If Not String.IsNullOrWhiteSpace(statusOverride) Then
            HistoryStatusTextBlock.Text = statusOverride
            Return
        End If

        If publishedRecords.Count = 0 Then
            HistoryStatusTextBlock.Text =
                "No grading history is available yet for your current subject roster."
            Return
        End If

        HistoryStatusTextBlock.Text =
            publishedRecords.Count.ToString() &
            " saved grade record(s) are already visible to students."
    End Sub

    Private Function ResolveTeacherContextLabel() As String
        If Not String.IsNullOrWhiteSpace(_currentTeacherName) AndAlso
           Not String.IsNullOrWhiteSpace(_currentTeacherId) Then
            Return _currentTeacherName & " | " & _currentTeacherId
        End If

        If Not String.IsNullOrWhiteSpace(_currentTeacherId) Then
            Return _currentTeacherId
        End If

        Return "No teacher loaded."
    End Function

    Private Function ResolveRosterTitle() As String
        If _activeSection = TeacherGradingSection.Published Then
            Return "Published Grades"
        End If

        Return "Students Ready for Grading"
    End Function

    Private Function ResolveRosterSubtitle() As String
        If _activeSection = TeacherGradingSection.Published Then
            Return "Review or update grades that are already posted to the student dashboard."
        End If

        Return "Current enrolled students in your handled subjects who still need posted grades."
    End Function

    Private Function BuildRosterStatusMessage(recordCount As Integer) As String
        If _activeSection = TeacherGradingSection.Published Then
            If recordCount = 0 Then
                Return "No published grades found for your handled subjects yet."
            End If

            Return recordCount.ToString() &
                " published grade record(s) are available for review."
        End If

        If recordCount = 0 Then
            Return "No pending grading items were found for your current subject roster."
        End If

        Return recordCount.ToString() &
            " pending student-subject roster item(s) are ready for grading."
    End Function

    Private Sub ApplySelectedRecord(record As StudentSubjectGradeRecord)
        _selectedGradeRecord = record

        If record Is Nothing Then
            ClearEditor()
            Return
        End If

        _isRefreshingEditor = True

        SelectedStudentTextBlock.Text =
            record.StudentName &
            " (" & record.StudentNumber & ")"
        SelectedSubjectTextBlock.Text = record.SubjectDisplayLabel
        SelectedSectionTextBlock.Text = record.SectionDisplayLabel
        QuizScoreTextBox.Text = If(record.QuizScore.HasValue,
                                   record.QuizScoreLabel,
                                   String.Empty)
        ProjectScoreTextBox.Text = If(record.ProjectScore.HasValue,
                                      record.ProjectScoreLabel,
                                      String.Empty)
        MidtermScoreTextBox.Text = If(record.MidtermScore.HasValue,
                                      record.MidtermScoreLabel,
                                      String.Empty)
        FinalExamScoreTextBox.Text = If(record.FinalExamScore.HasValue,
                                        record.FinalExamScoreLabel,
                                        String.Empty)

        _isRefreshingEditor = False

        SaveGradeButton.IsEnabled = True
        ResetGradeEditorButton.IsEnabled = True
        UpdateGradePreview()

        If record.HasPublishedGrade Then
            GradeEditorStatusTextBlock.Text =
                "Adjust the component scores below to update the published grade."
        Else
            GradeEditorStatusTextBlock.Text =
                "Enter complete scores from 0 to 100, then save to publish the grade."
        End If
    End Sub

    Private Sub ClearEditor()
        _selectedGradeRecord = Nothing
        _isRefreshingEditor = True

        SelectedStudentTextBlock.Text = "Select a student from the roster."
        SelectedSubjectTextBlock.Text = "--"
        SelectedSectionTextBlock.Text = "--"
        QuizScoreTextBox.Text = String.Empty
        ProjectScoreTextBox.Text = String.Empty
        MidtermScoreTextBox.Text = String.Empty
        FinalExamScoreTextBox.Text = String.Empty
        ComputedFinalGradeTextBlock.Text = "--"
        ComputedRemarksTextBlock.Text = "Pending"
        GradeEditorStatusTextBlock.Text = "Choose a student record to begin grading."

        _isRefreshingEditor = False

        SaveGradeButton.IsEnabled = False
        ResetGradeEditorButton.IsEnabled = False
    End Sub

    Private Sub RefreshExportActionState()
        If ExportStudentListPdfButton Is Nothing OrElse
           ExportStudentListTargetComboBox Is Nothing OrElse
           StudentListExportStatusTextBlock Is Nothing Then
            Return
        End If

        Dim hasTeacherContext As Boolean =
            Not String.IsNullOrWhiteSpace(_currentTeacherId)
        If Not hasTeacherContext Then
            ExportStudentListTargetComboBox.IsEnabled = False
            ExportStudentListPdfButton.IsEnabled = False
            ShowStudentListExportStatus("Load a teacher profile to export the student list PDF.",
                                        ExportStatusTone.Neutral)
            Return
        End If

        If Not String.IsNullOrWhiteSpace(_studentListExportTargetLoadError) Then
            ExportStudentListTargetComboBox.IsEnabled = False
            ExportStudentListPdfButton.IsEnabled = False
            ShowStudentListExportStatus(_studentListExportTargetLoadError,
                                        ExportStatusTone.Error)
            Return
        End If

        Dim hasExportTargets As Boolean = ExportStudentListTargetComboBox.Items.Count > 0
        ExportStudentListTargetComboBox.IsEnabled = hasExportTargets
        ExportStudentListPdfButton.IsEnabled = hasExportTargets

        If Not hasExportTargets Then
            ShowStudentListExportStatus("No handled subject sections are available for export.",
                                        ExportStatusTone.Neutral)
            Return
        End If

        Dim selectedExportOption As StudentListExportSelectionOption =
            GetSelectedStudentListExportOption()
        If selectedExportOption Is Nothing OrElse
           selectedExportOption.IsAllOption Then
            ShowStudentListExportStatus("Choose a handled subject and section, or leave All handled sections selected.",
                                        ExportStatusTone.Neutral)
            Return
        End If

        ShowStudentListExportStatus("Ready to export " & selectedExportOption.DisplayLabel & ".",
                                    ExportStatusTone.Neutral)
    End Sub

    Private Sub UpdateGradePreview()
        Dim quizScore As Decimal
        Dim projectScore As Decimal
        Dim midtermScore As Decimal
        Dim finalExamScore As Decimal
        Dim ignoredValidationMessage As String = String.Empty

        If Not TryReadScoreInputs(quizScore,
                                  projectScore,
                                  midtermScore,
                                  finalExamScore,
                                  ignoredValidationMessage,
                                  showValidation:=False) Then
            ComputedFinalGradeTextBlock.Text = "--"
            ComputedRemarksTextBlock.Text = "Pending"
            Return
        End If

        Dim finalGradeValue As Decimal =
            Decimal.Round((quizScore + projectScore + midtermScore + finalExamScore) / 4D,
                          2,
                          MidpointRounding.AwayFromZero)

        ComputedFinalGradeTextBlock.Text =
            finalGradeValue.ToString("0.00", CultureInfo.InvariantCulture)
        ComputedRemarksTextBlock.Text =
            If(finalGradeValue >= 75D, "Passed", "Failed")
    End Sub

    Private Function TryReadScoreInputs(ByRef quizScore As Decimal,
                                        ByRef projectScore As Decimal,
                                        ByRef midtermScore As Decimal,
                                        ByRef finalExamScore As Decimal,
                                        ByRef validationMessage As String,
                                        Optional showValidation As Boolean = True) As Boolean
        If Not TryReadScoreValue(QuizScoreTextBox.Text,
                                 "Quiz Score",
                                 quizScore,
                                 validationMessage) Then
            Return HandleScoreValidation(validationMessage, showValidation)
        End If

        If Not TryReadScoreValue(ProjectScoreTextBox.Text,
                                 "Project Score",
                                 projectScore,
                                 validationMessage) Then
            Return HandleScoreValidation(validationMessage, showValidation)
        End If

        If Not TryReadScoreValue(MidtermScoreTextBox.Text,
                                 "Midterm Score",
                                 midtermScore,
                                 validationMessage) Then
            Return HandleScoreValidation(validationMessage, showValidation)
        End If

        If Not TryReadScoreValue(FinalExamScoreTextBox.Text,
                                 "Final Exam Score",
                                 finalExamScore,
                                 validationMessage) Then
            Return HandleScoreValidation(validationMessage, showValidation)
        End If

        validationMessage = String.Empty
        Return True
    End Function

    Private Function HandleScoreValidation(validationMessage As String,
                                           showValidation As Boolean) As Boolean
        If Not showValidation Then
            validationMessage = String.Empty
        End If

        Return False
    End Function

    Private Function TryReadScoreValue(rawValue As String,
                                       fieldLabel As String,
                                       ByRef parsedValue As Decimal,
                                       ByRef validationMessage As String) As Boolean
        Dim normalizedValue As String = If(rawValue, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedValue) Then
            validationMessage = fieldLabel & " is required."
            Return False
        End If

        If Not Decimal.TryParse(normalizedValue,
                                NumberStyles.Number,
                                CultureInfo.InvariantCulture,
                                parsedValue) AndAlso
           Not Decimal.TryParse(normalizedValue,
                                NumberStyles.Number,
                                CultureInfo.CurrentCulture,
                                parsedValue) Then
            validationMessage = fieldLabel & " must be a valid number."
            Return False
        End If

        If parsedValue < 0D OrElse parsedValue > 100D Then
            validationMessage = fieldLabel & " must be between 0 and 100."
            Return False
        End If

        Return True
    End Function

    Private Function FindRecord(records As IEnumerable(Of StudentSubjectGradeRecord),
                                studentNumber As String,
                                subjectCode As String) As StudentSubjectGradeRecord
        Dim normalizedStudentNumber As String = If(studentNumber, String.Empty).Trim()
        Dim normalizedSubjectCode As String = If(subjectCode, String.Empty).Trim()
        If records Is Nothing OrElse
           normalizedStudentNumber = String.Empty OrElse
           normalizedSubjectCode = String.Empty Then
            Return Nothing
        End If

        For Each record As StudentSubjectGradeRecord In records
            If record Is Nothing Then
                Continue For
            End If

            If String.Equals(If(record.StudentNumber, String.Empty).Trim(),
                             normalizedStudentNumber,
                             StringComparison.OrdinalIgnoreCase) AndAlso
               String.Equals(If(record.SubjectCode, String.Empty).Trim(),
                             normalizedSubjectCode,
                             StringComparison.OrdinalIgnoreCase) Then
                Return record
            End If
        Next

        Return Nothing
    End Function

    Private Sub LoadStudentListExportTargets()
        If ExportStudentListTargetComboBox Is Nothing Then
            Return
        End If

        _studentListExportTargetLoadError = String.Empty
        ExportStudentListTargetComboBox.ItemsSource = Nothing
        ExportStudentListTargetComboBox.SelectedItem = Nothing

        If String.IsNullOrWhiteSpace(_currentTeacherId) Then
            RefreshExportActionState()
            Return
        End If

        Dim exportTargetsResult As ServiceResult(Of List(Of TeacherStudentListExportTarget)) =
            _teacherStudentListExportService.GetExportTargets(_currentTeacherId)
        If exportTargetsResult Is Nothing OrElse
           Not exportTargetsResult.IsSuccess Then
            _studentListExportTargetLoadError =
                If(exportTargetsResult Is Nothing,
                   "Unable to load the teacher student list choices.",
                   exportTargetsResult.Message)
            RefreshExportActionState()
            Return
        End If

        Dim selectionOptions As New List(Of StudentListExportSelectionOption) From {
            New StudentListExportSelectionOption() With {
                .DisplayLabel = "All handled sections",
                .IsAllOption = True
            }
        }

        For Each exportTarget As TeacherStudentListExportTarget In If(exportTargetsResult.Data,
                                                                      New List(Of TeacherStudentListExportTarget)())
            If exportTarget Is Nothing Then
                Continue For
            End If

            selectionOptions.Add(New StudentListExportSelectionOption() With {
                .DisplayLabel = exportTarget.DisplayLabel,
                .Target = exportTarget
            })
        Next

        ExportStudentListTargetComboBox.ItemsSource = selectionOptions
        ExportStudentListTargetComboBox.SelectedIndex =
            If(selectionOptions.Count > 0, 0, -1)
        RefreshExportActionState()
    End Sub

    Private Function GetSelectedStudentListExportOption() As StudentListExportSelectionOption
        Return TryCast(ExportStudentListTargetComboBox.SelectedItem,
                       StudentListExportSelectionOption)
    End Function

    Private Function ResolveSelectedStudentListExportTarget(selectedExportOption As StudentListExportSelectionOption) As TeacherStudentListExportTarget
        If selectedExportOption Is Nothing OrElse
           selectedExportOption.IsAllOption Then
            Return Nothing
        End If

        Return selectedExportOption.Target
    End Function

    Private Function BuildDefaultStudentListFileName(Optional selectedExportOption As StudentListExportSelectionOption = Nothing) As String
        Dim safeTeacherId As String = If(_currentTeacherId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(safeTeacherId) Then
            safeTeacherId = "teacher"
        End If

        safeTeacherId = BuildSafeFileNameSegment(safeTeacherId)

        Dim exportScopeLabel As String = String.Empty
        Dim selectedExportTarget As TeacherStudentListExportTarget =
            ResolveSelectedStudentListExportTarget(selectedExportOption)
        If selectedExportTarget IsNot Nothing Then
            Dim primarySubjectLabel As String =
                If(String.IsNullOrWhiteSpace(selectedExportTarget.SubjectCode),
                   selectedExportTarget.SubjectName,
                   selectedExportTarget.SubjectCode)
            Dim subjectToken As String =
                BuildSafeFileNameSegment(primarySubjectLabel)
            Dim sectionToken As String =
                BuildSafeFileNameSegment(selectedExportTarget.SectionLabel)

            If subjectToken <> String.Empty AndAlso sectionToken <> String.Empty Then
                exportScopeLabel = "-" & subjectToken & "-" & sectionToken
            ElseIf subjectToken <> String.Empty Then
                exportScopeLabel = "-" & subjectToken
            ElseIf sectionToken <> String.Empty Then
                exportScopeLabel = "-" & sectionToken
            End If
        End If

        Return "TeacherStudentList-" &
            safeTeacherId &
            exportScopeLabel &
            "-" &
            DateTime.Now.ToString("yyyyMMdd-HHmm") &
            ".pdf"
    End Function

    Private Function BuildSafeFileNameSegment(value As String) As String
        Dim normalizedValue As String = If(value, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedValue) Then
            Return String.Empty
        End If

        normalizedValue = normalizedValue.Replace(" "c, "-"c)

        For Each invalidCharacter As Char In IO.Path.GetInvalidFileNameChars()
            normalizedValue = normalizedValue.Replace(invalidCharacter, "_"c)
        Next

        Return normalizedValue
    End Function

    Private Function BuildStudentListExportSuccessMessage(exportedFilePath As String,
                                                          selectedExportOption As StudentListExportSelectionOption) As String
        Dim normalizedExportedFilePath As String = If(exportedFilePath, String.Empty).Trim()
        If selectedExportOption Is Nothing OrElse
           selectedExportOption.IsAllOption Then
            Return "Student list exported to " & normalizedExportedFilePath & "."
        End If

        Return "Student list for " &
            selectedExportOption.DisplayLabel &
            " exported to " &
            normalizedExportedFilePath &
            "."
    End Function

    Private Sub ShowStudentListExportStatus(message As String,
                                            tone As ExportStatusTone)
        If StudentListExportStatusTextBlock Is Nothing Then
            Return
        End If

        StudentListExportStatusTextBlock.Text = message

        Dim fallbackBrush As Brush =
            New SolidColorBrush(Color.FromRgb(&H7A, &H8C, &HA0))
        Dim accentBrush As Brush = fallbackBrush

        Select Case tone
            Case ExportStatusTone.Success
                accentBrush = TryCast(TryFindResource("DashboardSuccessBrush"), Brush)
                If accentBrush Is Nothing Then
                    accentBrush = New SolidColorBrush(Color.FromRgb(&H2A, &H7E, &H46))
                End If

            Case ExportStatusTone.Error
                accentBrush = TryCast(TryFindResource("DashboardDangerBrush"), Brush)
                If accentBrush Is Nothing Then
                    accentBrush = New SolidColorBrush(Color.FromRgb(&HC2, &H3D, &H3D))
                End If

            Case Else
                accentBrush = TryCast(TryFindResource("DashboardTextMutedBrush"), Brush)
                If accentBrush Is Nothing Then
                    accentBrush = fallbackBrush
                End If
        End Select

        StudentListExportStatusTextBlock.Foreground = accentBrush
    End Sub

    Private Sub ApplySectionButtonState(sectionButton As Button, isSelected As Boolean)
        sectionButton.Style = CType(FindResource(If(isSelected,
                                                    "DashboardProfileSegmentSelectedButtonStyle",
                                                    "DashboardProfileSegmentButtonStyle")), Style)
    End Sub
End Class
