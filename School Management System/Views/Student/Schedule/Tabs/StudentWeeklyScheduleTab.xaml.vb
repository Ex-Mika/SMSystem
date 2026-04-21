Imports System.Collections.Generic
Imports System.Data
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class StudentWeeklyScheduleTab
    Private ReadOnly _studentManagementService As New StudentManagementService()
    Private ReadOnly _subjectManagementService As New SubjectManagementService()
    Private ReadOnly _teacherScheduleManagementService As New TeacherScheduleManagementService()

    Private _currentStudentId As String = String.Empty
    Private _currentEnrollmentSnapshot As New StudentEnrollmentSnapshot()
    Private _currentStudentRecord As StudentRecord
    Private _subjectRecords As New List(Of SubjectRecord)()
    Private _timetableTable As DataTable

    Public Sub New()
        InitializeComponent()
        ApplyTimetable(StudentTimetablePresenter.CreateEmptyTable())
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedStudentId As String = If(studentId, String.Empty).Trim()
        _currentStudentId = normalizedStudentId
        WeeklyScheduleStudentIdTextBlock.Text = normalizedStudentId

        If String.IsNullOrWhiteSpace(normalizedStudentId) Then
            ApplyEmptyState("No student schedule loaded.")
        End If
    End Sub

    Public Sub SetEnrollmentSnapshot(snapshot As StudentEnrollmentSnapshot)
        _currentEnrollmentSnapshot = If(snapshot, New StudentEnrollmentSnapshot())

        Dim snapshotStudent As StudentRecord = _currentEnrollmentSnapshot.Student
        If snapshotStudent IsNot Nothing AndAlso
           Not String.IsNullOrWhiteSpace(snapshotStudent.StudentNumber) Then
            _currentStudentId = snapshotStudent.StudentNumber.Trim()
            WeeklyScheduleStudentIdTextBlock.Text = _currentStudentId
        End If

        LoadStudentTimetable(_currentStudentId)
    End Sub

    Public Function GetTimetableSnapshot() As DataTable
        If _timetableTable Is Nothing Then
            Return StudentTimetablePresenter.CreateEmptyTable()
        End If

        Return _timetableTable.Copy()
    End Function

    Private Sub LoadStudentTimetable(studentId As String)
        _currentStudentRecord = Nothing
        _subjectRecords = New List(Of SubjectRecord)()

        If String.IsNullOrWhiteSpace(studentId) Then
            ApplyEmptyState("No student schedule loaded.")
            Return
        End If

        _currentStudentRecord = ResolveStudentRecord(studentId)
        If _currentStudentRecord Is Nothing Then
            Return
        End If

        If Not _currentEnrollmentSnapshot.IsFullyEnrolled Then
            ApplyEmptyState(BuildEnrollmentPendingMessage(_currentEnrollmentSnapshot))
            Return
        End If

        Dim subjectResult = _subjectManagementService.GetSubjects()
        If Not subjectResult.IsSuccess Then
            ApplyEmptyState(subjectResult.Message)
            Return
        End If

        _subjectRecords = If(subjectResult.Data, New List(Of SubjectRecord)())

        Dim scheduleResult = _teacherScheduleManagementService.GetSchedules()
        If Not scheduleResult.IsSuccess Then
            ApplyEmptyState(scheduleResult.Message)
            Return
        End If

        Dim matchingSchedules As List(Of TeacherScheduleRecord) = FilterSchedulesForStudent(scheduleResult.Data)
        Dim timetableTable As DataTable = StudentTimetablePresenter.BuildTable(matchingSchedules, _subjectRecords)

        ApplyTimetable(timetableTable)
        UpdateScheduleSummary(matchingSchedules)
    End Sub

    Private Function ResolveStudentRecord(studentId As String) As StudentRecord
        Dim snapshotStudent As StudentRecord =
            If(_currentEnrollmentSnapshot, New StudentEnrollmentSnapshot()).Student
        If snapshotStudent IsNot Nothing AndAlso
           String.Equals(If(snapshotStudent.StudentNumber, String.Empty).Trim(),
                         studentId,
                         StringComparison.OrdinalIgnoreCase) Then
            Return snapshotStudent
        End If

        Dim studentResult = _studentManagementService.GetStudentByStudentNumber(studentId)
        If Not studentResult.IsSuccess Then
            ApplyEmptyState(studentResult.Message)
            Return Nothing
        End If

        If studentResult.Data Is Nothing Then
            ApplyEmptyState("No student record was found for this account.")
            Return Nothing
        End If

        Return studentResult.Data
    End Function

    Private Function BuildEnrollmentPendingMessage(snapshot As StudentEnrollmentSnapshot) As String
        Dim resolvedSnapshot As StudentEnrollmentSnapshot =
            If(snapshot, New StudentEnrollmentSnapshot())
        If resolvedSnapshot.Student Is Nothing Then
            Return "Complete your enrollment to load the timetable."
        End If

        If Not resolvedSnapshot.HasAssignedYearLevel Then
            Return "Your timetable will appear after the admin assigns your year level."
        End If

        If Not resolvedSnapshot.HasSelectedSubjects Then
            If Not String.IsNullOrWhiteSpace(resolvedSnapshot.NoticeMessage) Then
                Return resolvedSnapshot.NoticeMessage
            End If

            Return "Add your subjects in Enrollment before the timetable can load."
        End If

        If resolvedSnapshot.HasRemainingSubjects Then
            If resolvedSnapshot.AvailableSubjectCount = 1 Then
                Return "Finish enrolling in the remaining subject before the timetable loads."
            End If

            Return "Finish enrolling in the remaining " &
                resolvedSnapshot.AvailableSubjectCount.ToString() &
                " subjects before the timetable loads."
        End If

        If Not resolvedSnapshot.HasAssignedSection Then
            Return "Choose a section in Submission to load your class schedule."
        End If

        Return "Complete your enrollment to load the timetable."
    End Function

    Private Function FilterSchedulesForStudent(schedules As IEnumerable(Of TeacherScheduleRecord)) As List(Of TeacherScheduleRecord)
        Dim matches As New List(Of TeacherScheduleRecord)()
        If schedules Is Nothing OrElse _currentStudentRecord Is Nothing Then
            Return matches
        End If

        For Each schedule As TeacherScheduleRecord In schedules
            If StudentTimetablePresenter.MatchesStudentSchedule(schedule, _currentStudentRecord, _subjectRecords) Then
                matches.Add(schedule)
            End If
        Next

        Return matches
    End Function

    Private Sub ApplyTimetable(table As DataTable)
        _timetableTable = table
        StudentTimetablePresenter.ConfigureDataGrid(StudentTimetableDataGrid, _timetableTable)
        StudentTimetableDataGrid.ItemsSource = _timetableTable.DefaultView
    End Sub

    Private Sub ApplyEmptyState(statusMessage As String)
        ApplyTimetable(StudentTimetablePresenter.CreateEmptyTable())
        WeeklyScheduleSlotCountTextBlock.Text = "0"
        WeeklyScheduleDayCountTextBlock.Text = "0"
        WeeklyScheduleSectionTextBlock.Text =
            StudentScheduleHelper.BuildStudentSectionValue(_currentStudentRecord)
        WeeklyScheduleStatusTextBlock.Text = If(String.IsNullOrWhiteSpace(statusMessage),
                                               "No student schedule loaded.",
                                               statusMessage)
    End Sub

    Private Sub UpdateScheduleSummary(matchingSchedules As IEnumerable(Of TeacherScheduleRecord))
        Dim assignedSlots As Integer = 0
        Dim activeDays As New List(Of String)()

        If matchingSchedules IsNot Nothing Then
            For Each schedule As TeacherScheduleRecord In matchingSchedules
                If schedule Is Nothing Then
                    Continue For
                End If

                assignedSlots += 1

                Dim normalizedDay As String = StudentTimetablePresenter.NormalizeDayLabel(schedule.Day)
                If String.IsNullOrWhiteSpace(normalizedDay) Then
                    Continue For
                End If

                Dim alreadyCounted As Boolean = False
                For Each activeDay As String In activeDays
                    If String.Equals(activeDay, normalizedDay, StringComparison.OrdinalIgnoreCase) Then
                        alreadyCounted = True
                        Exit For
                    End If
                Next

                If Not alreadyCounted Then
                    activeDays.Add(normalizedDay)
                End If
            Next
        End If

        Dim sectionLabel As String =
            StudentScheduleHelper.BuildStudentSectionValue(_currentStudentRecord)
        WeeklyScheduleSlotCountTextBlock.Text = assignedSlots.ToString()
        WeeklyScheduleDayCountTextBlock.Text = activeDays.Count.ToString()
        WeeklyScheduleSectionTextBlock.Text = sectionLabel

        If assignedSlots = 0 Then
            WeeklyScheduleStatusTextBlock.Text = "No scheduled classes found for " & sectionLabel & "."
        Else
            WeeklyScheduleStatusTextBlock.Text =
                assignedSlots.ToString() & " scheduled classes loaded for " & sectionLabel & "."
        End If
    End Sub
End Class
