Imports System.Collections.Generic
Imports System.Data
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class TeacherScheduleView
    Private ReadOnly _subjectManagementService As New SubjectManagementService()
    Private ReadOnly _teacherScheduleManagementService As New TeacherScheduleManagementService()

    Private _subjectRecords As New List(Of SubjectRecord)()
    Private _timetableTable As DataTable

    Public Sub New()
        InitializeComponent()
        ApplyTimetable(StudentTimetablePresenter.CreateEmptyTable())
    End Sub

    Public Sub SetTeacherContext(teacherId As String, teacherName As String)
        Dim normalizedTeacherId As String = If(teacherId, String.Empty).Trim()
        ScheduleTeacherIdTextBlock.Text = normalizedTeacherId
        LoadTeacherTimetable(normalizedTeacherId)
    End Sub

    Public Function GetTimetableSnapshot() As DataTable
        If _timetableTable Is Nothing Then
            Return StudentTimetablePresenter.CreateEmptyTable()
        End If

        Return _timetableTable.Copy()
    End Function

    Private Sub LoadTeacherTimetable(teacherId As String)
        _subjectRecords = New List(Of SubjectRecord)()

        If String.IsNullOrWhiteSpace(teacherId) Then
            ApplyEmptyState("No teacher schedule loaded.")
            Return
        End If

        Dim subjectResult = _subjectManagementService.GetSubjects()
        If subjectResult IsNot Nothing AndAlso subjectResult.IsSuccess Then
            _subjectRecords = If(subjectResult.Data, New List(Of SubjectRecord)())
        End If

        Dim scheduleResult = _teacherScheduleManagementService.GetSchedules()
        If Not scheduleResult.IsSuccess Then
            ApplyEmptyState(scheduleResult.Message)
            Return
        End If

        Dim matchingSchedules As List(Of TeacherScheduleRecord) =
            FilterSchedulesForTeacher(scheduleResult.Data, teacherId)
        Dim timetableTable As DataTable =
            StudentTimetablePresenter.BuildTable(matchingSchedules, _subjectRecords)

        ApplyTimetable(timetableTable)
        UpdateScheduleSummary(matchingSchedules, teacherId)
    End Sub

    Private Function FilterSchedulesForTeacher(schedules As IEnumerable(Of TeacherScheduleRecord),
                                               teacherId As String) As List(Of TeacherScheduleRecord)
        Dim matches As New List(Of TeacherScheduleRecord)()
        Dim normalizedTeacherId As String = If(teacherId, String.Empty).Trim()
        If schedules Is Nothing OrElse String.IsNullOrWhiteSpace(normalizedTeacherId) Then
            Return matches
        End If

        For Each schedule As TeacherScheduleRecord In schedules
            If schedule Is Nothing Then
                Continue For
            End If

            If String.Equals(If(schedule.TeacherId, String.Empty).Trim(),
                             normalizedTeacherId,
                             StringComparison.OrdinalIgnoreCase) Then
                matches.Add(schedule)
            End If
        Next

        Return matches
    End Function

    Private Sub ApplyTimetable(table As DataTable)
        _timetableTable = table
        If _timetableTable Is Nothing Then
            _timetableTable = StudentTimetablePresenter.CreateEmptyTable()
        End If

        StudentTimetablePresenter.ConfigureDataGrid(TeacherTimetableDataGrid, _timetableTable)
        TeacherTimetableDataGrid.ItemsSource = _timetableTable.DefaultView
    End Sub

    Private Sub ApplyEmptyState(statusMessage As String)
        ApplyTimetable(StudentTimetablePresenter.CreateEmptyTable())
        AssignedSlotsTextBlock.Text = "0"
        ActiveDaysTextBlock.Text = "0"
        HandledSectionsTextBlock.Text = "0"
        TeacherScheduleStatusTextBlock.Text = If(String.IsNullOrWhiteSpace(statusMessage),
                                                 "No teacher schedule loaded.",
                                                 statusMessage)
    End Sub

    Private Sub UpdateScheduleSummary(matchingSchedules As IEnumerable(Of TeacherScheduleRecord),
                                      teacherId As String)
        Dim assignedSlots As Integer = 0
        Dim activeDays As New List(Of String)()
        Dim handledSections As New List(Of String)()

        If matchingSchedules IsNot Nothing Then
            For Each schedule As TeacherScheduleRecord In matchingSchedules
                If schedule Is Nothing Then
                    Continue For
                End If

                assignedSlots += 1
                AddUniqueValue(activeDays, StudentTimetablePresenter.NormalizeDayLabel(schedule.Day))
                AddUniqueValue(handledSections, If(schedule.Section, String.Empty).Trim())
            Next
        End If

        AssignedSlotsTextBlock.Text = assignedSlots.ToString()
        ActiveDaysTextBlock.Text = activeDays.Count.ToString()
        HandledSectionsTextBlock.Text = handledSections.Count.ToString()

        If assignedSlots = 0 Then
            TeacherScheduleStatusTextBlock.Text =
                "No scheduled classes found for teacher " & teacherId & "."
        Else
            TeacherScheduleStatusTextBlock.Text =
                assignedSlots.ToString() & " scheduled classes loaded for teacher " & teacherId & "."
        End If
    End Sub

    Private Sub AddUniqueValue(values As List(Of String), candidate As String)
        Dim normalizedCandidate As String = If(candidate, String.Empty).Trim()
        If values Is Nothing OrElse String.IsNullOrWhiteSpace(normalizedCandidate) Then
            Return
        End If

        For Each existingValue As String In values
            If String.Equals(existingValue, normalizedCandidate, StringComparison.OrdinalIgnoreCase) Then
                Return
            End If
        Next

        values.Add(normalizedCandidate)
    End Sub
End Class
