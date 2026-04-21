Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Media
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class AdminSchedulingView
    Private ReadOnly _dayHeaders As String() = New String() {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday"}
    Private Const TimetableSessionColumnName As String = "__Session"
    Private Const TimetableSearchColumnName As String = "__Search"

    Private _teacherOptions As New List(Of TeacherOption)()
    Private _subjectOptions As New List(Of SubjectOption)()
    Private _filteredTeacherOptions As New List(Of TeacherOption)()
    Private _filteredSubjectOptions As New List(Of SubjectOption)()
    Private _scheduleRecords As New List(Of TeacherScheduleRecord)()
    Private _timetableTable As DataTable
    Private _selectedTeacherId As String = String.Empty
    Private _searchTerm As String = String.Empty
    Private _isApplyingLookupFilters As Boolean = False

    Private ReadOnly _allDepartmentsFilterLabel As String = "All departments"
    Private ReadOnly _allProfessorsFilterLabel As String = "All professors"
    Private ReadOnly _withSchedulesFilterLabel As String = "With schedules"
    Private ReadOnly _withoutSchedulesFilterLabel As String = "Without schedules"
    Private ReadOnly _subjectManagementService As New SubjectManagementService()
    Private ReadOnly _teacherScheduleManagementService As New TeacherScheduleManagementService()
    Private ReadOnly _teacherManagementService As New TeacherManagementService()

    Private Class TeacherOption
        Public Property TeacherId As String
        Public Property FullName As String
        Public Property Department As String

        Public ReadOnly Property DisplayName As String
            Get
                Dim normalizedTeacherId As String = If(TeacherId, String.Empty).Trim()
                Dim normalizedFullName As String = If(FullName, String.Empty).Trim()

                If String.IsNullOrWhiteSpace(normalizedTeacherId) AndAlso String.IsNullOrWhiteSpace(normalizedFullName) Then
                    Return "--"
                End If

                If String.IsNullOrWhiteSpace(normalizedTeacherId) Then
                    Return normalizedFullName
                End If

                If String.IsNullOrWhiteSpace(normalizedFullName) Then
                    Return normalizedTeacherId
                End If

                Return normalizedFullName & " (" & normalizedTeacherId & ")"
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return DisplayName
        End Function
    End Class

    Private Class SubjectOption
        Public Property SubjectCode As String
        Public Property SubjectName As String
        Public Property Department As String
        Public Property YearLevel As String

        Public ReadOnly Property DisplayName As String
            Get
                Dim normalizedCode As String = If(SubjectCode, String.Empty).Trim()
                Dim normalizedName As String = If(SubjectName, String.Empty).Trim()

                If String.IsNullOrWhiteSpace(normalizedCode) AndAlso String.IsNullOrWhiteSpace(normalizedName) Then
                    Return "--"
                End If

                If String.IsNullOrWhiteSpace(normalizedCode) Then
                    Return normalizedName
                End If

                If String.IsNullOrWhiteSpace(normalizedName) Then
                    Return normalizedCode
                End If

                Return normalizedCode & " - " & normalizedName
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return DisplayName
        End Function
    End Class

    Public Sub New()
        InitializeComponent()

        InitializeDayOptions()
        InitializeSessionInputBoxes()
        InitializeLookupFilterControls()
        LoadScheduleRecords()
        LoadLookupData()

        RefreshProfessorContext()
        SetActionStatus("Select a professor to begin scheduling.")
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = NormalizeText(searchTerm)
        ApplyTimetableFilter()
    End Sub

    Public Sub RefreshData()
        LoadScheduleRecords()
        LoadLookupData()
    End Sub

    Private Sub InitializeDayOptions()
        DayComboBox.ItemsSource = _dayHeaders
        If _dayHeaders.Length > 0 Then
            DayComboBox.SelectedIndex = 0
        End If
    End Sub

    Private Sub InitializeSessionInputBoxes()
        If SessionStartHourTextBox IsNot Nothing Then
            SessionStartHourTextBox.Text = "00"
        End If

        If SessionStartMinuteTextBox IsNot Nothing Then
            SessionStartMinuteTextBox.Text = "00"
        End If

        If SessionEndHourTextBox IsNot Nothing Then
            SessionEndHourTextBox.Text = "00"
        End If

        If SessionEndMinuteTextBox IsNot Nothing Then
            SessionEndMinuteTextBox.Text = "00"
        End If
    End Sub

    Private Sub InitializeLookupFilterControls()
        _isApplyingLookupFilters = True
        Try
            If ProfessorScheduleFilterComboBox IsNot Nothing Then
                ProfessorScheduleFilterComboBox.ItemsSource = New List(Of String) From {
                    _allProfessorsFilterLabel,
                    _withSchedulesFilterLabel,
                    _withoutSchedulesFilterLabel
                }
                ProfessorScheduleFilterComboBox.SelectedIndex = 0
            End If

            If ProfessorDepartmentFilterComboBox IsNot Nothing Then
                ProfessorDepartmentFilterComboBox.ItemsSource = New List(Of String) From {_allDepartmentsFilterLabel}
                ProfessorDepartmentFilterComboBox.SelectedIndex = 0
            End If
        Finally
            _isApplyingLookupFilters = False
        End Try
    End Sub

    Private Sub LoadLookupData()
        Dim previousTeacherId As String = NormalizeText(_selectedTeacherId)
        Dim previousSubjectCode As String = NormalizeText(GetSelectedSubjectCode())

        _teacherOptions = ReadTeacherOptions()
        _subjectOptions = ReadSubjectOptions()

        UpdateDepartmentFilterOptions()
        ApplyLookupFilters(previousTeacherId, previousSubjectCode)
    End Sub

    Private Sub UpdateDepartmentFilterOptions()
        If ProfessorDepartmentFilterComboBox Is Nothing Then
            Return
        End If

        Dim previousFilterSelection As String = NormalizeText(TryCast(ProfessorDepartmentFilterComboBox.SelectedItem, String))
        If String.IsNullOrWhiteSpace(previousFilterSelection) Then
            previousFilterSelection = _allDepartmentsFilterLabel
        End If

        Dim departments As New List(Of String)()
        For Each optionEntry As TeacherOption In _teacherOptions
            If optionEntry Is Nothing Then
                Continue For
            End If

            Dim department As String = NormalizeText(optionEntry.Department)
            If String.IsNullOrWhiteSpace(department) Then
                Continue For
            End If

            If Not ContainsIgnoreCase(departments, department) Then
                departments.Add(department)
            End If
        Next

        departments.Sort(Function(left, right) String.Compare(left, right, StringComparison.OrdinalIgnoreCase))
        departments.Insert(0, _allDepartmentsFilterLabel)

        _isApplyingLookupFilters = True
        Try
            ProfessorDepartmentFilterComboBox.ItemsSource = departments
            If ContainsIgnoreCase(departments, previousFilterSelection) Then
                ProfessorDepartmentFilterComboBox.SelectedItem = departments.Find(Function(entry) String.Equals(entry, previousFilterSelection, StringComparison.OrdinalIgnoreCase))
            Else
                ProfessorDepartmentFilterComboBox.SelectedItem = _allDepartmentsFilterLabel
            End If
        Finally
            _isApplyingLookupFilters = False
        End Try
    End Sub

    Private Sub ApplyLookupFilters(Optional preferredTeacherId As String = Nothing,
                                   Optional preferredSubjectCode As String = Nothing)
        Dim teacherIdToKeep As String = NormalizeText(preferredTeacherId)
        If String.IsNullOrWhiteSpace(teacherIdToKeep) Then
            teacherIdToKeep = GetSelectedProfessorId()
        End If

        Dim subjectCodeToKeep As String = NormalizeText(preferredSubjectCode)
        If String.IsNullOrWhiteSpace(subjectCodeToKeep) Then
            subjectCodeToKeep = GetSelectedSubjectCode()
        End If

        Dim selectedDepartmentFilter As String = NormalizeDepartmentFilterValue(TryCast(ProfessorDepartmentFilterComboBox.SelectedItem, String))
        Dim selectedScheduleFilter As String = NormalizeText(TryCast(ProfessorScheduleFilterComboBox.SelectedItem, String))
        Dim professorFilter As String = NormalizeText(If(ProfessorFilterTextBox Is Nothing, String.Empty, ProfessorFilterTextBox.Text))

        _filteredTeacherOptions = BuildFilteredTeacherOptions(selectedDepartmentFilter, selectedScheduleFilter, professorFilter)

        _isApplyingLookupFilters = True
        Try
            ProfessorListDataGrid.ItemsSource = _filteredTeacherOptions
            If Not SelectProfessorById(teacherIdToKeep) Then
                ProfessorListDataGrid.SelectedItem = Nothing
            End If
        Finally
            _isApplyingLookupFilters = False
        End Try

        UpdateLookupCounts()
        UpdateProfessorScheduleBreakdown()
        UpdateScheduleBuilderSubtitle()
        RefreshProfessorContext(subjectCodeToKeep)
    End Sub

    Private Function BuildFilteredTeacherOptions(departmentFilter As String,
                                                 scheduleFilter As String,
                                                 professorFilter As String) As List(Of TeacherOption)
        Dim filtered As New List(Of TeacherOption)()

        For Each optionEntry As TeacherOption In _teacherOptions
            If optionEntry Is Nothing Then
                Continue For
            End If

            Dim teacherDepartment As String = NormalizeText(optionEntry.Department)
            If Not String.IsNullOrWhiteSpace(departmentFilter) AndAlso
               Not String.Equals(teacherDepartment, departmentFilter, StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim hasSchedule As Boolean = CountSlotsForTeacher(optionEntry.TeacherId) > 0
            If String.Equals(scheduleFilter, _withSchedulesFilterLabel, StringComparison.OrdinalIgnoreCase) AndAlso Not hasSchedule Then
                Continue For
            End If

            If String.Equals(scheduleFilter, _withoutSchedulesFilterLabel, StringComparison.OrdinalIgnoreCase) AndAlso hasSchedule Then
                Continue For
            End If

            If Not String.IsNullOrWhiteSpace(professorFilter) Then
                Dim candidateText As String =
                    NormalizeText(optionEntry.DisplayName) & "|" &
                    NormalizeText(optionEntry.TeacherId) & "|" &
                    teacherDepartment

                If candidateText.IndexOf(professorFilter, StringComparison.OrdinalIgnoreCase) < 0 Then
                    Continue For
                End If
            End If

            filtered.Add(optionEntry)
        Next

        filtered.Sort(Function(left, right) String.Compare(left.DisplayName, right.DisplayName, StringComparison.OrdinalIgnoreCase))
        Return filtered
    End Function

    Private Function BuildFilteredSubjectOptions(departmentFilter As String, subjectFilter As String) As List(Of SubjectOption)
        Dim filtered As New List(Of SubjectOption)()

        For Each optionEntry As SubjectOption In _subjectOptions
            If optionEntry Is Nothing Then
                Continue For
            End If

            Dim subjectDepartment As String = NormalizeText(optionEntry.Department)
            If Not String.IsNullOrWhiteSpace(departmentFilter) AndAlso
               Not String.IsNullOrWhiteSpace(subjectDepartment) AndAlso
               Not String.Equals(subjectDepartment, departmentFilter, StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            If Not String.IsNullOrWhiteSpace(subjectFilter) Then
                Dim candidateText As String =
                    NormalizeText(optionEntry.SubjectCode) & "|" &
                    NormalizeText(optionEntry.SubjectName) & "|" &
                    NormalizeText(optionEntry.DisplayName) & "|" &
                    subjectDepartment

                If candidateText.IndexOf(subjectFilter, StringComparison.OrdinalIgnoreCase) < 0 Then
                    Continue For
                End If
            End If

            filtered.Add(optionEntry)
        Next

        filtered.Sort(Function(left, right) String.Compare(left.DisplayName, right.DisplayName, StringComparison.OrdinalIgnoreCase))
        Return filtered
    End Function

    Private Function NormalizeDepartmentFilterValue(rawFilterValue As String) As String
        Dim normalized As String = NormalizeText(rawFilterValue)
        If String.IsNullOrWhiteSpace(normalized) Then
            Return String.Empty
        End If

        If String.Equals(normalized, _allDepartmentsFilterLabel, StringComparison.OrdinalIgnoreCase) Then
            Return String.Empty
        End If

        Return normalized
    End Function

    Private Function GetSelectedProfessorId() As String
        Dim selectedTeacher As TeacherOption = GetSelectedTeacherOption()
        If selectedTeacher Is Nothing Then
            Return String.Empty
        End If

        Return NormalizeText(selectedTeacher.TeacherId)
    End Function

    Private Function GetSelectedTeacherOption() As TeacherOption
        If ProfessorListDataGrid Is Nothing Then
            Return Nothing
        End If

        Return TryCast(ProfessorListDataGrid.SelectedItem, TeacherOption)
    End Function

    Private Sub LoadScheduleRecords()
        Dim result = _teacherScheduleManagementService.GetSchedules()
        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Scheduling",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            _scheduleRecords = New List(Of TeacherScheduleRecord)()
            Return
        End If

        Dim normalizedRecords As New List(Of TeacherScheduleRecord)()
        For Each record As TeacherScheduleRecord In result.Data
            Dim normalized As TeacherScheduleRecord = NormalizeScheduleRecord(record)
            If normalized IsNot Nothing Then
                normalizedRecords.Add(normalized)
            End If
        Next

        _scheduleRecords = normalizedRecords
    End Sub

    Private Function ReadTeacherOptions() As List(Of TeacherOption)
        Dim options As New List(Of TeacherOption)()
        Dim result = _teacherManagementService.GetTeachers()

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Scheduling",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return options
        End If

        For Each record As TeacherRecord In result.Data
            Dim teacherId As String = NormalizeText(record.EmployeeNumber)
            Dim fullName As String = NormalizeText(record.FullName)
            Dim department As String = NormalizeText(record.DepartmentDisplayName)

            If String.IsNullOrWhiteSpace(teacherId) AndAlso String.IsNullOrWhiteSpace(fullName) Then
                Continue For
            End If

            options.Add(New TeacherOption With {
                .TeacherId = teacherId,
                .FullName = fullName,
                .Department = department
            })
        Next

        options.Sort(Function(left, right) String.Compare(left.DisplayName, right.DisplayName, StringComparison.OrdinalIgnoreCase))
        Return options
    End Function

    Private Function ReadSubjectOptions() As List(Of SubjectOption)
        Dim options As New List(Of SubjectOption)()
        Dim result = _subjectManagementService.GetSubjects()

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Scheduling",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return options
        End If

        For Each record As SubjectRecord In result.Data
            Dim subjectCode As String = NormalizeText(record.SubjectCode)
            Dim subjectName As String = NormalizeText(record.SubjectName)

            If String.IsNullOrWhiteSpace(subjectCode) AndAlso
               String.IsNullOrWhiteSpace(subjectName) Then
                Continue For
            End If

            Dim department As String = NormalizeText(record.DepartmentDisplayName)
            If String.IsNullOrWhiteSpace(department) Then
                department = NormalizeText(record.CourseDisplayName)
            End If

            Dim yearLevel As String = NormalizeText(record.YearLevel)

            options.Add(New SubjectOption With {
                .SubjectCode = subjectCode,
                .SubjectName = subjectName,
                .Department = department,
                .YearLevel = yearLevel
            })
        Next

        options = DeduplicateSubjectOptions(options)
        options.Sort(Function(left, right) String.Compare(left.DisplayName, right.DisplayName, StringComparison.OrdinalIgnoreCase))
        Return options
    End Function

    Private Function DeduplicateSubjectOptions(source As IEnumerable(Of SubjectOption)) As List(Of SubjectOption)
        Dim deduped As New List(Of SubjectOption)()
        If source Is Nothing Then
            Return deduped
        End If

        Dim seenKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each candidate As SubjectOption In source
            If candidate Is Nothing Then
                Continue For
            End If

            Dim code As String = NormalizeText(candidate.SubjectCode)
            Dim name As String = NormalizeText(candidate.SubjectName)
            Dim department As String = NormalizeText(candidate.Department)
            Dim yearLevel As String = NormalizeText(candidate.YearLevel)
            Dim key As String = code & "|" & name & "|" & department & "|" & yearLevel
            If seenKeys.Contains(key) Then
                Continue For
            End If

            seenKeys.Add(key)
            deduped.Add(New SubjectOption With {
                .SubjectCode = code,
                .SubjectName = name,
                .Department = department,
                .YearLevel = yearLevel
            })
        Next

        Return deduped
    End Function

    Private Sub ProfessorDepartmentFilterComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If _isApplyingLookupFilters Then
            Return
        End If

        ApplyLookupFilters()
    End Sub

    Private Sub ProfessorScheduleFilterComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If _isApplyingLookupFilters Then
            Return
        End If

        ApplyLookupFilters()
    End Sub

    Private Sub ProfessorFilterTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        If _isApplyingLookupFilters Then
            Return
        End If

        ApplyLookupFilters()
    End Sub

    Private Sub ProfessorListDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If _isApplyingLookupFilters Then
            Return
        End If

        RefreshProfessorContext()
    End Sub

    Private Sub BackToProfessorListButton_Click(sender As Object, e As RoutedEventArgs)
        If ProfessorListDataGrid Is Nothing Then
            Return
        End If

        _isApplyingLookupFilters = True
        Try
            ProfessorListDataGrid.SelectedItem = Nothing
        Finally
            _isApplyingLookupFilters = False
        End Try

        RefreshProfessorContext()
    End Sub

    Private Sub SessionTimeTextBox_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        For Each typedCharacter As Char In e.Text
            If Not Char.IsDigit(typedCharacter) Then
                e.Handled = True
                Exit For
            End If
        Next
    End Sub

    Private Sub SessionTimeTextBox_LostFocus(sender As Object, e As RoutedEventArgs)
        Dim inputBox As System.Windows.Controls.TextBox = TryCast(sender, System.Windows.Controls.TextBox)
        If inputBox Is Nothing Then
            Return
        End If

        NormalizeSessionTextBox(inputBox)
    End Sub

    Private Sub RefreshSchedulingDataButton_Click(sender As Object, e As RoutedEventArgs)
        LoadScheduleRecords()
        LoadLookupData()
        SetActionStatus("Scheduling references refreshed.")
    End Sub

    Private Sub SaveScheduleSlotButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedTeacher As TeacherOption = GetSelectedTeacherOption()
        If selectedTeacher Is Nothing Then
            SetActionStatus("Select a professor first.", True)
            Return
        End If

        Dim selectedSubject As SubjectOption = TryCast(SubjectComboBox.SelectedItem, SubjectOption)
        If selectedSubject Is Nothing Then
            SetActionStatus("Select a subject before saving a slot.", True)
            Return
        End If

        Dim selectedDay As String = NormalizeDayLabel(TryCast(DayComboBox.SelectedItem, String))
        If String.IsNullOrWhiteSpace(selectedDay) Then
            SetActionStatus("Select a valid day.", True)
            Return
        End If

        Dim sessionValue As String = String.Empty
        If Not TryGetSessionRangeFromInputs(sessionValue) Then
            SetActionStatus("Session is required in 24-hour format (for example: 07:30 to 09:00).", True)
            Return
        End If

        Dim roomValue As String = NormalizeText(RoomTextBox.Text)
        Dim sectionValue As String = NormalizeText(SectionTextBox.Text)
        Dim normalizedTeacherId As String = NormalizeText(selectedTeacher.TeacherId)
        Dim normalizedSubjectCode As String = NormalizeText(selectedSubject.SubjectCode)
        Dim normalizedSubjectName As String = NormalizeText(selectedSubject.SubjectName)
        Dim saveResult =
            _teacherScheduleManagementService.SaveSchedule(New TeacherScheduleSaveRequest() With {
                .TeacherId = normalizedTeacherId,
                .Day = selectedDay,
                .Session = sessionValue,
                .Section = sectionValue,
                .SubjectCode = normalizedSubjectCode,
                .SubjectName = normalizedSubjectName,
                .Room = roomValue
            })
        If Not saveResult.IsSuccess Then
            SetActionStatus(saveResult.Message, True)
            Return
        End If

        LoadScheduleRecords()
        ApplyLookupFilters(normalizedTeacherId, normalizedSubjectCode)
        If String.Equals(NormalizeText(_selectedTeacherId), normalizedTeacherId, StringComparison.OrdinalIgnoreCase) Then
            SelectTimetableCell(sessionValue, selectedDay)
        End If
        SetActionStatus("Saved " & selectedDay & " " & sessionValue & ".")
    End Sub

    Private Sub RemoveScheduleSlotButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedTeacher As TeacherOption = GetSelectedTeacherOption()
        If selectedTeacher Is Nothing Then
            SetActionStatus("Select a professor first.", True)
            Return
        End If

        Dim selectedDay As String = NormalizeDayLabel(TryCast(DayComboBox.SelectedItem, String))
        If String.IsNullOrWhiteSpace(selectedDay) Then
            SetActionStatus("Select a valid day.", True)
            Return
        End If

        Dim sessionValue As String = String.Empty
        If Not TryGetSessionRangeFromInputs(sessionValue) Then
            SetActionStatus("Select the session start and end time to remove.", True)
            Return
        End If

        Dim deleteResult =
            _teacherScheduleManagementService.DeleteSchedule(NormalizeText(selectedTeacher.TeacherId),
                                                            selectedDay,
                                                            sessionValue)
        If Not deleteResult.IsSuccess Then
            SetActionStatus(deleteResult.Message, True)
            Return
        End If

        LoadScheduleRecords()
        ApplyLookupFilters(NormalizeText(selectedTeacher.TeacherId), GetSelectedSubjectCode())
        SetActionStatus("Removed " & selectedDay & " " & sessionValue & ".")
    End Sub

    Private Sub ClearScheduleFormButton_Click(sender As Object, e As RoutedEventArgs)
        If DayComboBox.Items.Count > 0 Then
            DayComboBox.SelectedIndex = 0
        End If

        If _filteredSubjectOptions.Count > 0 AndAlso SubjectComboBox.SelectedIndex < 0 Then
            SubjectComboBox.SelectedIndex = 0
        End If

        InitializeSessionInputBoxes()
        RoomTextBox.Text = String.Empty
        SectionTextBox.Text = String.Empty
        SetActionStatus("Form cleared.")
    End Sub

    Private Sub SchedulingTimetableDataGrid_SelectedCellsChanged(sender As Object, e As SelectedCellsChangedEventArgs)
        If SchedulingTimetableDataGrid.SelectedCells.Count = 0 Then
            Return
        End If

        Dim selectedCell As DataGridCellInfo = SchedulingTimetableDataGrid.SelectedCells(0)
        Dim selectedRowView As DataRowView = TryCast(selectedCell.Item, DataRowView)
        If selectedRowView Is Nothing OrElse selectedRowView.Row Is Nothing Then
            Return
        End If

        Dim sessionValue As String =
            NormalizeSession(ReadRowValue(selectedRowView.Row, TimetableSessionColumnName))
        If Not String.IsNullOrWhiteSpace(sessionValue) AndAlso Not String.Equals(sessionValue, "--", StringComparison.OrdinalIgnoreCase) Then
            SetSessionInputsFromSession(sessionValue)
        Else
            InitializeSessionInputBoxes()
        End If

        Dim selectedHeader As String = String.Empty
        If selectedCell.Column IsNot Nothing AndAlso selectedCell.Column.Header IsNot Nothing Then
            selectedHeader = NormalizeText(selectedCell.Column.Header.ToString())
        End If

        Dim selectedDay As String = NormalizeDayLabel(selectedHeader)
        If String.IsNullOrWhiteSpace(selectedDay) Then
            Return
        End If

        DayComboBox.SelectedItem = selectedDay

        Dim matchingRecord As TeacherScheduleRecord =
            FindScheduleRecord(_selectedTeacherId, selectedDay, sessionValue)
        If matchingRecord Is Nothing Then
            SectionTextBox.Text = String.Empty
            RoomTextBox.Text = String.Empty
            Return
        End If

        If Not SelectSubjectByCode(matchingRecord.SubjectCode) Then
            SelectSubjectByToken(matchingRecord.SubjectName)
        End If
        SectionTextBox.Text = NormalizeText(matchingRecord.Section)
        RoomTextBox.Text = NormalizeText(matchingRecord.Room)
    End Sub

    Private Sub RefreshProfessorContext(Optional preferredSubjectCode As String = Nothing)
        Dim selectedTeacher As TeacherOption = GetSelectedTeacherOption()
        If selectedTeacher Is Nothing Then
            _selectedTeacherId = String.Empty
            If TimetableSubtitleTextBlock IsNot Nothing Then
                TimetableSubtitleTextBlock.Text = "Select a professor to load timetable."
            End If
            _filteredSubjectOptions = New List(Of SubjectOption)()
            If SubjectComboBox IsNot Nothing Then
                SubjectComboBox.ItemsSource = _filteredSubjectOptions
            End If
        Else
            _selectedTeacherId = NormalizeText(selectedTeacher.TeacherId)
            If TimetableSubtitleTextBlock IsNot Nothing Then
                TimetableSubtitleTextBlock.Text = "Schedule for " & selectedTeacher.DisplayName
            End If
            ApplySubjectOptionsForTeacher(selectedTeacher, preferredSubjectCode)
        End If

        UpdateScheduleBuilderSubtitle()
        UpdateActionButtonState()
        RefreshTimetableForSelectedProfessor()
    End Sub

    Private Sub ApplySubjectOptionsForTeacher(selectedTeacher As TeacherOption,
                                              Optional preferredSubjectCode As String = Nothing)
        Dim teacherDepartment As String = NormalizeText(If(selectedTeacher Is Nothing, String.Empty, selectedTeacher.Department))
        Dim subjectCodeToKeep As String = NormalizeText(preferredSubjectCode)
        If String.IsNullOrWhiteSpace(subjectCodeToKeep) Then
            subjectCodeToKeep = GetSelectedSubjectCode()
        End If

        _filteredSubjectOptions = BuildFilteredSubjectOptions(teacherDepartment, String.Empty)
        If _filteredSubjectOptions.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(teacherDepartment) Then
            _filteredSubjectOptions = BuildFilteredSubjectOptions(String.Empty, String.Empty)
        End If

        _isApplyingLookupFilters = True
        Try
            SubjectComboBox.ItemsSource = _filteredSubjectOptions
            If Not SelectSubjectByCode(subjectCodeToKeep) AndAlso _filteredSubjectOptions.Count > 0 Then
                SubjectComboBox.SelectedIndex = 0
            End If
        Finally
            _isApplyingLookupFilters = False
        End Try
    End Sub

    Private Sub RefreshTimetableForSelectedProfessor()
        If String.IsNullOrWhiteSpace(_selectedTeacherId) Then
            _timetableTable = CreateEmptyTimetableTable()
        Else
            _timetableTable = CreateTimetableFromSchedules(GetSchedulesForTeacher(_selectedTeacherId))
        End If

        BuildTimetableColumns(_timetableTable)
        SchedulingTimetableDataGrid.ItemsSource = _timetableTable.DefaultView
        ApplyTimetableFilter()
    End Sub

    Private Function CreateTimetableFromSchedules(schedules As IEnumerable(Of TeacherScheduleRecord)) As DataTable
        Dim table As DataTable = CreateTimetableStructure()
        If schedules Is Nothing Then
            AddPlaceholderTimetableRow(table)
            Return table
        End If

        Dim orderedSessions As List(Of String) = GetOrderedSessions(schedules)
        If orderedSessions.Count = 0 Then
            AddPlaceholderTimetableRow(table)
            Return table
        End If

        For Each sessionEntry As String In orderedSessions
            AddTimetableRow(table, sessionEntry)
        Next

        For Each schedule As TeacherScheduleRecord In schedules
            Dim normalizedDay As String = NormalizeDayLabel(schedule.Day)
            Dim normalizedSession As String = NormalizeSession(schedule.Session)
            If String.IsNullOrWhiteSpace(normalizedDay) OrElse String.IsNullOrWhiteSpace(normalizedSession) Then
                Continue For
            End If

            Dim targetRow As DataRow = FindSessionRow(table, normalizedSession)
            If targetRow Is Nothing Then
                AddTimetableRow(table, normalizedSession)
                targetRow = FindSessionRow(table, normalizedSession)
            End If

            If targetRow IsNot Nothing Then
                targetRow(normalizedDay) = BuildTimetableCellDisplay(schedule)
            End If
        Next

        UpdateTimetableSearchValues(table)
        Return table
    End Function

    Private Function CreateTimetableStructure() As DataTable
        Dim table As New DataTable()
        table.Columns.Add(TimetableSessionColumnName, GetType(String))
        table.Columns.Add(TimetableSearchColumnName, GetType(String))
        For Each dayHeader As String In _dayHeaders
            table.Columns.Add(dayHeader, GetType(String))
        Next
        Return table
    End Function

    Private Function CreateEmptyTimetableTable() As DataTable
        Dim table As DataTable = CreateTimetableStructure()
        AddPlaceholderTimetableRow(table)
        Return table
    End Function

    Private Sub AddPlaceholderTimetableRow(table As DataTable)
        If table Is Nothing Then
            Return
        End If

        AddTimetableRow(table, "--")
        UpdateTimetableSearchValues(table)
    End Sub

    Private Sub AddTimetableRow(table As DataTable, sessionValue As String)
        If table Is Nothing Then
            Return
        End If

        Dim normalizedSession As String = NormalizeSession(sessionValue)
        If String.IsNullOrWhiteSpace(normalizedSession) Then
            normalizedSession = "--"
        End If

        Dim row As DataRow = table.NewRow()
        row(TimetableSessionColumnName) = normalizedSession
        row(TimetableSearchColumnName) = normalizedSession

        For Each dayHeader As String In _dayHeaders
            row(dayHeader) = "--"
        Next

        table.Rows.Add(row)
    End Sub

    Private Function FindSessionRow(table As DataTable, sessionValue As String) As DataRow
        If table Is Nothing OrElse String.IsNullOrWhiteSpace(sessionValue) Then
            Return Nothing
        End If

        Dim normalizedSession As String = NormalizeSession(sessionValue)

        For Each row As DataRow In table.Rows
            Dim candidateSession As String =
                NormalizeSession(ReadRowValue(row, TimetableSessionColumnName))
            If String.Equals(candidateSession, normalizedSession, StringComparison.OrdinalIgnoreCase) Then
                Return row
            End If
        Next

        Return Nothing
    End Function

    Private Sub UpdateTimetableSearchValues(table As DataTable)
        If table Is Nothing Then
            Return
        End If

        For Each row As DataRow In table.Rows
            Dim searchTokens As New List(Of String)()
            searchTokens.Add(ReadRowValue(row, TimetableSessionColumnName))

            For Each dayHeader As String In _dayHeaders
                Dim cellValue As String = ReadRowValue(row, dayHeader)
                If Not String.IsNullOrWhiteSpace(cellValue) AndAlso
                   Not String.Equals(cellValue, "--", StringComparison.OrdinalIgnoreCase) Then
                    searchTokens.Add(cellValue)
                End If
            Next

            row(TimetableSearchColumnName) =
                NormalizeText(String.Join(" ", searchTokens.ToArray()))
        Next
    End Sub

    Private Function GetOrderedSessions(schedules As IEnumerable(Of TeacherScheduleRecord)) As List(Of String)
        Dim orderedSessions As New List(Of String)()
        If schedules Is Nothing Then
            Return orderedSessions
        End If

        For Each schedule As TeacherScheduleRecord In schedules
            Dim normalizedSession As String = NormalizeSession(schedule.Session)
            If String.IsNullOrWhiteSpace(normalizedSession) Then
                Continue For
            End If

            If Not ContainsIgnoreCase(orderedSessions, normalizedSession) Then
                orderedSessions.Add(normalizedSession)
            End If
        Next

        orderedSessions.Sort(AddressOf CompareSessions)
        Return orderedSessions
    End Function

    Private Function CompareSessions(left As String, right As String) As Integer
        Dim leftStart As Integer
        Dim rightStart As Integer
        Dim leftHasTime As Boolean = TryParseSessionStart(left, leftStart)
        Dim rightHasTime As Boolean = TryParseSessionStart(right, rightStart)

        If leftHasTime AndAlso rightHasTime Then
            Return leftStart.CompareTo(rightStart)
        End If

        If leftHasTime Then
            Return -1
        End If

        If rightHasTime Then
            Return 1
        End If

        Return String.Compare(left, right, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function TryParseSessionStart(sessionValue As String, ByRef parsedStart As Integer) As Boolean
        Dim parsedEnd As Integer
        If TryParseSessionRange(sessionValue, parsedStart, parsedEnd) Then
            Return True
        End If

        Return False
    End Function

    Private Sub BuildTimetableColumns(sourceTable As DataTable)
        If SchedulingTimetableDataGrid Is Nothing OrElse sourceTable Is Nothing Then
            Return
        End If

        SchedulingTimetableDataGrid.Columns.Clear()

        Dim centeredTextStyle As New Style(GetType(TextBlock))
        centeredTextStyle.Setters.Add(New Setter(TextBlock.TextAlignmentProperty, TextAlignment.Center))
        centeredTextStyle.Setters.Add(New Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center))
        centeredTextStyle.Setters.Add(New Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap))

        For Each dayHeader As String In _dayHeaders
            If Not sourceTable.Columns.Contains(dayHeader) Then
                Continue For
            End If

            Dim dayColumn As New DataGridTextColumn() With {
                .Header = dayHeader,
                .Binding = New Binding(BuildDataTableBindingPath(dayHeader)),
                .IsReadOnly = True,
                .CanUserSort = False,
                .MinWidth = 120,
                .Width = New DataGridLength(1, DataGridLengthUnitType.Star),
                .ElementStyle = centeredTextStyle
            }
            SchedulingTimetableDataGrid.Columns.Add(dayColumn)
        Next
    End Sub

    Private Sub ApplyTimetableFilter()
        If _timetableTable Is Nothing Then
            UpdateTimetableSummary()
            Return
        End If

        If String.IsNullOrWhiteSpace(_searchTerm) Then
            _timetableTable.DefaultView.RowFilter = String.Empty
        Else
            Dim escapedTerm As String = EscapeLikeValue(_searchTerm)
            _timetableTable.DefaultView.RowFilter =
                "[" & TimetableSearchColumnName & "] LIKE '*" & escapedTerm & "*'"
        End If

        UpdateTimetableSummary()
    End Sub

    Private Sub UpdateLookupCounts()
        Dim totalProfessorCount As Integer = _teacherOptions.Count
        Dim totalSubjectCount As Integer = _subjectOptions.Count
        Dim visibleProfessorCount As Integer = _filteredTeacherOptions.Count
        Dim visibleSubjectCount As Integer = _filteredSubjectOptions.Count

        If visibleProfessorCount <> totalProfessorCount OrElse visibleSubjectCount <> totalSubjectCount Then
            SchedulingCountTextBlock.Text =
                visibleProfessorCount.ToString() & " of " & totalProfessorCount.ToString() & " professors | " &
                visibleSubjectCount.ToString() & " of " & totalSubjectCount.ToString() & " subjects"
        Else
            SchedulingCountTextBlock.Text = totalProfessorCount.ToString() & " professors | " & totalSubjectCount.ToString() & " subjects"
        End If
    End Sub

    Private Sub UpdateProfessorScheduleBreakdown()
        If ProfessorScheduleBreakdownTextBlock Is Nothing Then
            Return
        End If

        Dim withSchedules As Integer = 0
        For Each teacher As TeacherOption In _teacherOptions
            If teacher Is Nothing Then
                Continue For
            End If

            If CountSlotsForTeacher(teacher.TeacherId) > 0 Then
                withSchedules += 1
            End If
        Next

        Dim totalProfessors As Integer = _teacherOptions.Count
        Dim withoutSchedules As Integer = Math.Max(0, totalProfessors - withSchedules)
        Dim shownProfessors As Integer = _filteredTeacherOptions.Count
        ProfessorScheduleBreakdownTextBlock.Text =
            withSchedules.ToString() & " with schedules | " &
            withoutSchedules.ToString() & " without schedules | " &
            shownProfessors.ToString() & " shown"
    End Sub

    Private Sub UpdateTimetableSummary()
        If String.IsNullOrWhiteSpace(_selectedTeacherId) Then
            TimetableEntriesCountTextBlock.Text = "0 slots"
            Return
        End If

        Dim totalSlots As Integer = CountSlotsForTeacher(_selectedTeacherId)
        If String.IsNullOrWhiteSpace(_searchTerm) Then
            TimetableEntriesCountTextBlock.Text = totalSlots.ToString() & " slots"
        Else
            Dim visibleRows As Integer = If(_timetableTable Is Nothing, 0, _timetableTable.DefaultView.Count)
            TimetableEntriesCountTextBlock.Text = visibleRows.ToString() & " rows matched"
        End If
    End Sub

    Private Sub UpdateScheduleBuilderSubtitle()
        If _teacherOptions.Count = 0 Then
            ScheduleBuilderSubtitleTextBlock.Text = "No professors found. Add teachers first."
            Return
        End If

        If _filteredTeacherOptions.Count = 0 Then
            ScheduleBuilderSubtitleTextBlock.Text = "No professors match the current list filters."
            Return
        End If

        If String.IsNullOrWhiteSpace(_selectedTeacherId) Then
            ScheduleBuilderSubtitleTextBlock.Text = "Select a professor from the list to start building."
            Return
        End If

        If _filteredSubjectOptions.Count = 0 Then
            ScheduleBuilderSubtitleTextBlock.Text = "No subjects available for the selected professor."
            Return
        End If

        ScheduleBuilderSubtitleTextBlock.Text = "Assign subject, day, session, section, and room for the selected professor."
    End Sub

    Private Sub UpdateActionButtonState()
        Dim hasTeacher As Boolean = Not String.IsNullOrWhiteSpace(_selectedTeacherId)
        Dim hasSubjects As Boolean = _filteredSubjectOptions.Count > 0

        SaveScheduleSlotButton.IsEnabled = hasTeacher AndAlso hasSubjects
        RemoveScheduleSlotButton.IsEnabled = hasTeacher
        DayComboBox.IsEnabled = hasTeacher
        SessionStartHourTextBox.IsEnabled = hasTeacher
        SessionStartMinuteTextBox.IsEnabled = hasTeacher
        SessionEndHourTextBox.IsEnabled = hasTeacher
        SessionEndMinuteTextBox.IsEnabled = hasTeacher
        SectionTextBox.IsEnabled = hasTeacher
        RoomTextBox.IsEnabled = hasTeacher
        SubjectComboBox.IsEnabled = hasTeacher AndAlso hasSubjects

        ScheduleBuilderEditorPanel.Visibility = If(hasTeacher, System.Windows.Visibility.Visible, System.Windows.Visibility.Collapsed)
        ScheduleBuilderEmptyStatePanel.Visibility = If(hasTeacher, System.Windows.Visibility.Collapsed, System.Windows.Visibility.Visible)

        If hasTeacher Then
            ProfessorListPanelRow.Height = New System.Windows.GridLength(0)
            ListBuilderSpacerRow.Height = New System.Windows.GridLength(0)
            ScheduleBuilderPanelRow.Height = New System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
            ProfessorListPanelBorder.Visibility = System.Windows.Visibility.Collapsed
            ScheduleBuilderPanelBorder.Visibility = System.Windows.Visibility.Visible
            ScheduleLeftColumn.Width = New System.Windows.GridLength(0.8, System.Windows.GridUnitType.Star)
            ScheduleMiddleColumn.Width = New System.Windows.GridLength(10)
            ScheduleRightColumn.Width = New System.Windows.GridLength(1.2, System.Windows.GridUnitType.Star)
            ProfessorTimetableBorder.Visibility = System.Windows.Visibility.Visible
        Else
            ProfessorListPanelRow.Height = New System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
            ListBuilderSpacerRow.Height = New System.Windows.GridLength(0)
            ScheduleBuilderPanelRow.Height = New System.Windows.GridLength(0)
            ProfessorListPanelBorder.Visibility = System.Windows.Visibility.Visible
            ScheduleBuilderPanelBorder.Visibility = System.Windows.Visibility.Collapsed
            ScheduleLeftColumn.Width = New System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
            ScheduleMiddleColumn.Width = New System.Windows.GridLength(0)
            ScheduleRightColumn.Width = New System.Windows.GridLength(0)
            ProfessorTimetableBorder.Visibility = System.Windows.Visibility.Collapsed
        End If
    End Sub

    Private Function GetSchedulesForTeacher(teacherId As String) As List(Of TeacherScheduleRecord)
        Dim matches As New List(Of TeacherScheduleRecord)()
        Dim normalizedTeacherId As String = NormalizeText(teacherId)
        If String.IsNullOrWhiteSpace(normalizedTeacherId) Then
            Return matches
        End If

        For Each record As TeacherScheduleRecord In _scheduleRecords
            If record Is Nothing Then
                Continue For
            End If

            If String.Equals(NormalizeText(record.TeacherId), normalizedTeacherId, StringComparison.OrdinalIgnoreCase) Then
                matches.Add(record)
            End If
        Next

        Return matches
    End Function

    Private Function CountSlotsForTeacher(teacherId As String) As Integer
        Return GetSchedulesForTeacher(teacherId).Count
    End Function

    Private Function FindScheduleRecord(teacherId As String,
                                        dayValue As String,
                                        sessionValue As String) As TeacherScheduleRecord
        Dim normalizedTeacherId As String = NormalizeText(teacherId)
        Dim normalizedDay As String = NormalizeDayLabel(dayValue)
        Dim normalizedSession As String = NormalizeSession(sessionValue)
        If String.IsNullOrWhiteSpace(normalizedTeacherId) OrElse
           String.IsNullOrWhiteSpace(normalizedDay) OrElse
           String.IsNullOrWhiteSpace(normalizedSession) OrElse
           String.Equals(normalizedSession, "--", StringComparison.OrdinalIgnoreCase) Then
            Return Nothing
        End If

        For Each record As TeacherScheduleRecord In _scheduleRecords
            If record Is Nothing Then
                Continue For
            End If

            If String.Equals(NormalizeText(record.TeacherId), normalizedTeacherId, StringComparison.OrdinalIgnoreCase) AndAlso
               String.Equals(NormalizeDayLabel(record.Day), normalizedDay, StringComparison.OrdinalIgnoreCase) AndAlso
               String.Equals(NormalizeSession(record.Session), normalizedSession, StringComparison.OrdinalIgnoreCase) Then
                Return record
            End If
        Next

        Return Nothing
    End Function

    Private Function NormalizeScheduleRecord(record As TeacherScheduleRecord) As TeacherScheduleRecord
        If record Is Nothing Then
            Return Nothing
        End If

        Dim normalizedTeacherId As String = NormalizeText(record.TeacherId)
        Dim normalizedDay As String = NormalizeDayLabel(record.Day)
        Dim normalizedSession As String = NormalizeSession(record.Session)
        Dim normalizedSection As String = NormalizeText(record.Section)
        Dim normalizedSubjectCode As String = NormalizeText(record.SubjectCode)
        Dim normalizedSubjectName As String = NormalizeText(record.SubjectName)

        If String.IsNullOrWhiteSpace(normalizedTeacherId) OrElse
           String.IsNullOrWhiteSpace(normalizedDay) OrElse
           String.IsNullOrWhiteSpace(normalizedSession) Then
            Return Nothing
        End If

        If String.IsNullOrWhiteSpace(normalizedSubjectCode) AndAlso String.IsNullOrWhiteSpace(normalizedSubjectName) Then
            Return Nothing
        End If

        Return New TeacherScheduleRecord With {
            .ScheduleId = record.ScheduleId,
            .TeacherRecordId = record.TeacherRecordId,
            .TeacherId = normalizedTeacherId,
            .TeacherName = NormalizeText(record.TeacherName),
            .Day = normalizedDay,
            .Session = normalizedSession,
            .Section = normalizedSection,
            .SubjectCode = normalizedSubjectCode,
            .SubjectName = normalizedSubjectName,
            .Room = NormalizeText(record.Room)
        }
    End Function

    Private Function NormalizeDayLabel(dayValue As String) As String
        Dim normalized As String = NormalizeText(dayValue)
        If String.IsNullOrWhiteSpace(normalized) Then
            Return String.Empty
        End If

        Select Case normalized.ToLowerInvariant()
            Case "monday", "mon"
                Return "Monday"
            Case "tuesday", "tue", "tues"
                Return "Tuesday"
            Case "wednesday", "wed"
                Return "Wednesday"
            Case "thursday", "thu", "thur", "thurs"
                Return "Thursday"
            Case "friday", "fri"
                Return "Friday"
            Case Else
                Return String.Empty
        End Select
    End Function

    Private Function NormalizeSession(sessionValue As String) As String
        Dim startMinutes As Integer
        Dim endMinutes As Integer
        If TryParseSessionRange(sessionValue, startMinutes, endMinutes) Then
            Return FormatSessionRange(startMinutes, endMinutes)
        End If

        Return NormalizeText(sessionValue)
    End Function

    Private Sub NormalizeSessionTextBox(inputBox As System.Windows.Controls.TextBox)
        If inputBox Is Nothing Then
            Return
        End If

        Dim isHourBox As Boolean =
            inputBox.Name.IndexOf("Hour", StringComparison.OrdinalIgnoreCase) >= 0
        Dim maxValue As Integer = If(isHourBox, 24, 59)

        Dim numericValue As Integer
        If Not Integer.TryParse(NormalizeText(inputBox.Text), NumberStyles.Integer, CultureInfo.InvariantCulture, numericValue) Then
            numericValue = 0
        End If

        numericValue = Math.Max(0, Math.Min(maxValue, numericValue))
        inputBox.Text = numericValue.ToString("00", CultureInfo.InvariantCulture)
    End Sub

    Private Function TryGetSessionRangeFromInputs(ByRef sessionValue As String) As Boolean
        sessionValue = String.Empty

        NormalizeSessionTextBox(SessionStartHourTextBox)
        NormalizeSessionTextBox(SessionStartMinuteTextBox)
        NormalizeSessionTextBox(SessionEndHourTextBox)
        NormalizeSessionTextBox(SessionEndMinuteTextBox)

        Dim startMinutes As Integer
        Dim endMinutes As Integer
        If Not TryParseSessionInput(SessionStartHourTextBox.Text, SessionStartMinuteTextBox.Text, startMinutes) Then
            Return False
        End If

        If Not TryParseSessionInput(SessionEndHourTextBox.Text, SessionEndMinuteTextBox.Text, endMinutes) Then
            Return False
        End If

        If endMinutes <= startMinutes Then
            Return False
        End If

        sessionValue = FormatSessionRange(startMinutes, endMinutes)
        Return True
    End Function

    Private Sub SetSessionInputsFromSession(sessionValue As String)
        Dim startMinutes As Integer
        Dim endMinutes As Integer
        If Not TryParseSessionRange(sessionValue, startMinutes, endMinutes) Then
            InitializeSessionInputBoxes()
            Return
        End If

        SessionStartHourTextBox.Text = (startMinutes \ 60).ToString("00", CultureInfo.InvariantCulture)
        SessionStartMinuteTextBox.Text = (startMinutes Mod 60).ToString("00", CultureInfo.InvariantCulture)
        SessionEndHourTextBox.Text = (endMinutes \ 60).ToString("00", CultureInfo.InvariantCulture)
        SessionEndMinuteTextBox.Text = (endMinutes Mod 60).ToString("00", CultureInfo.InvariantCulture)
    End Sub

    Private Function TryParseSessionInput(hourValue As String,
                                          minuteValue As String,
                                          ByRef totalMinutes As Integer) As Boolean
        totalMinutes = 0

        Dim normalizedHour As String = NormalizeText(hourValue)
        Dim normalizedMinute As String = NormalizeText(minuteValue)
        If String.IsNullOrWhiteSpace(normalizedHour) Then
            normalizedHour = "00"
        End If
        If String.IsNullOrWhiteSpace(normalizedMinute) Then
            normalizedMinute = "00"
        End If

        Dim parsedHour As Integer
        Dim parsedMinute As Integer
        If Not Integer.TryParse(normalizedHour, NumberStyles.Integer, CultureInfo.InvariantCulture, parsedHour) OrElse
           Not Integer.TryParse(normalizedMinute, NumberStyles.Integer, CultureInfo.InvariantCulture, parsedMinute) Then
            Return False
        End If

        If parsedHour < 0 OrElse parsedHour > 24 Then
            Return False
        End If

        If parsedMinute < 0 OrElse parsedMinute > 59 Then
            Return False
        End If

        If parsedHour = 24 AndAlso parsedMinute <> 0 Then
            Return False
        End If

        totalMinutes = (parsedHour * 60) + parsedMinute
        Return totalMinutes >= 0 AndAlso totalMinutes <= 1440
    End Function

    Private Function TryParseSessionRange(sessionValue As String,
                                          ByRef parsedStartMinutes As Integer,
                                          ByRef parsedEndMinutes As Integer) As Boolean
        Dim normalizedSession As String = NormalizeText(sessionValue)
        If String.IsNullOrWhiteSpace(normalizedSession) Then
            Return False
        End If

        Dim startToken As String = String.Empty
        Dim endToken As String = String.Empty
        If Not TrySplitSessionRange(normalizedSession, startToken, endToken) Then
            Return False
        End If

        If Not TryParseClockToken(startToken, parsedStartMinutes) OrElse
           Not TryParseClockToken(endToken, parsedEndMinutes) Then
            Return False
        End If

        If parsedEndMinutes <= parsedStartMinutes Then
            Return False
        End If

        Return True
    End Function

    Private Function TrySplitSessionRange(sessionValue As String, ByRef startToken As String, ByRef endToken As String) As Boolean
        startToken = String.Empty
        endToken = String.Empty

        Dim separator As String = " - "
        Dim separatorIndex As Integer = sessionValue.IndexOf(separator, StringComparison.Ordinal)
        If separatorIndex <= 0 Then
            separator = " to "
            separatorIndex = sessionValue.IndexOf(separator, StringComparison.OrdinalIgnoreCase)
        End If

        If separatorIndex <= 0 Then
            separator = "-"
            separatorIndex = sessionValue.IndexOf(separator, StringComparison.Ordinal)
        End If

        If separatorIndex <= 0 Then
            Return False
        End If

        startToken = NormalizeText(sessionValue.Substring(0, separatorIndex))
        endToken = NormalizeText(sessionValue.Substring(separatorIndex + separator.Length))
        If String.IsNullOrWhiteSpace(startToken) OrElse String.IsNullOrWhiteSpace(endToken) Then
            Return False
        End If

        Return True
    End Function

    Private Function TryParseClockToken(value As String, ByRef parsedMinutes As Integer) As Boolean
        parsedMinutes = 0
        Dim normalized As String = NormalizeText(value)
        If String.IsNullOrWhiteSpace(normalized) Then
            Return False
        End If

        Dim splitTokens As String() = normalized.Split(":"c)
        If splitTokens.Length = 2 Then
            Dim parsedHour As Integer
            Dim parsedMinute As Integer
            If Integer.TryParse(NormalizeText(splitTokens(0)), NumberStyles.Integer, CultureInfo.InvariantCulture, parsedHour) AndAlso
               Integer.TryParse(NormalizeText(splitTokens(1)), NumberStyles.Integer, CultureInfo.InvariantCulture, parsedMinute) Then
                If parsedHour >= 0 AndAlso parsedHour <= 24 AndAlso
                   parsedMinute >= 0 AndAlso parsedMinute <= 59 AndAlso
                   (parsedHour < 24 OrElse parsedMinute = 0) Then
                    parsedMinutes = (parsedHour * 60) + parsedMinute
                    Return True
                End If
            End If
        End If

        Dim supportedFormats As String() = New String() {
            "HH:mm",
            "H:mm",
            "hh:mm tt",
            "h:mm tt",
            "hh:mmtt",
            "h:mmtt"
        }

        Dim parsedTime As DateTime
        If DateTime.TryParseExact(normalized,
                                  supportedFormats,
                                  CultureInfo.InvariantCulture,
                                  DateTimeStyles.AllowWhiteSpaces,
                                  parsedTime) Then
            parsedMinutes = (parsedTime.Hour * 60) + parsedTime.Minute
            Return True
        End If

        If DateTime.TryParse(normalized, CultureInfo.CurrentCulture, DateTimeStyles.NoCurrentDateDefault, parsedTime) Then
            parsedMinutes = (parsedTime.Hour * 60) + parsedTime.Minute
            Return True
        End If

        If DateTime.TryParse(normalized, CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault, parsedTime) Then
            parsedMinutes = (parsedTime.Hour * 60) + parsedTime.Minute
            Return True
        End If

        Return False
    End Function

    Private Function FormatSessionRange(startMinutes As Integer, endMinutes As Integer) As String
        Return FormatClockMinutes(startMinutes) &
               " - " &
               FormatClockMinutes(endMinutes)
    End Function

    Private Function FormatClockMinutes(totalMinutes As Integer) As String
        Dim normalizedTotal As Integer = Math.Max(0, Math.Min(1440, totalMinutes))
        Dim hourValue As Integer = normalizedTotal \ 60
        Dim minuteValue As Integer = normalizedTotal Mod 60
        Return hourValue.ToString("00", CultureInfo.InvariantCulture) &
               ":" &
               minuteValue.ToString("00", CultureInfo.InvariantCulture)
    End Function

    Private Function BuildTimetableCellDisplay(record As TeacherScheduleRecord) As String
        If record Is Nothing Then
            Return "--"
        End If

        Dim sessionValue As String = NormalizeSession(record.Session)
        Dim subjectCode As String = NormalizeText(record.SubjectCode)
        Dim subjectName As String = NormalizeText(record.SubjectName)
        Dim sectionValue As String = NormalizeText(record.Section)
        Dim yearLevelValue As String = ResolveSubjectYearLevel(record)
        Dim sectionLine As String = BuildCompactSectionLabel(sectionValue, yearLevelValue)
        Dim subjectValue As String = subjectCode
        If Not String.IsNullOrWhiteSpace(subjectCode) AndAlso
           Not String.IsNullOrWhiteSpace(subjectName) AndAlso
           Not String.Equals(subjectCode, subjectName, StringComparison.OrdinalIgnoreCase) Then
            subjectValue = subjectCode & " - " & subjectName
        ElseIf String.IsNullOrWhiteSpace(subjectValue) Then
            subjectValue = subjectName
        End If

        Dim roomValue As String = NormalizeText(record.Room)
        If String.IsNullOrWhiteSpace(sessionValue) Then
            sessionValue = "--"
        End If
        If String.IsNullOrWhiteSpace(subjectValue) Then
            subjectValue = "--"
        End If
        If String.IsNullOrWhiteSpace(roomValue) Then
            roomValue = "--"
        End If

        Return String.Join(Environment.NewLine, New String() {
            sessionValue,
            subjectValue,
            sectionLine,
            roomValue
        })
    End Function

    Private Function BuildCompactSectionLabel(sectionValue As String, yearLevelValue As String) As String
        Return StudentScheduleHelper.BuildCompactSectionValue(sectionValue,
                                                              yearLevelValue)
    End Function

    Private Function NormalizeSectionToken(sectionValue As String) As String
        Dim normalizedSection As String = NormalizeText(sectionValue)
        If String.IsNullOrWhiteSpace(normalizedSection) Then
            Return String.Empty
        End If

        If normalizedSection.StartsWith("Section:", StringComparison.OrdinalIgnoreCase) Then
            normalizedSection = NormalizeText(normalizedSection.Substring("Section:".Length))
        ElseIf normalizedSection.StartsWith("Section ", StringComparison.OrdinalIgnoreCase) Then
            normalizedSection = NormalizeText(normalizedSection.Substring("Section ".Length))
        End If

        Return normalizedSection.Replace(" ", String.Empty)
    End Function

    Private Function NormalizeYearLevelToken(yearLevelValue As String) As String
        Dim normalizedYearLevel As String = NormalizeText(yearLevelValue)
        If String.IsNullOrWhiteSpace(normalizedYearLevel) Then
            Return String.Empty
        End If

        Dim compactDigits As String = String.Empty
        For Each characterValue As Char In normalizedYearLevel
            If Char.IsDigit(characterValue) Then
                compactDigits &= characterValue
            ElseIf compactDigits.Length > 0 Then
                Exit For
            End If
        Next

        If compactDigits.Length > 0 Then
            Return compactDigits
        End If

        Dim lowerYearLevel As String = normalizedYearLevel.ToLowerInvariant()
        If lowerYearLevel.Contains("first") Then
            Return "1"
        End If
        If lowerYearLevel.Contains("second") Then
            Return "2"
        End If
        If lowerYearLevel.Contains("third") Then
            Return "3"
        End If
        If lowerYearLevel.Contains("fourth") Then
            Return "4"
        End If
        If lowerYearLevel.Contains("fifth") Then
            Return "5"
        End If
        If lowerYearLevel.Contains("sixth") Then
            Return "6"
        End If

        Return normalizedYearLevel.Replace(" ", String.Empty)
    End Function

    Private Function ResolveSubjectYearLevel(record As TeacherScheduleRecord) As String
        If record Is Nothing Then
            Return String.Empty
        End If

        Dim subjectCode As String = NormalizeText(record.SubjectCode)
        Dim subjectName As String = NormalizeText(record.SubjectName)

        For Each optionEntry As SubjectOption In _subjectOptions
            If optionEntry Is Nothing Then
                Continue For
            End If

            If Not String.IsNullOrWhiteSpace(subjectCode) AndAlso
               String.Equals(NormalizeText(optionEntry.SubjectCode),
                             subjectCode,
                             StringComparison.OrdinalIgnoreCase) Then
                Return NormalizeText(optionEntry.YearLevel)
            End If

            If String.IsNullOrWhiteSpace(subjectCode) AndAlso
               Not String.IsNullOrWhiteSpace(subjectName) AndAlso
               String.Equals(NormalizeText(optionEntry.SubjectName),
                             subjectName,
                             StringComparison.OrdinalIgnoreCase) Then
                Return NormalizeText(optionEntry.YearLevel)
            End If
        Next

        Return String.Empty
    End Function

    Private Sub SelectTimetableCell(sessionValue As String, dayHeader As String)
        If SchedulingTimetableDataGrid Is Nothing OrElse _timetableTable Is Nothing Then
            Return
        End If

        Dim normalizedSession As String = NormalizeSession(sessionValue)
        Dim normalizedDay As String = NormalizeDayLabel(dayHeader)
        If String.IsNullOrWhiteSpace(normalizedSession) OrElse String.IsNullOrWhiteSpace(normalizedDay) Then
            Return
        End If

        For Each rowView As DataRowView In _timetableTable.DefaultView
            Dim candidateSession As String =
                NormalizeSession(ReadRowValue(rowView.Row, TimetableSessionColumnName))
            If Not String.Equals(candidateSession, normalizedSession, StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            For Each column As DataGridColumn In SchedulingTimetableDataGrid.Columns
                If column Is Nothing OrElse column.Header Is Nothing Then
                    Continue For
                End If

                If String.Equals(NormalizeText(column.Header.ToString()), normalizedDay, StringComparison.OrdinalIgnoreCase) Then
                    Dim targetCell As New DataGridCellInfo(rowView, column)

                    SchedulingTimetableDataGrid.UnselectAllCells()
                    SchedulingTimetableDataGrid.CurrentCell = targetCell
                    SchedulingTimetableDataGrid.SelectedCells.Add(targetCell)
                    SchedulingTimetableDataGrid.ScrollIntoView(rowView, column)
                    Return
                End If
            Next
        Next
    End Sub

    Private Function SelectProfessorById(teacherId As String) As Boolean
        If ProfessorListDataGrid Is Nothing Then
            Return False
        End If

        Dim normalizedTeacherId As String = NormalizeText(teacherId)
        If String.IsNullOrWhiteSpace(normalizedTeacherId) Then
            Return False
        End If

        For Each optionEntry As TeacherOption In _filteredTeacherOptions
            If optionEntry Is Nothing Then
                Continue For
            End If

            If String.Equals(NormalizeText(optionEntry.TeacherId), normalizedTeacherId, StringComparison.OrdinalIgnoreCase) Then
                ProfessorListDataGrid.SelectedItem = optionEntry
                ProfessorListDataGrid.ScrollIntoView(optionEntry)
                Return True
            End If
        Next

        Return False
    End Function

    Private Function SelectSubjectByCode(subjectCode As String) As Boolean
        Dim normalizedSubjectCode As String = NormalizeText(subjectCode)
        If String.IsNullOrWhiteSpace(normalizedSubjectCode) Then
            Return False
        End If

        Dim source As IEnumerable(Of SubjectOption) = _filteredSubjectOptions
        If source Is Nothing OrElse _filteredSubjectOptions.Count = 0 Then
            source = _subjectOptions
        End If

        For Each optionEntry As SubjectOption In source
            If optionEntry Is Nothing Then
                Continue For
            End If

            If String.Equals(NormalizeText(optionEntry.SubjectCode), normalizedSubjectCode, StringComparison.OrdinalIgnoreCase) Then
                SubjectComboBox.SelectedItem = optionEntry
                Return True
            End If
        Next

        Return False
    End Function

    Private Function SelectSubjectByToken(subjectToken As String) As Boolean
        Dim normalizedToken As String = NormalizeText(subjectToken)
        If String.IsNullOrWhiteSpace(normalizedToken) Then
            Return False
        End If

        Dim source As IEnumerable(Of SubjectOption) = _filteredSubjectOptions
        If source Is Nothing OrElse _filteredSubjectOptions.Count = 0 Then
            source = _subjectOptions
        End If

        For Each optionEntry As SubjectOption In source
            If optionEntry Is Nothing Then
                Continue For
            End If

            Dim code As String = NormalizeText(optionEntry.SubjectCode)
            Dim name As String = NormalizeText(optionEntry.SubjectName)

            If String.Equals(code, normalizedToken, StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(name, normalizedToken, StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(optionEntry.DisplayName, normalizedToken, StringComparison.OrdinalIgnoreCase) Then
                SubjectComboBox.SelectedItem = optionEntry
                Return True
            End If
        Next

        Return False
    End Function

    Private Function GetSelectedSubjectCode() As String
        Dim selectedSubject As SubjectOption = TryCast(SubjectComboBox.SelectedItem, SubjectOption)
        If selectedSubject Is Nothing Then
            Return String.Empty
        End If

        Return NormalizeText(selectedSubject.SubjectCode)
    End Function

    Private Sub SetActionStatus(message As String, Optional isError As Boolean = False)
        If SchedulingActionStatusTextBlock Is Nothing Then
            Return
        End If

        SchedulingActionStatusTextBlock.Text = NormalizeText(message)
        If String.IsNullOrWhiteSpace(SchedulingActionStatusTextBlock.Text) Then
            SchedulingActionStatusTextBlock.Text = "Ready."
        End If

        Dim brushKey As String = If(isError, "DashboardDangerBrush", "DashboardTextMutedBrush")
        Dim brush As Brush = TryCast(FindResource(brushKey), Brush)
        If brush IsNot Nothing Then
            SchedulingActionStatusTextBlock.Foreground = brush
        End If
    End Sub

    Private Function BuildDataTableBindingPath(columnName As String) As String
        Dim safeColumnName As String = If(columnName, String.Empty)
        Return "[" & safeColumnName.Replace("]", "]]") & "]"
    End Function

    Private Function ReadRowValue(row As DataRow, columnName As String) As String
        If row Is Nothing OrElse row.Table Is Nothing OrElse Not row.Table.Columns.Contains(columnName) Then
            Return String.Empty
        End If

        If row.IsNull(columnName) Then
            Return String.Empty
        End If

        Return NormalizeText(row(columnName).ToString())
    End Function

    Private Function EscapeLikeValue(value As String) As String
        Dim safeValue As String = If(value, String.Empty)
        Return safeValue.
            Replace("'", "''").
            Replace("[", "[[]").
            Replace("]", "[]]").
            Replace("*", "[*]").
            Replace("%", "[%]")
    End Function

    Private Function NormalizeText(value As String) As String
        Return If(value, String.Empty).Trim()
    End Function

    Private Function ContainsIgnoreCase(values As IEnumerable(Of String), candidate As String) As Boolean
        If values Is Nothing Then
            Return False
        End If

        For Each value As String In values
            If String.Equals(value, candidate, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Return False
    End Function
End Class
