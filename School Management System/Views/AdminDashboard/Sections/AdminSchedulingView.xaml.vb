Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Text.Json
Imports System.Windows.Data
Imports System.Windows.Media

Class AdminSchedulingView
    Private ReadOnly _dayHeaders As String() = New String() {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday"}
    Private ReadOnly _teachersStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "teachers.json")
    Private ReadOnly _subjectsStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "subjects.json")
    Private ReadOnly _coursesStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "courses.json")
    Private ReadOnly _schedulesStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "professor-schedules.json")
    Private ReadOnly _jsonOptions As New JsonSerializerOptions() With {
        .WriteIndented = True,
        .PropertyNameCaseInsensitive = True
    }

    Private _teacherOptions As New List(Of TeacherOption)()
    Private _subjectOptions As New List(Of SubjectOption)()
    Private _filteredTeacherOptions As New List(Of TeacherOption)()
    Private _filteredSubjectOptions As New List(Of SubjectOption)()
    Private _scheduleRecords As New List(Of ScheduleStorageRecord)()
    Private _timetableTable As DataTable
    Private _selectedTeacherId As String = String.Empty
    Private _searchTerm As String = String.Empty
    Private _isApplyingLookupFilters As Boolean = False

    Private ReadOnly _allDepartmentsFilterLabel As String = "All departments"
    Private ReadOnly _allProfessorsFilterLabel As String = "All professors"
    Private ReadOnly _withSchedulesFilterLabel As String = "With schedules"
    Private ReadOnly _withoutSchedulesFilterLabel As String = "Without schedules"

    Private Class TeacherStorageRecord
        Public Property TeacherId As String
        Public Property FullName As String
        Public Property Department As String
    End Class

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

    Private Class ScheduleStorageRecord
        Public Property TeacherId As String
        Public Property TeacherName As String
        Public Property Day As String
        Public Property Session As String
        Public Property Section As String
        Public Property SubjectCode As String
        Public Property SubjectName As String
        Public Property Room As String
    End Class

    Public Sub New()
        InitializeComponent()

        InitializeDayOptions()
        InitializeSessionTimeOptions()
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

    Private Sub InitializeDayOptions()
        DayComboBox.ItemsSource = _dayHeaders
        If _dayHeaders.Length > 0 Then
            DayComboBox.SelectedIndex = 0
        End If
    End Sub

    Private Sub InitializeSessionTimeOptions()
        Dim clockTimes As List(Of String) = BuildClockTimes()
        SessionStartComboBox.ItemsSource = clockTimes
        SessionEndComboBox.ItemsSource = clockTimes
    End Sub

    Private Function BuildClockTimes() As List(Of String)
        Dim values As New List(Of String)()
        For hour As Integer = 0 To 23
            For minute As Integer = 0 To 30 Step 30
                values.Add(hour.ToString("00", CultureInfo.InvariantCulture) &
                           ":" &
                           minute.ToString("00", CultureInfo.InvariantCulture))
            Next
        Next

        Return values
    End Function

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
        Dim subjectFilter As String = NormalizeText(If(SubjectFilterTextBox Is Nothing, String.Empty, SubjectFilterTextBox.Text))

        _filteredTeacherOptions = BuildFilteredTeacherOptions(selectedDepartmentFilter, selectedScheduleFilter, professorFilter)
        _filteredSubjectOptions = BuildFilteredSubjectOptions(selectedDepartmentFilter, subjectFilter)

        _isApplyingLookupFilters = True
        Try
            ProfessorComboBox.ItemsSource = _filteredTeacherOptions
            SubjectComboBox.ItemsSource = _filteredSubjectOptions

            If Not SelectProfessorById(teacherIdToKeep) AndAlso _filteredTeacherOptions.Count > 0 Then
                ProfessorComboBox.SelectedIndex = 0
            End If

            If Not SelectSubjectByCode(subjectCodeToKeep) AndAlso _filteredSubjectOptions.Count > 0 Then
                SubjectComboBox.SelectedIndex = 0
            End If
        Finally
            _isApplyingLookupFilters = False
        End Try

        ProfessorComboBox.IsEnabled = _filteredTeacherOptions.Count > 0
        SubjectComboBox.IsEnabled = _filteredSubjectOptions.Count > 0

        UpdateLookupCounts()
        UpdateProfessorScheduleBreakdown()
        UpdateScheduleBuilderSubtitle()
        RefreshProfessorContext()
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
        Dim selectedTeacher As TeacherOption = TryCast(ProfessorComboBox.SelectedItem, TeacherOption)
        If selectedTeacher Is Nothing Then
            Return String.Empty
        End If

        Return NormalizeText(selectedTeacher.TeacherId)
    End Function

    Private Sub LoadScheduleRecords()
        Dim loadedRecords As List(Of ScheduleStorageRecord) = ReadSchedulesFromStorage()
        Dim normalizedRecords As New List(Of ScheduleStorageRecord)()

        For Each record As ScheduleStorageRecord In loadedRecords
            Dim normalized As ScheduleStorageRecord = NormalizeScheduleRecord(record)
            If normalized IsNot Nothing Then
                normalizedRecords.Add(normalized)
            End If
        Next

        _scheduleRecords = normalizedRecords
    End Sub

    Private Function ReadTeacherOptions() As List(Of TeacherOption)
        Dim options As New List(Of TeacherOption)()
        If Not File.Exists(_teachersStoragePath) Then
            Return options
        End If

        Try
            Dim json As String = File.ReadAllText(_teachersStoragePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return options
            End If

            Dim records As List(Of TeacherStorageRecord) =
                JsonSerializer.Deserialize(Of List(Of TeacherStorageRecord))(json, _jsonOptions)
            If records Is Nothing Then
                Return options
            End If

            For Each record As TeacherStorageRecord In records
                Dim teacherId As String = NormalizeText(record.TeacherId)
                Dim fullName As String = NormalizeText(record.FullName)
                Dim department As String = NormalizeText(record.Department)

                If String.IsNullOrWhiteSpace(teacherId) AndAlso String.IsNullOrWhiteSpace(fullName) Then
                    Continue For
                End If

                options.Add(New TeacherOption With {
                    .TeacherId = teacherId,
                    .FullName = fullName,
                    .Department = department
                })
            Next
        Catch ex As Exception
            MessageBox.Show("Unable to load teachers for scheduling." & Environment.NewLine & ex.Message,
                            "Scheduling",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
        End Try

        options.Sort(Function(left, right) String.Compare(left.DisplayName, right.DisplayName, StringComparison.OrdinalIgnoreCase))
        Return options
    End Function

    Private Function ReadSubjectOptions() As List(Of SubjectOption)
        Dim options As List(Of SubjectOption) =
            ReadSubjectOptionsFromJsonFile(
                _subjectsStoragePath,
                New String() {"SubjectCode", "Subject Code", "Code", "CourseCode", "Course Code"},
                New String() {"SubjectName", "Subject Name", "Name", "Title", "CourseTitle", "Course Title"},
                New String() {"Department", "Dept", "Program", "Course", "Course Code", "CourseCode"},
                "subjects")

        If options.Count = 0 Then
            options = ReadSubjectOptionsFromJsonFile(
                _coursesStoragePath,
                New String() {"CourseCode", "Course Code", "Code", "SubjectCode", "Subject Code"},
                New String() {"CourseTitle", "Course Title", "Title", "Name", "SubjectName", "Subject Name"},
                New String() {"Department", "Dept", "Program", "Department Name", "College"},
                "courses")
        End If

        options = DeduplicateSubjectOptions(options)
        options.Sort(Function(left, right) String.Compare(left.DisplayName, right.DisplayName, StringComparison.OrdinalIgnoreCase))
        Return options
    End Function

    Private Function ReadSubjectOptionsFromJsonFile(storagePath As String,
                                                    codeAliases As String(),
                                                    nameAliases As String(),
                                                    departmentAliases As String(),
                                                    sourceLabel As String) As List(Of SubjectOption)
        Dim options As New List(Of SubjectOption)()
        If Not File.Exists(storagePath) Then
            Return options
        End If

        Try
            Dim json As String = File.ReadAllText(storagePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return options
            End If

            Using document As JsonDocument = JsonDocument.Parse(json)
                If document.RootElement.ValueKind <> JsonValueKind.Array Then
                    Return options
                End If

                For Each item As JsonElement In document.RootElement.EnumerateArray()
                    If item.ValueKind <> JsonValueKind.Object Then
                        Continue For
                    End If

                    Dim code As String = NormalizeText(ExtractJsonValue(item, codeAliases))
                    Dim name As String = NormalizeText(ExtractJsonValue(item, nameAliases))
                    Dim department As String = NormalizeText(ExtractJsonValue(item, departmentAliases))

                    If String.IsNullOrWhiteSpace(code) AndAlso String.IsNullOrWhiteSpace(name) Then
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(code) Then
                        code = name
                    End If

                    If String.IsNullOrWhiteSpace(name) Then
                        name = code
                    End If

                    options.Add(New SubjectOption With {
                        .SubjectCode = code,
                        .SubjectName = name,
                        .Department = department
                    })
                Next
            End Using
        Catch ex As Exception
            MessageBox.Show("Unable to load " & sourceLabel & " for scheduling." & Environment.NewLine & ex.Message,
                            "Scheduling",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
        End Try

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
            Dim key As String = code & "|" & name & "|" & department
            If seenKeys.Contains(key) Then
                Continue For
            End If

            seenKeys.Add(key)
            deduped.Add(New SubjectOption With {
                .SubjectCode = code,
                .SubjectName = name,
                .Department = department
            })
        Next

        Return deduped
    End Function

    Private Function ExtractJsonValue(element As JsonElement, propertyAliases As IEnumerable(Of String)) As String
        If element.ValueKind <> JsonValueKind.Object OrElse propertyAliases Is Nothing Then
            Return String.Empty
        End If

        For Each aliasName As String In propertyAliases
            Dim normalizedAlias As String = NormalizeText(aliasName)
            If String.IsNullOrWhiteSpace(normalizedAlias) Then
                Continue For
            End If

            For Each propertyEntry As JsonProperty In element.EnumerateObject()
                If Not String.Equals(propertyEntry.Name, normalizedAlias, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Select Case propertyEntry.Value.ValueKind
                    Case JsonValueKind.String
                        Return NormalizeText(propertyEntry.Value.GetString())
                    Case JsonValueKind.Number, JsonValueKind.True, JsonValueKind.False
                        Return NormalizeText(propertyEntry.Value.ToString())
                End Select
            Next
        Next

        Return String.Empty
    End Function

    Private Function ReadSchedulesFromStorage() As List(Of ScheduleStorageRecord)
        Dim schedules As New List(Of ScheduleStorageRecord)()
        If Not File.Exists(_schedulesStoragePath) Then
            Return schedules
        End If

        Try
            Dim json As String = File.ReadAllText(_schedulesStoragePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return schedules
            End If

            Dim records As List(Of ScheduleStorageRecord) =
                JsonSerializer.Deserialize(Of List(Of ScheduleStorageRecord))(json, _jsonOptions)
            If records Is Nothing Then
                Return schedules
            End If

            schedules.AddRange(records)
        Catch ex As Exception
            MessageBox.Show("Unable to load saved schedule records." & Environment.NewLine & ex.Message,
                            "Scheduling",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
        End Try

        Return schedules
    End Function

    Private Function PersistSchedulesToStorage() As Boolean
        Try
            Dim storageDirectory As String = Path.GetDirectoryName(_schedulesStoragePath)
            If Not String.IsNullOrWhiteSpace(storageDirectory) Then
                Directory.CreateDirectory(storageDirectory)
            End If

            Dim json As String = JsonSerializer.Serialize(_scheduleRecords, _jsonOptions)
            File.WriteAllText(_schedulesStoragePath, json)
            Return True
        Catch ex As Exception
            MessageBox.Show("Unable to save schedule records." & Environment.NewLine & ex.Message,
                            "Scheduling",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            Return False
        End Try
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

    Private Sub SubjectFilterTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        If _isApplyingLookupFilters Then
            Return
        End If

        ApplyLookupFilters()
    End Sub

    Private Sub ProfessorComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If _isApplyingLookupFilters Then
            Return
        End If

        RefreshProfessorContext()
    End Sub

    Private Sub RefreshSchedulingDataButton_Click(sender As Object, e As RoutedEventArgs)
        LoadScheduleRecords()
        LoadLookupData()
        SetActionStatus("Scheduling references refreshed.")
    End Sub

    Private Sub SaveScheduleSlotButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedTeacher As TeacherOption = TryCast(ProfessorComboBox.SelectedItem, TeacherOption)
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
        Dim normalizedTeacherName As String = NormalizeText(selectedTeacher.FullName)
        Dim normalizedSubjectCode As String = NormalizeText(selectedSubject.SubjectCode)
        Dim normalizedSubjectName As String = NormalizeText(selectedSubject.SubjectName)

        RemoveSchedulesBySlot(normalizedTeacherId, selectedDay, sessionValue)

        _scheduleRecords.Add(New ScheduleStorageRecord With {
            .TeacherId = normalizedTeacherId,
            .TeacherName = normalizedTeacherName,
            .Day = selectedDay,
            .Session = sessionValue,
            .Section = sectionValue,
            .SubjectCode = normalizedSubjectCode,
            .SubjectName = normalizedSubjectName,
            .Room = roomValue
        })

        If Not PersistSchedulesToStorage() Then
            Return
        End If

        ApplyLookupFilters(normalizedTeacherId, normalizedSubjectCode)
        If String.Equals(NormalizeText(_selectedTeacherId), normalizedTeacherId, StringComparison.OrdinalIgnoreCase) Then
            SelectTimetableCell(sessionValue, selectedDay)
        End If
        SetActionStatus("Saved " & selectedDay & " " & sessionValue & ".")
    End Sub

    Private Sub RemoveScheduleSlotButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedTeacher As TeacherOption = TryCast(ProfessorComboBox.SelectedItem, TeacherOption)
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

        Dim removedCount As Integer = RemoveSchedulesBySlot(NormalizeText(selectedTeacher.TeacherId), selectedDay, sessionValue)
        If removedCount = 0 Then
            SetActionStatus("No matching slot found for this professor.", True)
            Return
        End If

        If Not PersistSchedulesToStorage() Then
            Return
        End If

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

        SessionStartComboBox.SelectedIndex = -1
        SessionEndComboBox.SelectedIndex = -1
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

        Dim sessionValue As String = NormalizeSession(ReadRowValue(selectedRowView.Row, "Session"))
        If Not String.IsNullOrWhiteSpace(sessionValue) AndAlso Not String.Equals(sessionValue, "--", StringComparison.OrdinalIgnoreCase) Then
            SetSessionSelectorsFromSession(sessionValue)
        Else
            SessionStartComboBox.SelectedIndex = -1
            SessionEndComboBox.SelectedIndex = -1
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

        Dim cellValue As String = NormalizeText(ReadRowValue(selectedRowView.Row, selectedDay))
        If String.IsNullOrWhiteSpace(cellValue) OrElse String.Equals(cellValue, "--", StringComparison.OrdinalIgnoreCase) Then
            SectionTextBox.Text = String.Empty
            RoomTextBox.Text = String.Empty
            Return
        End If

        Dim subjectToken As String = ExtractSubjectToken(cellValue)
        Dim sectionToken As String = ExtractSectionToken(cellValue)
        Dim roomToken As String = ExtractRoomToken(cellValue)

        SelectSubjectByToken(subjectToken)
        SectionTextBox.Text = sectionToken
        If Not String.IsNullOrWhiteSpace(roomToken) Then
            RoomTextBox.Text = roomToken
        Else
            RoomTextBox.Text = String.Empty
        End If
    End Sub

    Private Sub RefreshProfessorContext()
        Dim selectedTeacher As TeacherOption = TryCast(ProfessorComboBox.SelectedItem, TeacherOption)
        If selectedTeacher Is Nothing Then
            _selectedTeacherId = String.Empty
            SelectedProfessorTextBlock.Text = "--"
            TimetableSubtitleTextBlock.Text = "Select a professor to load timetable."
        Else
            _selectedTeacherId = NormalizeText(selectedTeacher.TeacherId)
            SelectedProfessorTextBlock.Text = selectedTeacher.DisplayName
            TimetableSubtitleTextBlock.Text = "Schedule for " & selectedTeacher.DisplayName
        End If

        UpdateActionButtonState()
        RefreshTimetableForSelectedProfessor()
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

    Private Function CreateTimetableFromSchedules(schedules As IEnumerable(Of ScheduleStorageRecord)) As DataTable
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
            Dim row As DataRow = table.NewRow()
            row("Session") = sessionEntry
            For Each dayHeader As String In _dayHeaders
                row(dayHeader) = "--"
            Next
            table.Rows.Add(row)
        Next

        For Each schedule As ScheduleStorageRecord In schedules
            Dim normalizedDay As String = NormalizeDayLabel(schedule.Day)
            Dim normalizedSession As String = NormalizeSession(schedule.Session)
            If String.IsNullOrWhiteSpace(normalizedDay) OrElse String.IsNullOrWhiteSpace(normalizedSession) Then
                Continue For
            End If

            Dim targetRow As DataRow = FindSessionRow(table, normalizedSession)
            If targetRow Is Nothing Then
                targetRow = table.NewRow()
                targetRow("Session") = normalizedSession
                For Each dayHeader As String In _dayHeaders
                    targetRow(dayHeader) = "--"
                Next
                table.Rows.Add(targetRow)
            End If

            targetRow(normalizedDay) = BuildTimetableCellDisplay(schedule)
        Next

        Return table
    End Function

    Private Function CreateTimetableStructure() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Session", GetType(String))
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

        Dim row As DataRow = table.NewRow()
        row("Session") = "--"
        For Each dayHeader As String In _dayHeaders
            row(dayHeader) = "--"
        Next
        table.Rows.Add(row)
    End Sub

    Private Function FindSessionRow(table As DataTable, sessionValue As String) As DataRow
        If table Is Nothing OrElse String.IsNullOrWhiteSpace(sessionValue) Then
            Return Nothing
        End If

        For Each row As DataRow In table.Rows
            Dim candidate As String = NormalizeSession(ReadRowValue(row, "Session"))
            If String.Equals(candidate, sessionValue, StringComparison.OrdinalIgnoreCase) Then
                Return row
            End If
        Next

        Return Nothing
    End Function

    Private Function GetOrderedSessions(schedules As IEnumerable(Of ScheduleStorageRecord)) As List(Of String)
        Dim orderedSessions As New List(Of String)()
        If schedules Is Nothing Then
            Return orderedSessions
        End If

        For Each schedule As ScheduleStorageRecord In schedules
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
        Dim leftStart As DateTime
        Dim rightStart As DateTime
        Dim leftHasTime As Boolean = TryParseSessionStart(left, leftStart)
        Dim rightHasTime As Boolean = TryParseSessionStart(right, rightStart)

        If leftHasTime AndAlso rightHasTime Then
            Return DateTime.Compare(leftStart, rightStart)
        End If

        If leftHasTime Then
            Return -1
        End If

        If rightHasTime Then
            Return 1
        End If

        Return String.Compare(left, right, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function TryParseSessionStart(sessionValue As String, ByRef parsedStart As DateTime) As Boolean
        Dim parsedEnd As DateTime
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

        Dim sessionColumn As New DataGridTextColumn() With {
            .Header = "Session",
            .Binding = New Binding(BuildDataTableBindingPath("Session")),
            .IsReadOnly = True,
            .CanUserSort = False,
            .MinWidth = 170,
            .Width = New DataGridLength(1.2, DataGridLengthUnitType.Star)
        }
        SchedulingTimetableDataGrid.Columns.Add(sessionColumn)

        For Each dayHeader As String In _dayHeaders
            If Not sourceTable.Columns.Contains(dayHeader) Then
                Continue For
            End If

            Dim dayColumn As New DataGridTextColumn() With {
                .Header = dayHeader,
                .Binding = New Binding(BuildDataTableBindingPath(dayHeader)),
                .IsReadOnly = True,
                .CanUserSort = False,
                .Width = New DataGridLength(1, DataGridLengthUnitType.Star)
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
            Dim filterParts As New List(Of String) From {
                "[Session] LIKE '*" & escapedTerm & "*'"
            }

            For Each dayHeader As String In _dayHeaders
                filterParts.Add("[" & dayHeader & "] LIKE '*" & escapedTerm & "*'")
            Next

            _timetableTable.DefaultView.RowFilter = String.Join(" OR ", filterParts)
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

        If _subjectOptions.Count = 0 Then
            ScheduleBuilderSubtitleTextBlock.Text = "No subjects found. Add subjects first."
            Return
        End If

        If _filteredTeacherOptions.Count = 0 Then
            ScheduleBuilderSubtitleTextBlock.Text = "No professors match the current filters."
            Return
        End If

        If _filteredSubjectOptions.Count = 0 Then
            ScheduleBuilderSubtitleTextBlock.Text = "No subjects match the current filters."
            Return
        End If

        ScheduleBuilderSubtitleTextBlock.Text = "Choose a professor and assign subjects, session, section, and room."
    End Sub

    Private Sub UpdateActionButtonState()
        Dim hasTeacher As Boolean = Not String.IsNullOrWhiteSpace(_selectedTeacherId)
        Dim hasSubjects As Boolean = _filteredSubjectOptions.Count > 0

        SaveScheduleSlotButton.IsEnabled = hasTeacher AndAlso hasSubjects
        RemoveScheduleSlotButton.IsEnabled = hasTeacher
        DayComboBox.IsEnabled = hasTeacher
        SessionStartComboBox.IsEnabled = hasTeacher
        SessionEndComboBox.IsEnabled = hasTeacher
        SectionTextBox.IsEnabled = hasTeacher
        RoomTextBox.IsEnabled = hasTeacher
    End Sub

    Private Function GetSchedulesForTeacher(teacherId As String) As List(Of ScheduleStorageRecord)
        Dim matches As New List(Of ScheduleStorageRecord)()
        Dim normalizedTeacherId As String = NormalizeText(teacherId)
        If String.IsNullOrWhiteSpace(normalizedTeacherId) Then
            Return matches
        End If

        For Each record As ScheduleStorageRecord In _scheduleRecords
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

    Private Function RemoveSchedulesBySlot(teacherId As String, dayValue As String, sessionValue As String) As Integer
        Dim normalizedTeacherId As String = NormalizeText(teacherId)
        Dim normalizedDay As String = NormalizeDayLabel(dayValue)
        Dim normalizedSession As String = NormalizeSession(sessionValue)
        Dim removedCount As Integer = 0

        For index As Integer = _scheduleRecords.Count - 1 To 0 Step -1
            Dim candidate As ScheduleStorageRecord = _scheduleRecords(index)
            If candidate Is Nothing Then
                Continue For
            End If

            If String.Equals(NormalizeText(candidate.TeacherId), normalizedTeacherId, StringComparison.OrdinalIgnoreCase) AndAlso
               String.Equals(NormalizeDayLabel(candidate.Day), normalizedDay, StringComparison.OrdinalIgnoreCase) AndAlso
               String.Equals(NormalizeSession(candidate.Session), normalizedSession, StringComparison.OrdinalIgnoreCase) Then
                _scheduleRecords.RemoveAt(index)
                removedCount += 1
            End If
        Next

        Return removedCount
    End Function

    Private Function NormalizeScheduleRecord(record As ScheduleStorageRecord) As ScheduleStorageRecord
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

        Return New ScheduleStorageRecord With {
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
        Dim startTime As DateTime
        Dim endTime As DateTime
        If TryParseSessionRange(sessionValue, startTime, endTime) Then
            Return FormatSessionRange(startTime, endTime)
        End If

        Return NormalizeText(sessionValue)
    End Function

    Private Function TryGetSessionRangeFromInputs(ByRef sessionValue As String) As Boolean
        sessionValue = String.Empty

        Dim selectedStart As String = NormalizeText(TryCast(SessionStartComboBox.SelectedItem, String))
        Dim selectedEnd As String = NormalizeText(TryCast(SessionEndComboBox.SelectedItem, String))
        If String.IsNullOrWhiteSpace(selectedStart) OrElse String.IsNullOrWhiteSpace(selectedEnd) Then
            Return False
        End If

        Dim startTime As DateTime
        Dim endTime As DateTime
        If Not TryParseClockTime(selectedStart, startTime) OrElse Not TryParseClockTime(selectedEnd, endTime) Then
            Return False
        End If

        If endTime <= startTime Then
            Return False
        End If

        sessionValue = FormatSessionRange(startTime, endTime)
        Return True
    End Function

    Private Sub SetSessionSelectorsFromSession(sessionValue As String)
        Dim startTime As DateTime
        Dim endTime As DateTime
        If Not TryParseSessionRange(sessionValue, startTime, endTime) Then
            SessionStartComboBox.SelectedIndex = -1
            SessionEndComboBox.SelectedIndex = -1
            Return
        End If

        SelectSessionComboValue(SessionStartComboBox, startTime.ToString("HH:mm", CultureInfo.InvariantCulture))
        SelectSessionComboValue(SessionEndComboBox, endTime.ToString("HH:mm", CultureInfo.InvariantCulture))
    End Sub

    Private Sub SelectSessionComboValue(combo As System.Windows.Controls.ComboBox, value As String)
        If combo Is Nothing Then
            Return
        End If

        Dim normalizedValue As String = NormalizeText(value)
        For Each item As Object In combo.Items
            Dim candidate As String = NormalizeText(TryCast(item, String))
            If String.Equals(candidate, normalizedValue, StringComparison.OrdinalIgnoreCase) Then
                combo.SelectedItem = item
                Return
            End If
        Next

        combo.SelectedIndex = -1
    End Sub

    Private Function TryParseSessionRange(sessionValue As String, ByRef parsedStart As DateTime, ByRef parsedEnd As DateTime) As Boolean
        Dim normalizedSession As String = NormalizeText(sessionValue)
        If String.IsNullOrWhiteSpace(normalizedSession) Then
            Return False
        End If

        Dim startToken As String = String.Empty
        Dim endToken As String = String.Empty
        If Not TrySplitSessionRange(normalizedSession, startToken, endToken) Then
            Return False
        End If

        If Not TryParseClockTime(startToken, parsedStart) OrElse Not TryParseClockTime(endToken, parsedEnd) Then
            Return False
        End If

        If parsedEnd <= parsedStart Then
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

    Private Function TryParseClockTime(value As String, ByRef parsedTime As DateTime) As Boolean
        Dim normalized As String = NormalizeText(value)
        If String.IsNullOrWhiteSpace(normalized) Then
            Return False
        End If

        Dim supportedFormats As String() = New String() {
            "HH:mm",
            "H:mm",
            "hh:mm tt",
            "h:mm tt",
            "hh:mmtt",
            "h:mmtt"
        }

        If DateTime.TryParseExact(normalized,
                                  supportedFormats,
                                  CultureInfo.InvariantCulture,
                                  DateTimeStyles.AllowWhiteSpaces,
                                  parsedTime) Then
            Return True
        End If

        If DateTime.TryParse(normalized, CultureInfo.CurrentCulture, DateTimeStyles.NoCurrentDateDefault, parsedTime) Then
            Return True
        End If

        Return DateTime.TryParse(normalized, CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault, parsedTime)
    End Function

    Private Function FormatSessionRange(startTime As DateTime, endTime As DateTime) As String
        Return startTime.ToString("HH:mm", CultureInfo.InvariantCulture) &
               " - " &
               endTime.ToString("HH:mm", CultureInfo.InvariantCulture)
    End Function

    Private Function BuildTimetableCellDisplay(record As ScheduleStorageRecord) As String
        If record Is Nothing Then
            Return "--"
        End If

        Dim subjectCode As String = NormalizeText(record.SubjectCode)
        Dim subjectName As String = NormalizeText(record.SubjectName)
        Dim section As String = NormalizeText(record.Section)
        Dim room As String = NormalizeText(record.Room)

        Dim subjectToken As String = subjectCode
        If String.IsNullOrWhiteSpace(subjectToken) Then
            subjectToken = subjectName
        End If

        If String.IsNullOrWhiteSpace(subjectToken) Then
            subjectToken = "--"
        End If

        If Not String.IsNullOrWhiteSpace(section) Then
            subjectToken &= " [Sec: " & section & "]"
        End If

        If Not String.IsNullOrWhiteSpace(room) Then
            subjectToken &= " (" & room & ")"
        End If

        Return subjectToken
    End Function

    Private Function ExtractSubjectToken(cellDisplayValue As String) As String
        Dim normalized As String = RemoveRoomSuffix(cellDisplayValue)
        If String.IsNullOrWhiteSpace(normalized) Then
            Return String.Empty
        End If

        Dim sectionStartIndex As Integer = normalized.LastIndexOf(" [Sec:", StringComparison.OrdinalIgnoreCase)
        If sectionStartIndex > 0 AndAlso normalized.EndsWith("]", StringComparison.Ordinal) Then
            Return NormalizeText(normalized.Substring(0, sectionStartIndex))
        End If

        Return normalized
    End Function

    Private Function ExtractSectionToken(cellDisplayValue As String) As String
        Dim normalized As String = RemoveRoomSuffix(cellDisplayValue)
        If String.IsNullOrWhiteSpace(normalized) Then
            Return String.Empty
        End If

        Dim sectionStartIndex As Integer = normalized.LastIndexOf(" [Sec:", StringComparison.OrdinalIgnoreCase)
        If sectionStartIndex <= 0 OrElse Not normalized.EndsWith("]", StringComparison.Ordinal) Then
            Return String.Empty
        End If

        Dim tokenStartIndex As Integer = sectionStartIndex + " [Sec:".Length
        If tokenStartIndex >= normalized.Length - 1 Then
            Return String.Empty
        End If

        Dim tokenLength As Integer = normalized.Length - tokenStartIndex - 1
        Return NormalizeText(normalized.Substring(tokenStartIndex, tokenLength))
    End Function

    Private Function ExtractRoomToken(cellDisplayValue As String) As String
        Dim normalized As String = NormalizeText(cellDisplayValue)
        If String.IsNullOrWhiteSpace(normalized) Then
            Return String.Empty
        End If

        Dim roomStartIndex As Integer = normalized.LastIndexOf(" (", StringComparison.Ordinal)
        If roomStartIndex <= 0 OrElse Not normalized.EndsWith(")", StringComparison.Ordinal) Then
            Return String.Empty
        End If

        Dim roomValue As String = normalized.Substring(roomStartIndex + 2, normalized.Length - roomStartIndex - 3)
        Return NormalizeText(roomValue)
    End Function

    Private Function RemoveRoomSuffix(cellDisplayValue As String) As String
        Dim normalized As String = NormalizeText(cellDisplayValue)
        If String.IsNullOrWhiteSpace(normalized) Then
            Return String.Empty
        End If

        Dim roomStartIndex As Integer = normalized.LastIndexOf(" (", StringComparison.Ordinal)
        If roomStartIndex > 0 AndAlso normalized.EndsWith(")", StringComparison.Ordinal) Then
            Return NormalizeText(normalized.Substring(0, roomStartIndex))
        End If

        Return normalized
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
            Dim candidateSession As String = NormalizeSession(If(rowView("Session"), String.Empty).ToString())
            If Not String.Equals(candidateSession, normalizedSession, StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            For Each column As DataGridColumn In SchedulingTimetableDataGrid.Columns
                If column Is Nothing OrElse column.Header Is Nothing Then
                    Continue For
                End If

                If String.Equals(NormalizeText(column.Header.ToString()), normalizedDay, StringComparison.OrdinalIgnoreCase) Then
                    SchedulingTimetableDataGrid.SelectedItem = rowView
                    SchedulingTimetableDataGrid.CurrentCell = New DataGridCellInfo(rowView, column)
                    SchedulingTimetableDataGrid.ScrollIntoView(rowView, column)
                    Return
                End If
            Next
        Next
    End Sub

    Private Function SelectProfessorById(teacherId As String) As Boolean
        Dim normalizedTeacherId As String = NormalizeText(teacherId)
        If String.IsNullOrWhiteSpace(normalizedTeacherId) Then
            Return False
        End If

        For Each optionEntry As TeacherOption In _teacherOptions
            If optionEntry Is Nothing Then
                Continue For
            End If

            If String.Equals(NormalizeText(optionEntry.TeacherId), normalizedTeacherId, StringComparison.OrdinalIgnoreCase) Then
                ProfessorComboBox.SelectedItem = optionEntry
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

        For Each optionEntry As SubjectOption In _subjectOptions
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

        For Each optionEntry As SubjectOption In _subjectOptions
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
