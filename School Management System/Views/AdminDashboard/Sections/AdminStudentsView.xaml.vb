Imports System.Data
Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json
Imports System.Windows.Media.Imaging
Imports Microsoft.Win32

Class AdminStudentsView
    Private Enum StudentFormMode
        Add
        Edit
    End Enum

    Private Structure StudentFormValues
        Public StudentId As String
        Public FullName As String
        Public YearLevel As String
        Public Course As String
        Public Section As String
        Public PhotoPath As String
    End Structure

    Private Class StudentStorageRecord
        Public Property StudentId As String
        Public Property FullName As String
        Public Property YearLevel As String
        Public Property Course As String
        Public Property Section As String
        Public Property PhotoPath As String
    End Class

    Private _studentsTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As StudentFormMode = StudentFormMode.Add
    Private _editingStudentOriginalId As String = String.Empty
    Private ReadOnly _studentsStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "students.json")
    Private ReadOnly _studentStorageJsonOptions As New JsonSerializerOptions() With {
        .WriteIndented = True
    }

    Public Sub New()
        InitializeComponent()
        LoadStudentsTable()
        HideStudentForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyStudentsFilter()
    End Sub

    Public Sub SetStudentsTable(table As DataTable)
        _studentsTable = NormalizeStudentsTable(table)
        StudentsDataGrid.ItemsSource = _studentsTable.DefaultView
        ApplyStudentsFilter()
        UpdateStudentsCount()
        EnsureSelectedRowForDetails()
        RefreshStudentDetailsPanel()
    End Sub

    Private Sub LoadStudentsTable()
        Dim studentsTable As DataTable = FetchStudentsTable()
        SetStudentsTable(studentsTable)
    End Sub

    Private Function FetchStudentsTable() As DataTable
        Return ReadStudentsFromStorage()
    End Function

    Private Function NormalizeStudentsTable(source As DataTable) As DataTable
        Dim table As DataTable = If(source Is Nothing, CreateEmptyStudentsTable(), source.Copy())

        Dim studentIdColumn As DataColumn = EnsureStudentsColumn(table, "Student ID", "StudentId", "Student_ID", "ID", "Student Number")
        Dim fullNameColumn As DataColumn = EnsureStudentsColumn(table, "Full Name", "FullName", "Name", "Student Name")
        Dim yearLevelColumn As DataColumn = EnsureStudentsColumn(table, "Year Level", "YearLevel", "Year")
        Dim courseColumn As DataColumn = EnsureStudentsColumn(table, "Course", "Program", "Course Name", "Degree")
        Dim sectionColumn As DataColumn = EnsureStudentsColumn(table, "Section", "Class", "Block")
        Dim photoPathColumn As DataColumn = EnsureStudentsColumn(table, "Photo Path", "PhotoPath", "Photo", "Image", "ImagePath", "Avatar")

        RemoveStudentsColumns(table,
                              "Status",
                              "Enrollment Status",
                              "State",
                              "Photo",
                              "Image",
                              "ImagePath",
                              "Avatar")

        studentIdColumn.SetOrdinal(0)
        fullNameColumn.SetOrdinal(1)
        yearLevelColumn.SetOrdinal(2)
        courseColumn.SetOrdinal(3)
        sectionColumn.SetOrdinal(4)
        photoPathColumn.SetOrdinal(5)

        For Each row As DataRow In table.Rows
            If row.IsNull("Student ID") Then
                row("Student ID") = String.Empty
            Else
                row("Student ID") = row("Student ID").ToString().Trim()
            End If

            If row.IsNull("Full Name") Then
                row("Full Name") = String.Empty
            Else
                row("Full Name") = row("Full Name").ToString().Trim()
            End If

            If row.IsNull("Year Level") Then
                row("Year Level") = String.Empty
            Else
                row("Year Level") = row("Year Level").ToString().Trim()
            End If

            If row.IsNull("Course") Then
                row("Course") = String.Empty
            Else
                row("Course") = row("Course").ToString().Trim()
            End If

            If row.IsNull("Section") Then
                row("Section") = String.Empty
            Else
                row("Section") = row("Section").ToString().Trim()
            End If

            If row.IsNull("Photo Path") Then
                row("Photo Path") = String.Empty
            Else
                row("Photo Path") = row("Photo Path").ToString().Trim()
            End If
        Next

        Return table
    End Function

    Private Sub RemoveStudentsColumns(table As DataTable, ParamArray columnNames() As String)
        If table Is Nothing OrElse columnNames Is Nothing Then
            Return
        End If

        For columnIndex As Integer = table.Columns.Count - 1 To 0 Step -1
            Dim candidateColumn As DataColumn = table.Columns(columnIndex)
            For Each requestedName As String In columnNames
                If String.Equals(candidateColumn.ColumnName, requestedName, StringComparison.OrdinalIgnoreCase) Then
                    table.Columns.Remove(candidateColumn)
                    Exit For
                End If
            Next
        Next
    End Sub

    Private Function EnsureStudentsColumn(table As DataTable, targetColumnName As String, ParamArray aliases() As String) As DataColumn
        Dim existingColumn As DataColumn = FindTableColumn(table, targetColumnName)
        If existingColumn Is Nothing Then
            existingColumn = FindTableColumn(table, aliases)
            If existingColumn IsNot Nothing Then
                existingColumn.ColumnName = targetColumnName
            End If
        End If

        If existingColumn Is Nothing Then
            existingColumn = New DataColumn(targetColumnName, GetType(String))
            table.Columns.Add(existingColumn)
        End If

        Return existingColumn
    End Function

    Private Function FindTableColumn(table As DataTable, ParamArray columnNames() As String) As DataColumn
        If table Is Nothing OrElse columnNames Is Nothing Then
            Return Nothing
        End If

        For Each requestedName As String In columnNames
            Dim normalizedRequestedName As String = If(requestedName, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedRequestedName) Then
                Continue For
            End If

            For Each candidateColumn As DataColumn In table.Columns
                If String.Equals(candidateColumn.ColumnName, normalizedRequestedName, StringComparison.OrdinalIgnoreCase) Then
                    Return candidateColumn
                End If
            Next
        Next

        Return Nothing
    End Function

    Private Function CreateEmptyStudentsTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Student ID", GetType(String))
        table.Columns.Add("Full Name", GetType(String))
        table.Columns.Add("Year Level", GetType(String))
        table.Columns.Add("Course", GetType(String))
        table.Columns.Add("Section", GetType(String))
        table.Columns.Add("Photo Path", GetType(String))
        Return table
    End Function

    Private Function ReadStudentsFromStorage() As DataTable
        Dim table As DataTable = CreateEmptyStudentsTable()
        If Not File.Exists(_studentsStoragePath) Then
            Return table
        End If

        Try
            Dim json As String = File.ReadAllText(_studentsStoragePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return table
            End If

            Dim records As List(Of StudentStorageRecord) =
                JsonSerializer.Deserialize(Of List(Of StudentStorageRecord))(json, _studentStorageJsonOptions)
            If records Is Nothing Then
                Return table
            End If

            For Each record As StudentStorageRecord In records
                Dim row As DataRow = table.NewRow()
                row("Student ID") = If(record.StudentId, String.Empty).Trim()
                row("Full Name") = If(record.FullName, String.Empty).Trim()
                row("Year Level") = If(record.YearLevel, String.Empty).Trim()
                row("Course") = If(record.Course, String.Empty).Trim()
                row("Section") = If(record.Section, String.Empty).Trim()
                row("Photo Path") = If(record.PhotoPath, String.Empty).Trim()
                table.Rows.Add(row)
            Next
        Catch ex As Exception
            MessageBox.Show("Unable to load saved students data." & Environment.NewLine & ex.Message,
                            "Students",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
        End Try

        Return table
    End Function

    Private Function PersistStudentsToStorage() As Boolean
        If _studentsTable Is Nothing Then
            Return True
        End If

        Try
            Dim storageDirectory As String = Path.GetDirectoryName(_studentsStoragePath)
            If Not String.IsNullOrWhiteSpace(storageDirectory) Then
                Directory.CreateDirectory(storageDirectory)
            End If

            Dim records As New List(Of StudentStorageRecord)()
            For Each row As DataRow In _studentsTable.Rows
                If row.RowState = DataRowState.Deleted Then
                    Continue For
                End If

                records.Add(New StudentStorageRecord With {
                    .StudentId = ReadRowValue(row, "Student ID"),
                    .FullName = ReadRowValue(row, "Full Name"),
                    .YearLevel = ReadRowValue(row, "Year Level"),
                    .Course = ReadRowValue(row, "Course"),
                    .Section = ReadRowValue(row, "Section"),
                    .PhotoPath = ReadRowValue(row, "Photo Path")
                })
            Next

            Dim json As String = JsonSerializer.Serialize(records, _studentStorageJsonOptions)
            File.WriteAllText(_studentsStoragePath, json)
            Return True
        Catch ex As Exception
            MessageBox.Show("Unable to save students data." & Environment.NewLine & ex.Message,
                            "Students",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            Return False
        End Try
    End Function

    Private Sub ApplyStudentsFilter()
        If _studentsTable Is Nothing Then
            UpdateStudentsCount()
            Return
        End If

        If String.IsNullOrWhiteSpace(_searchTerm) Then
            _studentsTable.DefaultView.RowFilter = String.Empty
        Else
            Dim escapedTerm As String = EscapeLikeValue(_searchTerm)
            _studentsTable.DefaultView.RowFilter =
                "[Student ID] LIKE '*" & escapedTerm & "*' OR " &
                "[Full Name] LIKE '*" & escapedTerm & "*' OR " &
                "[Year Level] LIKE '*" & escapedTerm & "*' OR " &
                "[Course] LIKE '*" & escapedTerm & "*' OR " &
                "[Section] LIKE '*" & escapedTerm & "*'"
        End If

        UpdateStudentsCount()
        EnsureSelectedRowForDetails()
        RefreshStudentDetailsPanel()
    End Sub

    Private Function EscapeLikeValue(value As String) As String
        Dim safeValue As String = If(value, String.Empty)
        Return safeValue.
            Replace("'", "''").
            Replace("[", "[[]").
            Replace("]", "[]]").
            Replace("*", "[*]").
            Replace("%", "[%]")
    End Function

    Private Sub UpdateStudentsCount()
        If StudentsCountTextBlock Is Nothing Then
            Return
        End If

        If _studentsTable Is Nothing Then
            StudentsCountTextBlock.Text = "0 students"
            Return
        End If

        Dim visibleCount As Integer = _studentsTable.DefaultView.Count
        Dim totalCount As Integer = _studentsTable.Rows.Count
        Dim hasSearch As Boolean = Not String.IsNullOrWhiteSpace(_searchTerm)

        If hasSearch AndAlso visibleCount <> totalCount Then
            StudentsCountTextBlock.Text = visibleCount.ToString() & " of " & totalCount.ToString() & " students"
        Else
            StudentsCountTextBlock.Text = totalCount.ToString() & " students"
        End If
    End Sub

    Private Sub OpenAddStudentButton_Click(sender As Object, e As RoutedEventArgs)
        BeginAddStudent()
    End Sub

    Private Sub EditSelectedStudentButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        BeginEditStudent(selectedRow)
    End Sub

    Private Sub DeleteSelectedStudentButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        DeleteStudentRow(selectedRow)
    End Sub

    Private Sub StudentsDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        RefreshStudentDetailsPanel()
    End Sub

    Private Sub BrowseStudentPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dialog As New OpenFileDialog() With {
            .Title = "Select Student Photo",
            .Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp|All Files|*.*",
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        If dialog.ShowDialog() = True Then
            StudentFormPhotoPathTextBox.Text = If(dialog.FileName, String.Empty).Trim()
        End If
    End Sub

    Private Sub ClearStudentPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        StudentFormPhotoPathTextBox.Text = String.Empty
    End Sub

    Private Sub StudentFormPhotoPathTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        UpdateImageControlSource(StudentFormPhotoPreviewImage, If(StudentFormPhotoPathTextBox.Text, String.Empty))
    End Sub

    Private Sub BeginAddStudent()
        OpenStudentForm(StudentFormMode.Add, "Add Student", "Create a new student record.", "Add Student", Nothing)
        StudentFormStudentIdTextBox.Focus()
    End Sub

    Private Sub BeginEditStudent(row As DataRow)
        If row Is Nothing Then
            Return
        End If

        OpenStudentForm(StudentFormMode.Edit, "Edit Student", "Update student details.", "Save Changes", row)
        StudentFormFirstNameTextBox.Focus()
    End Sub

    Private Sub SaveStudentFormButton_Click(sender As Object, e As RoutedEventArgs)
        Dim formValues As New StudentFormValues()
        If Not TryReadStudentForm(formValues) Then
            Return
        End If

        EnsureStudentsTableLoaded()
        Dim targetRow As DataRow = ResolveTargetRowForSave(formValues.StudentId)
        If targetRow Is Nothing Then
            Return
        End If

        WriteStudentValues(targetRow, formValues)
        If _activeFormMode = StudentFormMode.Add Then
            _studentsTable.Rows.Add(targetRow)
        End If

        _studentsTable.AcceptChanges()
        If Not PersistStudentsToStorage() Then
            Return
        End If

        ApplyStudentsFilter()
        SelectStudentById(formValues.StudentId)
        HideStudentForm()
    End Sub

    Private Sub OpenStudentForm(mode As StudentFormMode,
                                title As String,
                                subtitle As String,
                                actionText As String,
                                row As DataRow)
        _activeFormMode = mode
        StudentFormTitleTextBlock.Text = title
        StudentFormSubtitleTextBlock.Text = subtitle
        SaveStudentFormButton.Content = actionText
        StudentFormStudentIdTextBox.IsReadOnly = False

        If row Is Nothing Then
            _editingStudentOriginalId = String.Empty
            ClearStudentFormInputs()
        Else
            _editingStudentOriginalId = ReadRowValue(row, "Student ID")
            PopulateStudentForm(row)
        End If

        ShowStudentForm()
    End Sub

    Private Sub EnsureStudentsTableLoaded()
        If _studentsTable Is Nothing Then
            SetStudentsTable(Nothing)
        End If
    End Sub

    Private Function ResolveTargetRowForSave(studentId As String) As DataRow
        Select Case _activeFormMode
            Case StudentFormMode.Add
                If FindStudentRowById(studentId) IsNot Nothing Then
                    MessageBox.Show("Student ID already exists.", "Duplicate Student ID", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return _studentsTable.NewRow()

            Case StudentFormMode.Edit
                Dim targetRow As DataRow = FindStudentRowById(_editingStudentOriginalId)
                If targetRow Is Nothing Then
                    MessageBox.Show("The selected student no longer exists.", "Edit Student", MessageBoxButton.OK, MessageBoxImage.Information)
                    HideStudentForm()
                    Return Nothing
                End If

                If Not String.Equals(_editingStudentOriginalId, studentId, StringComparison.OrdinalIgnoreCase) AndAlso
                   FindStudentRowById(studentId) IsNot Nothing Then
                    MessageBox.Show("Student ID already exists.", "Duplicate Student ID", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return targetRow
        End Select

        Return Nothing
    End Function

    Private Sub WriteStudentValues(row As DataRow, values As StudentFormValues)
        row("Student ID") = values.StudentId
        row("Full Name") = values.FullName
        row("Year Level") = values.YearLevel
        row("Course") = values.Course
        row("Section") = values.Section
        row("Photo Path") = values.PhotoPath
    End Sub

    Private Sub CancelStudentFormButton_Click(sender As Object, e As RoutedEventArgs)
        HideStudentForm()
    End Sub

    Private Sub DeleteStudentRow(row As DataRow)
        If row Is Nothing OrElse row.RowState = DataRowState.Deleted Then
            Return
        End If

        Dim fullName As String = ReadRowValue(row, "Full Name")
        Dim studentId As String = ReadRowValue(row, "Student ID")
        Dim recordLabel As String = If(String.IsNullOrWhiteSpace(fullName), studentId, fullName)

        Dim confirmation As MessageBoxResult = MessageBox.Show("Delete " & recordLabel & "?", "Delete Student", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        row.Delete()
        _studentsTable.AcceptChanges()
        If Not PersistStudentsToStorage() Then
            Return
        End If

        ApplyStudentsFilter()

        If _activeFormMode = StudentFormMode.Edit AndAlso
           String.Equals(_editingStudentOriginalId, studentId, StringComparison.OrdinalIgnoreCase) Then
            HideStudentForm()
        End If

        EnsureSelectedRowForDetails()
        RefreshStudentDetailsPanel()
    End Sub

    Private Sub EnsureSelectedRowForDetails()
        If StudentsDataGrid Is Nothing OrElse _studentsTable Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow IsNot Nothing Then
            Return
        End If

        If _studentsTable.DefaultView.Count > 0 Then
            StudentsDataGrid.SelectedItem = _studentsTable.DefaultView(0)
        End If
    End Sub

    Private Function TryGetSelectedGridRow() As DataRow
        Dim selectedRowView As DataRowView = TryCast(StudentsDataGrid.SelectedItem, DataRowView)
        If selectedRowView Is Nothing OrElse selectedRowView.Row Is Nothing Then
            Return Nothing
        End If

        If selectedRowView.Row.RowState = DataRowState.Deleted Then
            Return Nothing
        End If

        Return selectedRowView.Row
    End Function

    Private Sub RefreshStudentDetailsPanel()
        If StudentDetailsStudentIdTextBlock Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        Dim hasSelection As Boolean = selectedRow IsNot Nothing

        If EditSelectedStudentButton IsNot Nothing Then
            EditSelectedStudentButton.IsEnabled = hasSelection
        End If

        If DeleteSelectedStudentButton IsNot Nothing Then
            DeleteSelectedStudentButton.IsEnabled = hasSelection
        End If

        If Not hasSelection Then
            If StudentDetailsSubtitleTextBlock IsNot Nothing Then
                StudentDetailsSubtitleTextBlock.Text = "Select a student from the table."
            End If

            SetDetailsValue(StudentDetailsStudentIdTextBlock, String.Empty)
            SetDetailsValue(StudentDetailsFullNameTextBlock, String.Empty)
            SetDetailsValue(StudentDetailsYearLevelTextBlock, String.Empty)
            SetDetailsValue(StudentDetailsCourseTextBlock, String.Empty)
            SetDetailsValue(StudentDetailsSectionTextBlock, String.Empty)
            UpdateImageControlSource(StudentDetailsPhotoImage, String.Empty)
            Return
        End If

        If StudentDetailsSubtitleTextBlock IsNot Nothing Then
            StudentDetailsSubtitleTextBlock.Text = "Selected student record."
        End If

        SetDetailsValue(StudentDetailsStudentIdTextBlock, ReadRowValue(selectedRow, "Student ID"))
        SetDetailsValue(StudentDetailsFullNameTextBlock, ReadRowValue(selectedRow, "Full Name"))
        SetDetailsValue(StudentDetailsYearLevelTextBlock, ReadRowValue(selectedRow, "Year Level"))
        SetDetailsValue(StudentDetailsCourseTextBlock, ReadRowValue(selectedRow, "Course"))
        SetDetailsValue(StudentDetailsSectionTextBlock, ReadRowValue(selectedRow, "Section"))
        UpdateImageControlSource(StudentDetailsPhotoImage, ReadRowValue(selectedRow, "Photo Path"))
    End Sub

    Private Sub SetDetailsValue(target As TextBlock, value As String)
        If target Is Nothing Then
            Return
        End If

        Dim normalizedValue As String = If(value, String.Empty).Trim()
        target.Text = If(String.IsNullOrWhiteSpace(normalizedValue), "--", normalizedValue)
    End Sub

    Private Sub ShowStudentForm()
        If StudentFormPanel IsNot Nothing Then
            StudentFormPanel.Visibility = Visibility.Visible
        End If

        If StudentDetailsPanel IsNot Nothing Then
            StudentDetailsPanel.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub HideStudentForm()
        _activeFormMode = StudentFormMode.Add
        _editingStudentOriginalId = String.Empty
        StudentFormTitleTextBlock.Text = "Add Student"
        StudentFormSubtitleTextBlock.Text = "Create a new student record."
        SaveStudentFormButton.Content = "Add Student"
        StudentFormStudentIdTextBox.IsReadOnly = False
        ClearStudentFormInputs()

        If StudentFormPanel IsNot Nothing Then
            StudentFormPanel.Visibility = Visibility.Collapsed
        End If

        If StudentDetailsPanel IsNot Nothing Then
            StudentDetailsPanel.Visibility = Visibility.Visible
        End If

        RefreshStudentDetailsPanel()
    End Sub

    Private Sub PopulateStudentForm(row As DataRow)
        StudentFormStudentIdTextBox.Text = ReadRowValue(row, "Student ID")

        Dim fullName As String = ReadRowValue(row, "Full Name")
        Dim firstName As String = String.Empty
        Dim lastName As String = String.Empty
        Dim middleName As String = String.Empty

        SplitFullName(fullName, firstName, lastName, middleName)

        StudentFormFirstNameTextBox.Text = firstName
        StudentFormLastNameTextBox.Text = lastName
        StudentFormMiddleNameTextBox.Text = middleName
        SetComboBoxValue(StudentFormYearLevelComboBox, ReadRowValue(row, "Year Level"))
        StudentFormCourseTextBox.Text = ReadRowValue(row, "Course")
        StudentFormSectionTextBox.Text = ReadRowValue(row, "Section")
        StudentFormPhotoPathTextBox.Text = ReadRowValue(row, "Photo Path")
        UpdateImageControlSource(StudentFormPhotoPreviewImage, StudentFormPhotoPathTextBox.Text)
    End Sub

    Private Sub ClearStudentFormInputs()
        StudentFormStudentIdTextBox.Text = String.Empty
        StudentFormFirstNameTextBox.Text = String.Empty
        StudentFormLastNameTextBox.Text = String.Empty
        StudentFormMiddleNameTextBox.Text = String.Empty
        StudentFormYearLevelComboBox.SelectedIndex = -1
        StudentFormCourseTextBox.Text = String.Empty
        StudentFormSectionTextBox.Text = String.Empty
        StudentFormPhotoPathTextBox.Text = String.Empty
        UpdateImageControlSource(StudentFormPhotoPreviewImage, String.Empty)
    End Sub

    Private Function TryReadStudentForm(ByRef values As StudentFormValues) As Boolean
        values.StudentId = If(StudentFormStudentIdTextBox.Text, String.Empty).Trim()

        Dim firstName As String = If(StudentFormFirstNameTextBox.Text, String.Empty).Trim()
        Dim lastName As String = If(StudentFormLastNameTextBox.Text, String.Empty).Trim()
        Dim middleName As String = If(StudentFormMiddleNameTextBox.Text, String.Empty).Trim()

        values.FullName = BuildFullName(firstName, lastName, middleName)
        values.YearLevel = ReadComboBoxValue(StudentFormYearLevelComboBox)
        values.Course = If(StudentFormCourseTextBox.Text, String.Empty).Trim()
        values.Section = If(StudentFormSectionTextBox.Text, String.Empty).Trim()
        values.PhotoPath = If(StudentFormPhotoPathTextBox.Text, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(values.StudentId) Then
            MessageBox.Show("Student ID is required.", "Student Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(firstName) Then
            MessageBox.Show("First Name is required.", "Student Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(lastName) Then
            MessageBox.Show("Last Name is required.", "Student Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        Return True
    End Function

    Private Function BuildFullName(firstName As String, lastName As String, middleName As String) As String
        Dim normalizedFirst As String = If(firstName, String.Empty).Trim()
        Dim normalizedLast As String = If(lastName, String.Empty).Trim()
        Dim normalizedMiddle As String = If(middleName, String.Empty).Trim()

        Dim parts As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(normalizedFirst) Then
            parts.Add(normalizedFirst)
        End If

        If Not String.IsNullOrWhiteSpace(normalizedLast) Then
            parts.Add(normalizedLast)
        End If

        If Not String.IsNullOrWhiteSpace(normalizedMiddle) Then
            parts.Add(normalizedMiddle)
        End If

        Return String.Join(" ", parts)
    End Function

    Private Sub SplitFullName(fullName As String, ByRef firstName As String, ByRef lastName As String, ByRef middleName As String)
        firstName = String.Empty
        lastName = String.Empty
        middleName = String.Empty

        Dim normalizedFullName As String = If(fullName, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedFullName) Then
            Return
        End If

        If normalizedFullName.Contains(",") Then
            Dim commaParts() As String = normalizedFullName.Split(New Char() {","c}, 2, StringSplitOptions.None)
            lastName = If(commaParts(0), String.Empty).Trim()

            Dim rightSide As String = String.Empty
            If commaParts.Length > 1 Then
                rightSide = If(commaParts(1), String.Empty).Trim()
            End If

            If Not String.IsNullOrWhiteSpace(rightSide) Then
                Dim rightTokens() As String = rightSide.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                If rightTokens.Length > 0 Then
                    firstName = rightTokens(0)
                End If

                If rightTokens.Length > 1 Then
                    middleName = String.Join(" ", rightTokens, 1, rightTokens.Length - 1)
                End If
            End If

            Return
        End If

        Dim tokens() As String = normalizedFullName.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
        If tokens.Length = 1 Then
            firstName = tokens(0)
            Return
        End If

        If tokens.Length = 2 Then
            firstName = tokens(0)
            lastName = tokens(1)
            Return
        End If

        firstName = tokens(0)
        lastName = tokens(1)
        middleName = String.Join(" ", tokens, 2, tokens.Length - 2)
    End Sub

    Private Function ReadComboBoxValue(comboBox As ComboBox) As String
        If comboBox Is Nothing Then
            Return String.Empty
        End If

        Dim selectedItem As ComboBoxItem = TryCast(comboBox.SelectedItem, ComboBoxItem)
        If selectedItem IsNot Nothing Then
            Return If(selectedItem.Content, String.Empty).ToString().Trim()
        End If

        Return If(comboBox.Text, String.Empty).Trim()
    End Function

    Private Sub SetComboBoxValue(comboBox As ComboBox, value As String)
        If comboBox Is Nothing Then
            Return
        End If

        Dim requestedValue As String = If(value, String.Empty).Trim()
        comboBox.SelectedIndex = -1

        For Each rawItem As Object In comboBox.Items
            Dim comboItem As ComboBoxItem = TryCast(rawItem, ComboBoxItem)
            If comboItem Is Nothing Then
                Continue For
            End If

            Dim itemValue As String = If(comboItem.Content, String.Empty).ToString().Trim()
            If String.Equals(itemValue, requestedValue, StringComparison.OrdinalIgnoreCase) Then
                comboBox.SelectedItem = comboItem
                Return
            End If
        Next
    End Sub

    Private Function FindStudentRowById(studentId As String) As DataRow
        If _studentsTable Is Nothing Then
            Return Nothing
        End If

        Dim normalizedStudentId As String = If(studentId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedStudentId) Then
            Return Nothing
        End If

        For Each row As DataRow In _studentsTable.Rows
            If row.RowState = DataRowState.Deleted Then
                Continue For
            End If

            Dim candidateId As String = ReadRowValue(row, "Student ID")
            If String.Equals(candidateId, normalizedStudentId, StringComparison.OrdinalIgnoreCase) Then
                Return row
            End If
        Next

        Return Nothing
    End Function

    Private Sub SelectStudentById(studentId As String)
        If _studentsTable Is Nothing Then
            Return
        End If

        Dim normalizedStudentId As String = If(studentId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedStudentId) Then
            Return
        End If

        For Each rowView As DataRowView In _studentsTable.DefaultView
            Dim candidateId As String = If(rowView("Student ID"), String.Empty).ToString().Trim()
            If String.Equals(candidateId, normalizedStudentId, StringComparison.OrdinalIgnoreCase) Then
                StudentsDataGrid.SelectedItem = rowView
                StudentsDataGrid.ScrollIntoView(rowView)
                RefreshStudentDetailsPanel()
                Return
            End If
        Next

        RefreshStudentDetailsPanel()
    End Sub

    Private Function ReadRowValue(row As DataRow, columnName As String) As String
        If row Is Nothing OrElse row.Table Is Nothing OrElse Not row.Table.Columns.Contains(columnName) Then
            Return String.Empty
        End If

        If row.IsNull(columnName) Then
            Return String.Empty
        End If

        Return row(columnName).ToString().Trim()
    End Function

    Private Sub UpdateImageControlSource(targetImage As System.Windows.Controls.Image, imagePath As String)
        If targetImage Is Nothing Then
            Return
        End If

        Dim normalizedPath As String = If(imagePath, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedPath) OrElse Not File.Exists(normalizedPath) Then
            targetImage.Source = Nothing
            Return
        End If

        Try
            Dim bitmap As New BitmapImage()
            bitmap.BeginInit()
            bitmap.CacheOption = BitmapCacheOption.OnLoad
            bitmap.UriSource = New Uri(normalizedPath, UriKind.Absolute)
            bitmap.EndInit()
            bitmap.Freeze()
            targetImage.Source = bitmap
        Catch
            targetImage.Source = Nothing
        End Try
    End Sub
End Class
