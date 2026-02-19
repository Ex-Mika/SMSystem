Imports System.Data
Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json
Imports System.Windows.Media.Imaging
Imports Microsoft.Win32

Class AdminTeachersView
    Private Enum TeacherFormMode
        Add
        Edit
    End Enum

    Private Structure TeacherFormValues
        Public TeacherId As String
        Public FullName As String
        Public Department As String
        Public Advisory As String
        Public PhotoPath As String
    End Structure

    Private Class TeacherStorageRecord
        Public Property TeacherId As String
        Public Property FullName As String
        Public Property Department As String
        Public Property Advisory As String
        Public Property PhotoPath As String
    End Class

    Private _teachersTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As TeacherFormMode = TeacherFormMode.Add
    Private _editingTeacherOriginalId As String = String.Empty
    Private ReadOnly _teachersStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "teachers.json")
    Private ReadOnly _teacherStorageJsonOptions As New JsonSerializerOptions() With {
        .WriteIndented = True
    }

    Public Sub New()
        InitializeComponent()
        LoadTeachersTable()
        HideTeacherForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyTeachersFilter()
    End Sub

    Public Sub SetTeachersTable(table As DataTable)
        _teachersTable = NormalizeTeachersTable(table)
        TeachersDataGrid.ItemsSource = _teachersTable.DefaultView
        ApplyTeachersFilter()
        UpdateTeachersCount()
        EnsureSelectedTeacherForDetails()
        RefreshTeacherDetailsPanel()
    End Sub

    Private Sub LoadTeachersTable()
        Dim teachersTable As DataTable = FetchTeachersTable()
        SetTeachersTable(teachersTable)
    End Sub

    Private Function FetchTeachersTable() As DataTable
        Return ReadTeachersFromStorage()
    End Function

    Private Function NormalizeTeachersTable(source As DataTable) As DataTable
        Dim table As DataTable = If(source Is Nothing, CreateEmptyTeachersTable(), source.Copy())

        Dim teacherIdColumn As DataColumn = EnsureTeachersColumn(table, "Teacher ID", "TeacherId", "Teacher_ID", "ID", "Teacher Number")
        Dim fullNameColumn As DataColumn = EnsureTeachersColumn(table, "Full Name", "FullName", "Name", "Teacher Name")
        Dim departmentColumn As DataColumn = EnsureTeachersColumn(table, "Department", "Program", "Department Name")
        Dim advisoryColumn As DataColumn = EnsureTeachersColumn(table, "Advisory", "Class", "Section", "Block")
        Dim photoPathColumn As DataColumn = EnsureTeachersColumn(table, "Photo Path", "PhotoPath", "Photo", "Image", "ImagePath", "Avatar")

        RemoveTeachersColumns(table,
                              "Status",
                              "Employment Status",
                              "State",
                              "Enrollment Status",
                              "Photo",
                              "Image",
                              "ImagePath",
                              "Avatar")

        teacherIdColumn.SetOrdinal(0)
        fullNameColumn.SetOrdinal(1)
        departmentColumn.SetOrdinal(2)
        advisoryColumn.SetOrdinal(3)
        photoPathColumn.SetOrdinal(4)

        For Each row As DataRow In table.Rows
            If row.IsNull("Teacher ID") Then
                row("Teacher ID") = String.Empty
            Else
                row("Teacher ID") = row("Teacher ID").ToString().Trim()
            End If

            If row.IsNull("Full Name") Then
                row("Full Name") = String.Empty
            Else
                row("Full Name") = row("Full Name").ToString().Trim()
            End If

            If row.IsNull("Department") Then
                row("Department") = String.Empty
            Else
                row("Department") = row("Department").ToString().Trim()
            End If

            If row.IsNull("Advisory") Then
                row("Advisory") = String.Empty
            Else
                row("Advisory") = row("Advisory").ToString().Trim()
            End If

            If row.IsNull("Photo Path") Then
                row("Photo Path") = String.Empty
            Else
                row("Photo Path") = row("Photo Path").ToString().Trim()
            End If
        Next

        Return table
    End Function

    Private Sub RemoveTeachersColumns(table As DataTable, ParamArray columnNames() As String)
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

    Private Function EnsureTeachersColumn(table As DataTable, targetColumnName As String, ParamArray aliases() As String) As DataColumn
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

    Private Function CreateEmptyTeachersTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Teacher ID", GetType(String))
        table.Columns.Add("Full Name", GetType(String))
        table.Columns.Add("Department", GetType(String))
        table.Columns.Add("Advisory", GetType(String))
        table.Columns.Add("Photo Path", GetType(String))
        Return table
    End Function

    Private Function ReadTeachersFromStorage() As DataTable
        Dim table As DataTable = CreateEmptyTeachersTable()
        If Not File.Exists(_teachersStoragePath) Then
            Return table
        End If

        Try
            Dim json As String = File.ReadAllText(_teachersStoragePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return table
            End If

            Dim records As List(Of TeacherStorageRecord) =
                JsonSerializer.Deserialize(Of List(Of TeacherStorageRecord))(json, _teacherStorageJsonOptions)
            If records Is Nothing Then
                Return table
            End If

            For Each record As TeacherStorageRecord In records
                Dim row As DataRow = table.NewRow()
                row("Teacher ID") = If(record.TeacherId, String.Empty).Trim()
                row("Full Name") = If(record.FullName, String.Empty).Trim()
                row("Department") = If(record.Department, String.Empty).Trim()
                row("Advisory") = If(record.Advisory, String.Empty).Trim()
                row("Photo Path") = If(record.PhotoPath, String.Empty).Trim()
                table.Rows.Add(row)
            Next
        Catch ex As Exception
            MessageBox.Show("Unable to load saved teachers data." & Environment.NewLine & ex.Message,
                            "Teachers",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
        End Try

        Return table
    End Function

    Private Function PersistTeachersToStorage() As Boolean
        If _teachersTable Is Nothing Then
            Return True
        End If

        Try
            Dim storageDirectory As String = Path.GetDirectoryName(_teachersStoragePath)
            If Not String.IsNullOrWhiteSpace(storageDirectory) Then
                Directory.CreateDirectory(storageDirectory)
            End If

            Dim records As New List(Of TeacherStorageRecord)()
            For Each row As DataRow In _teachersTable.Rows
                If row.RowState = DataRowState.Deleted Then
                    Continue For
                End If

                records.Add(New TeacherStorageRecord With {
                    .TeacherId = ReadRowValue(row, "Teacher ID"),
                    .FullName = ReadRowValue(row, "Full Name"),
                    .Department = ReadRowValue(row, "Department"),
                    .Advisory = ReadRowValue(row, "Advisory"),
                    .PhotoPath = ReadRowValue(row, "Photo Path")
                })
            Next

            Dim json As String = JsonSerializer.Serialize(records, _teacherStorageJsonOptions)
            File.WriteAllText(_teachersStoragePath, json)
            Return True
        Catch ex As Exception
            MessageBox.Show("Unable to save teachers data." & Environment.NewLine & ex.Message,
                            "Teachers",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            Return False
        End Try
    End Function

    Private Sub ApplyTeachersFilter()
        If _teachersTable Is Nothing Then
            UpdateTeachersCount()
            Return
        End If

        If String.IsNullOrWhiteSpace(_searchTerm) Then
            _teachersTable.DefaultView.RowFilter = String.Empty
        Else
            Dim escapedTerm As String = EscapeLikeValue(_searchTerm)
            _teachersTable.DefaultView.RowFilter =
                "[Teacher ID] LIKE '*" & escapedTerm & "*' OR " &
                "[Full Name] LIKE '*" & escapedTerm & "*' OR " &
                "[Department] LIKE '*" & escapedTerm & "*' OR " &
                "[Advisory] LIKE '*" & escapedTerm & "*'"
        End If

        UpdateTeachersCount()
        EnsureSelectedTeacherForDetails()
        RefreshTeacherDetailsPanel()
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

    Private Sub UpdateTeachersCount()
        If TeachersCountTextBlock Is Nothing Then
            Return
        End If

        If _teachersTable Is Nothing Then
            TeachersCountTextBlock.Text = "0 teachers"
            Return
        End If

        Dim visibleCount As Integer = _teachersTable.DefaultView.Count
        Dim totalCount As Integer = _teachersTable.Rows.Count
        Dim hasSearch As Boolean = Not String.IsNullOrWhiteSpace(_searchTerm)

        If hasSearch AndAlso visibleCount <> totalCount Then
            TeachersCountTextBlock.Text = visibleCount.ToString() & " of " & totalCount.ToString() & " teachers"
        Else
            TeachersCountTextBlock.Text = totalCount.ToString() & " teachers"
        End If
    End Sub

    Private Sub OpenAddTeacherButton_Click(sender As Object, e As RoutedEventArgs)
        BeginAddTeacher()
    End Sub

    Private Sub EditSelectedTeacherButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        BeginEditTeacher(selectedRow)
    End Sub

    Private Sub DeleteSelectedTeacherButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        DeleteTeacherRow(selectedRow)
    End Sub

    Private Sub TeachersDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        RefreshTeacherDetailsPanel()
    End Sub

    Private Sub BrowseTeacherPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dialog As New OpenFileDialog() With {
            .Title = "Select Teacher Photo",
            .Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp|All Files|*.*",
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        If dialog.ShowDialog() = True Then
            TeacherFormPhotoPathTextBox.Text = If(dialog.FileName, String.Empty).Trim()
        End If
    End Sub

    Private Sub ClearTeacherPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        TeacherFormPhotoPathTextBox.Text = String.Empty
    End Sub

    Private Sub TeacherFormPhotoPathTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        UpdateImageControlSource(TeacherFormPhotoPreviewImage, If(TeacherFormPhotoPathTextBox.Text, String.Empty))
    End Sub

    Private Sub BeginAddTeacher()
        OpenTeacherForm(TeacherFormMode.Add, "Add Teacher", "Create a new teacher record.", "Add Teacher", Nothing)
        TeacherFormTeacherIdTextBox.Focus()
    End Sub

    Private Sub BeginEditTeacher(row As DataRow)
        If row Is Nothing Then
            Return
        End If

        OpenTeacherForm(TeacherFormMode.Edit, "Edit Teacher", "Update teacher details.", "Save Changes", row)
        TeacherFormFirstNameTextBox.Focus()
    End Sub

    Private Sub SaveTeacherFormButton_Click(sender As Object, e As RoutedEventArgs)
        Dim formValues As New TeacherFormValues()
        If Not TryReadTeacherForm(formValues) Then
            Return
        End If

        EnsureTeachersTableLoaded()
        Dim targetRow As DataRow = ResolveTargetRowForSave(formValues.TeacherId)
        If targetRow Is Nothing Then
            Return
        End If

        WriteTeacherValues(targetRow, formValues)
        If _activeFormMode = TeacherFormMode.Add Then
            _teachersTable.Rows.Add(targetRow)
        End If

        _teachersTable.AcceptChanges()
        If Not PersistTeachersToStorage() Then
            Return
        End If

        ApplyTeachersFilter()
        SelectTeacherById(formValues.TeacherId)
        HideTeacherForm()
    End Sub

    Private Sub OpenTeacherForm(mode As TeacherFormMode,
                                title As String,
                                subtitle As String,
                                actionText As String,
                                row As DataRow)
        _activeFormMode = mode
        TeacherFormTitleTextBlock.Text = title
        TeacherFormSubtitleTextBlock.Text = subtitle
        SaveTeacherFormButton.Content = actionText
        TeacherFormTeacherIdTextBox.IsReadOnly = False

        If row Is Nothing Then
            _editingTeacherOriginalId = String.Empty
            ClearTeacherFormInputs()
        Else
            _editingTeacherOriginalId = ReadRowValue(row, "Teacher ID")
            PopulateTeacherForm(row)
        End If

        ShowTeacherForm()
    End Sub

    Private Sub EnsureTeachersTableLoaded()
        If _teachersTable Is Nothing Then
            SetTeachersTable(Nothing)
        End If
    End Sub

    Private Function ResolveTargetRowForSave(teacherId As String) As DataRow
        Select Case _activeFormMode
            Case TeacherFormMode.Add
                If FindTeacherRowById(teacherId) IsNot Nothing Then
                    MessageBox.Show("Teacher ID already exists.", "Duplicate Teacher ID", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return _teachersTable.NewRow()

            Case TeacherFormMode.Edit
                Dim targetRow As DataRow = FindTeacherRowById(_editingTeacherOriginalId)
                If targetRow Is Nothing Then
                    MessageBox.Show("The selected teacher no longer exists.", "Edit Teacher", MessageBoxButton.OK, MessageBoxImage.Information)
                    HideTeacherForm()
                    Return Nothing
                End If

                If Not String.Equals(_editingTeacherOriginalId, teacherId, StringComparison.OrdinalIgnoreCase) AndAlso
                   FindTeacherRowById(teacherId) IsNot Nothing Then
                    MessageBox.Show("Teacher ID already exists.", "Duplicate Teacher ID", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return targetRow
        End Select

        Return Nothing
    End Function

    Private Sub WriteTeacherValues(row As DataRow, values As TeacherFormValues)
        row("Teacher ID") = values.TeacherId
        row("Full Name") = values.FullName
        row("Department") = values.Department
        row("Advisory") = values.Advisory
        row("Photo Path") = values.PhotoPath
    End Sub

    Private Sub CancelTeacherFormButton_Click(sender As Object, e As RoutedEventArgs)
        HideTeacherForm()
    End Sub

    Private Sub DeleteTeacherRow(row As DataRow)
        If row Is Nothing OrElse row.RowState = DataRowState.Deleted Then
            Return
        End If

        Dim fullName As String = ReadRowValue(row, "Full Name")
        Dim teacherId As String = ReadRowValue(row, "Teacher ID")
        Dim recordLabel As String = If(String.IsNullOrWhiteSpace(fullName), teacherId, fullName)

        Dim confirmation As MessageBoxResult = MessageBox.Show("Delete " & recordLabel & "?", "Delete Teacher", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        row.Delete()
        _teachersTable.AcceptChanges()
        If Not PersistTeachersToStorage() Then
            Return
        End If

        ApplyTeachersFilter()

        If _activeFormMode = TeacherFormMode.Edit AndAlso
           String.Equals(_editingTeacherOriginalId, teacherId, StringComparison.OrdinalIgnoreCase) Then
            HideTeacherForm()
        End If

        EnsureSelectedTeacherForDetails()
        RefreshTeacherDetailsPanel()
    End Sub

    Private Sub EnsureSelectedTeacherForDetails()
        If TeachersDataGrid Is Nothing OrElse _teachersTable Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow IsNot Nothing Then
            Return
        End If

        If _teachersTable.DefaultView.Count > 0 Then
            TeachersDataGrid.SelectedItem = _teachersTable.DefaultView(0)
        End If
    End Sub

    Private Function TryGetSelectedGridRow() As DataRow
        Dim selectedRowView As DataRowView = TryCast(TeachersDataGrid.SelectedItem, DataRowView)
        If selectedRowView Is Nothing OrElse selectedRowView.Row Is Nothing Then
            Return Nothing
        End If

        If selectedRowView.Row.RowState = DataRowState.Deleted Then
            Return Nothing
        End If

        Return selectedRowView.Row
    End Function

    Private Sub RefreshTeacherDetailsPanel()
        If TeacherDetailsTeacherIdTextBlock Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        Dim hasSelection As Boolean = selectedRow IsNot Nothing

        If EditSelectedTeacherButton IsNot Nothing Then
            EditSelectedTeacherButton.IsEnabled = hasSelection
        End If

        If DeleteSelectedTeacherButton IsNot Nothing Then
            DeleteSelectedTeacherButton.IsEnabled = hasSelection
        End If

        If Not hasSelection Then
            If TeacherDetailsSubtitleTextBlock IsNot Nothing Then
                TeacherDetailsSubtitleTextBlock.Text = "Select a teacher from the table."
            End If

            SetDetailsValue(TeacherDetailsTeacherIdTextBlock, String.Empty)
            SetDetailsValue(TeacherDetailsFullNameTextBlock, String.Empty)
            SetDetailsValue(TeacherDetailsDepartmentTextBlock, String.Empty)
            SetDetailsValue(TeacherDetailsAdvisoryTextBlock, String.Empty)
            UpdateImageControlSource(TeacherDetailsPhotoImage, String.Empty)
            Return
        End If

        If TeacherDetailsSubtitleTextBlock IsNot Nothing Then
            TeacherDetailsSubtitleTextBlock.Text = "Selected teacher record."
        End If

        SetDetailsValue(TeacherDetailsTeacherIdTextBlock, ReadRowValue(selectedRow, "Teacher ID"))
        SetDetailsValue(TeacherDetailsFullNameTextBlock, ReadRowValue(selectedRow, "Full Name"))
        SetDetailsValue(TeacherDetailsDepartmentTextBlock, ReadRowValue(selectedRow, "Department"))
        SetDetailsValue(TeacherDetailsAdvisoryTextBlock, ReadRowValue(selectedRow, "Advisory"))
        UpdateImageControlSource(TeacherDetailsPhotoImage, ReadRowValue(selectedRow, "Photo Path"))
    End Sub

    Private Sub SetDetailsValue(target As TextBlock, value As String)
        If target Is Nothing Then
            Return
        End If

        Dim normalizedValue As String = If(value, String.Empty).Trim()
        target.Text = If(String.IsNullOrWhiteSpace(normalizedValue), "--", normalizedValue)
    End Sub

    Private Sub ShowTeacherForm()
        If TeacherFormPanel IsNot Nothing Then
            TeacherFormPanel.Visibility = Visibility.Visible
        End If

        If TeacherDetailsPanel IsNot Nothing Then
            TeacherDetailsPanel.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub HideTeacherForm()
        _activeFormMode = TeacherFormMode.Add
        _editingTeacherOriginalId = String.Empty
        TeacherFormTitleTextBlock.Text = "Add Teacher"
        TeacherFormSubtitleTextBlock.Text = "Create a new teacher record."
        SaveTeacherFormButton.Content = "Add Teacher"
        TeacherFormTeacherIdTextBox.IsReadOnly = False
        ClearTeacherFormInputs()

        If TeacherFormPanel IsNot Nothing Then
            TeacherFormPanel.Visibility = Visibility.Collapsed
        End If

        If TeacherDetailsPanel IsNot Nothing Then
            TeacherDetailsPanel.Visibility = Visibility.Visible
        End If

        RefreshTeacherDetailsPanel()
    End Sub

    Private Sub PopulateTeacherForm(row As DataRow)
        TeacherFormTeacherIdTextBox.Text = ReadRowValue(row, "Teacher ID")

        Dim fullName As String = ReadRowValue(row, "Full Name")
        Dim firstName As String = String.Empty
        Dim lastName As String = String.Empty
        Dim middleName As String = String.Empty

        SplitFullName(fullName, firstName, lastName, middleName)

        TeacherFormFirstNameTextBox.Text = firstName
        TeacherFormLastNameTextBox.Text = lastName
        TeacherFormMiddleNameTextBox.Text = middleName
        TeacherFormDepartmentTextBox.Text = ReadRowValue(row, "Department")
        TeacherFormAdvisoryTextBox.Text = ReadRowValue(row, "Advisory")
        TeacherFormPhotoPathTextBox.Text = ReadRowValue(row, "Photo Path")
        UpdateImageControlSource(TeacherFormPhotoPreviewImage, TeacherFormPhotoPathTextBox.Text)
    End Sub

    Private Sub ClearTeacherFormInputs()
        TeacherFormTeacherIdTextBox.Text = String.Empty
        TeacherFormFirstNameTextBox.Text = String.Empty
        TeacherFormLastNameTextBox.Text = String.Empty
        TeacherFormMiddleNameTextBox.Text = String.Empty
        TeacherFormDepartmentTextBox.Text = String.Empty
        TeacherFormAdvisoryTextBox.Text = String.Empty
        TeacherFormPhotoPathTextBox.Text = String.Empty
        UpdateImageControlSource(TeacherFormPhotoPreviewImage, String.Empty)
    End Sub

    Private Function TryReadTeacherForm(ByRef values As TeacherFormValues) As Boolean
        values.TeacherId = If(TeacherFormTeacherIdTextBox.Text, String.Empty).Trim()

        Dim firstName As String = If(TeacherFormFirstNameTextBox.Text, String.Empty).Trim()
        Dim lastName As String = If(TeacherFormLastNameTextBox.Text, String.Empty).Trim()
        Dim middleName As String = If(TeacherFormMiddleNameTextBox.Text, String.Empty).Trim()

        values.FullName = BuildFullName(firstName, lastName, middleName)
        values.Department = If(TeacherFormDepartmentTextBox.Text, String.Empty).Trim()
        values.Advisory = If(TeacherFormAdvisoryTextBox.Text, String.Empty).Trim()
        values.PhotoPath = If(TeacherFormPhotoPathTextBox.Text, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(values.TeacherId) Then
            MessageBox.Show("Teacher ID is required.", "Teacher Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(firstName) Then
            MessageBox.Show("First Name is required.", "Teacher Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(lastName) Then
            MessageBox.Show("Last Name is required.", "Teacher Form", MessageBoxButton.OK, MessageBoxImage.Information)
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

    Private Function FindTeacherRowById(teacherId As String) As DataRow
        If _teachersTable Is Nothing Then
            Return Nothing
        End If

        Dim normalizedTeacherId As String = If(teacherId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedTeacherId) Then
            Return Nothing
        End If

        For Each row As DataRow In _teachersTable.Rows
            If row.RowState = DataRowState.Deleted Then
                Continue For
            End If

            Dim candidateId As String = ReadRowValue(row, "Teacher ID")
            If String.Equals(candidateId, normalizedTeacherId, StringComparison.OrdinalIgnoreCase) Then
                Return row
            End If
        Next

        Return Nothing
    End Function

    Private Sub SelectTeacherById(teacherId As String)
        If _teachersTable Is Nothing Then
            Return
        End If

        Dim normalizedTeacherId As String = If(teacherId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedTeacherId) Then
            Return
        End If

        For Each rowView As DataRowView In _teachersTable.DefaultView
            Dim candidateId As String = If(rowView("Teacher ID"), String.Empty).ToString().Trim()
            If String.Equals(candidateId, normalizedTeacherId, StringComparison.OrdinalIgnoreCase) Then
                TeachersDataGrid.SelectedItem = rowView
                TeachersDataGrid.ScrollIntoView(rowView)
                RefreshTeacherDetailsPanel()
                Return
            End If
        Next

        RefreshTeacherDetailsPanel()
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
