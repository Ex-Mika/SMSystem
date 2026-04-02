Imports System.Data
Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json

Class AdminCoursesView
    Private Enum CourseFormMode
        Add
        Edit
    End Enum

    Private Structure CourseFormValues
        Public CourseCode As String
        Public CourseTitle As String
        Public Department As String
        Public Units As String
    End Structure

    Private Class CourseStorageRecord
        Public Property CourseCode As String
        Public Property CourseTitle As String
        Public Property Department As String
        Public Property Units As String
    End Class

    Private _coursesTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As CourseFormMode = CourseFormMode.Add
    Private _editingCourseOriginalCode As String = String.Empty
    Private ReadOnly _coursesStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "courses.json")
    Private ReadOnly _courseStorageJsonOptions As New JsonSerializerOptions() With {
        .WriteIndented = True
    }

    Public Sub New()
        InitializeComponent()
        LoadCoursesTable()
        HideCourseForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyCoursesFilter()
    End Sub

    Public Sub RefreshData()
        LoadCoursesTable(GetSelectedCourseCode())
    End Sub

    Public Sub SetCoursesTable(table As DataTable)
        _coursesTable = NormalizeCoursesTable(table)
        CoursesDataGrid.ItemsSource = _coursesTable.DefaultView
        ApplyCoursesFilter()
        UpdateCoursesCount()
        EnsureSelectedCourseForDetails()
        RefreshCourseDetailsPanel()
    End Sub

    Private Sub LoadCoursesTable(Optional courseCodeToSelect As String = "")
        Dim coursesTable As DataTable = FetchCoursesTable()
        SetCoursesTable(coursesTable)

        If Not String.IsNullOrWhiteSpace(courseCodeToSelect) Then
            SelectCourseByCode(courseCodeToSelect)
        End If
    End Sub

    Private Function FetchCoursesTable() As DataTable
        Return ReadCoursesFromStorage()
    End Function

    Private Function NormalizeCoursesTable(source As DataTable) As DataTable
        Dim table As DataTable = If(source Is Nothing, CreateEmptyCoursesTable(), source.Copy())

        Dim courseCodeColumn As DataColumn = EnsureCoursesColumn(table, "Course Code", "CourseCode", "Code", "Course_ID")
        Dim courseTitleColumn As DataColumn = EnsureCoursesColumn(table, "Course Title", "CourseTitle", "Title", "Course Name")
        Dim departmentColumn As DataColumn = EnsureCoursesColumn(table, "Department", "Program", "Department Name", "College")
        Dim unitsColumn As DataColumn = EnsureCoursesColumn(table, "Units", "Unit", "Credits", "Credit")

        RemoveCoursesColumns(table,
                             "Year Level",
                             "Section",
                             "Instructor",
                             "Status",
                             "Photo",
                             "Photo Path",
                             "Image",
                             "ImagePath",
                             "Avatar")

        courseCodeColumn.SetOrdinal(0)
        courseTitleColumn.SetOrdinal(1)
        departmentColumn.SetOrdinal(2)
        unitsColumn.SetOrdinal(3)

        For Each row As DataRow In table.Rows
            If row.IsNull("Course Code") Then
                row("Course Code") = String.Empty
            Else
                row("Course Code") = row("Course Code").ToString().Trim()
            End If

            If row.IsNull("Course Title") Then
                row("Course Title") = String.Empty
            Else
                row("Course Title") = row("Course Title").ToString().Trim()
            End If

            If row.IsNull("Department") Then
                row("Department") = String.Empty
            Else
                row("Department") = row("Department").ToString().Trim()
            End If

            If row.IsNull("Units") Then
                row("Units") = String.Empty
            Else
                row("Units") = row("Units").ToString().Trim()
            End If
        Next

        Return table
    End Function

    Private Sub RemoveCoursesColumns(table As DataTable, ParamArray columnNames() As String)
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

    Private Function EnsureCoursesColumn(table As DataTable, targetColumnName As String, ParamArray aliases() As String) As DataColumn
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

    Private Function CreateEmptyCoursesTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Course Code", GetType(String))
        table.Columns.Add("Course Title", GetType(String))
        table.Columns.Add("Department", GetType(String))
        table.Columns.Add("Units", GetType(String))
        Return table
    End Function

    Private Function ReadCoursesFromStorage() As DataTable
        Dim table As DataTable = CreateEmptyCoursesTable()
        If Not File.Exists(_coursesStoragePath) Then
            Return table
        End If

        Try
            Dim json As String = File.ReadAllText(_coursesStoragePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return table
            End If

            Dim records As List(Of CourseStorageRecord) =
                JsonSerializer.Deserialize(Of List(Of CourseStorageRecord))(json, _courseStorageJsonOptions)
            If records Is Nothing Then
                Return table
            End If

            For Each record As CourseStorageRecord In records
                Dim row As DataRow = table.NewRow()
                row("Course Code") = If(record.CourseCode, String.Empty).Trim()
                row("Course Title") = If(record.CourseTitle, String.Empty).Trim()
                row("Department") = If(record.Department, String.Empty).Trim()
                row("Units") = If(record.Units, String.Empty).Trim()
                table.Rows.Add(row)
            Next
        Catch ex As Exception
            MessageBox.Show("Unable to load saved courses data." & Environment.NewLine & ex.Message,
                            "Courses",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
        End Try

        Return table
    End Function

    Private Function PersistCoursesToStorage() As Boolean
        If _coursesTable Is Nothing Then
            Return True
        End If

        Try
            Dim storageDirectory As String = Path.GetDirectoryName(_coursesStoragePath)
            If Not String.IsNullOrWhiteSpace(storageDirectory) Then
                Directory.CreateDirectory(storageDirectory)
            End If

            Dim records As New List(Of CourseStorageRecord)()
            For Each row As DataRow In _coursesTable.Rows
                If row.RowState = DataRowState.Deleted Then
                    Continue For
                End If

                records.Add(New CourseStorageRecord With {
                    .CourseCode = ReadRowValue(row, "Course Code"),
                    .CourseTitle = ReadRowValue(row, "Course Title"),
                    .Department = ReadRowValue(row, "Department"),
                    .Units = ReadRowValue(row, "Units")
                })
            Next

            Dim json As String = JsonSerializer.Serialize(records, _courseStorageJsonOptions)
            File.WriteAllText(_coursesStoragePath, json)
            Return True
        Catch ex As Exception
            MessageBox.Show("Unable to save courses data." & Environment.NewLine & ex.Message,
                            "Courses",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            Return False
        End Try
    End Function

    Private Sub ApplyCoursesFilter()
        If _coursesTable Is Nothing Then
            UpdateCoursesCount()
            Return
        End If

        If String.IsNullOrWhiteSpace(_searchTerm) Then
            _coursesTable.DefaultView.RowFilter = String.Empty
        Else
            Dim escapedTerm As String = EscapeLikeValue(_searchTerm)
            _coursesTable.DefaultView.RowFilter =
                "[Course Code] LIKE '*" & escapedTerm & "*' OR " &
                "[Course Title] LIKE '*" & escapedTerm & "*' OR " &
                "[Department] LIKE '*" & escapedTerm & "*' OR " &
                "[Units] LIKE '*" & escapedTerm & "*'"
        End If

        UpdateCoursesCount()
        EnsureSelectedCourseForDetails()
        RefreshCourseDetailsPanel()
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

    Private Sub UpdateCoursesCount()
        If CoursesCountTextBlock Is Nothing Then
            Return
        End If

        If _coursesTable Is Nothing Then
            CoursesCountTextBlock.Text = "0 courses"
            Return
        End If

        Dim visibleCount As Integer = _coursesTable.DefaultView.Count
        Dim totalCount As Integer = _coursesTable.Rows.Count
        Dim hasSearch As Boolean = Not String.IsNullOrWhiteSpace(_searchTerm)

        If hasSearch AndAlso visibleCount <> totalCount Then
            CoursesCountTextBlock.Text = visibleCount.ToString() & " of " & totalCount.ToString() & " courses"
        Else
            CoursesCountTextBlock.Text = totalCount.ToString() & " courses"
        End If
    End Sub

    Private Sub OpenAddCourseButton_Click(sender As Object, e As RoutedEventArgs)
        BeginAddCourse()
    End Sub

    Private Sub EditSelectedCourseButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        BeginEditCourse(selectedRow)
    End Sub

    Private Sub DeleteSelectedCourseButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        DeleteCourseRow(selectedRow)
    End Sub

    Private Sub CoursesDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        RefreshCourseDetailsPanel()
    End Sub

    Private Sub BeginAddCourse()
        OpenCourseForm(CourseFormMode.Add, "Add Course", "Create a new course record.", "Add Course", Nothing)
        CourseFormCodeTextBox.Focus()
    End Sub

    Private Sub BeginEditCourse(row As DataRow)
        If row Is Nothing Then
            Return
        End If

        OpenCourseForm(CourseFormMode.Edit, "Edit Course", "Update course details.", "Save Changes", row)
        CourseFormTitleTextBox.Focus()
    End Sub

    Private Sub SaveCourseFormButton_Click(sender As Object, e As RoutedEventArgs)
        Dim formValues As New CourseFormValues()
        If Not TryReadCourseForm(formValues) Then
            Return
        End If

        EnsureCoursesTableLoaded()
        Dim targetRow As DataRow = ResolveTargetRowForSave(formValues.CourseCode)
        If targetRow Is Nothing Then
            Return
        End If

        WriteCourseValues(targetRow, formValues)
        If _activeFormMode = CourseFormMode.Add Then
            _coursesTable.Rows.Add(targetRow)
        End If

        _coursesTable.AcceptChanges()
        If Not PersistCoursesToStorage() Then
            Return
        End If

        ApplyCoursesFilter()
        SelectCourseByCode(formValues.CourseCode)
        HideCourseForm()
    End Sub

    Private Sub OpenCourseForm(mode As CourseFormMode,
                               title As String,
                               subtitle As String,
                               actionText As String,
                               row As DataRow)
        _activeFormMode = mode
        CourseFormTitleTextBlock.Text = title
        CourseFormSubtitleTextBlock.Text = subtitle
        SaveCourseFormButton.Content = actionText
        CourseFormCodeTextBox.IsReadOnly = False

        If row Is Nothing Then
            _editingCourseOriginalCode = String.Empty
            ClearCourseFormInputs()
        Else
            _editingCourseOriginalCode = ReadRowValue(row, "Course Code")
            PopulateCourseForm(row)
        End If

        ShowCourseForm()
    End Sub

    Private Sub EnsureCoursesTableLoaded()
        If _coursesTable Is Nothing Then
            SetCoursesTable(Nothing)
        End If
    End Sub

    Private Function ResolveTargetRowForSave(courseCode As String) As DataRow
        Select Case _activeFormMode
            Case CourseFormMode.Add
                If FindCourseRowByCode(courseCode) IsNot Nothing Then
                    MessageBox.Show("Course Code already exists.", "Duplicate Course Code", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return _coursesTable.NewRow()

            Case CourseFormMode.Edit
                Dim targetRow As DataRow = FindCourseRowByCode(_editingCourseOriginalCode)
                If targetRow Is Nothing Then
                    MessageBox.Show("The selected course no longer exists.", "Edit Course", MessageBoxButton.OK, MessageBoxImage.Information)
                    HideCourseForm()
                    Return Nothing
                End If

                If Not String.Equals(_editingCourseOriginalCode, courseCode, StringComparison.OrdinalIgnoreCase) AndAlso
                   FindCourseRowByCode(courseCode) IsNot Nothing Then
                    MessageBox.Show("Course Code already exists.", "Duplicate Course Code", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return targetRow
        End Select

        Return Nothing
    End Function

    Private Sub WriteCourseValues(row As DataRow, values As CourseFormValues)
        row("Course Code") = values.CourseCode
        row("Course Title") = values.CourseTitle
        row("Department") = values.Department
        row("Units") = values.Units
    End Sub

    Private Sub CancelCourseFormButton_Click(sender As Object, e As RoutedEventArgs)
        HideCourseForm()
    End Sub

    Private Sub DeleteCourseRow(row As DataRow)
        If row Is Nothing OrElse row.RowState = DataRowState.Deleted Then
            Return
        End If

        Dim courseTitle As String = ReadRowValue(row, "Course Title")
        Dim courseCode As String = ReadRowValue(row, "Course Code")
        Dim recordLabel As String = If(String.IsNullOrWhiteSpace(courseTitle), courseCode, courseTitle)

        Dim confirmation As MessageBoxResult = MessageBox.Show("Delete " & recordLabel & "?", "Delete Course", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        row.Delete()
        _coursesTable.AcceptChanges()
        If Not PersistCoursesToStorage() Then
            Return
        End If

        ApplyCoursesFilter()

        If _activeFormMode = CourseFormMode.Edit AndAlso
           String.Equals(_editingCourseOriginalCode, courseCode, StringComparison.OrdinalIgnoreCase) Then
            HideCourseForm()
        End If

        EnsureSelectedCourseForDetails()
        RefreshCourseDetailsPanel()
    End Sub

    Private Sub EnsureSelectedCourseForDetails()
        If CoursesDataGrid Is Nothing OrElse _coursesTable Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow IsNot Nothing Then
            Return
        End If

        If _coursesTable.DefaultView.Count > 0 Then
            CoursesDataGrid.SelectedItem = _coursesTable.DefaultView(0)
        End If
    End Sub

    Private Function GetSelectedCourseCode() As String
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return String.Empty
        End If

        Return ReadRowValue(selectedRow, "Course Code")
    End Function

    Private Function TryGetSelectedGridRow() As DataRow
        Dim selectedRowView As DataRowView = TryCast(CoursesDataGrid.SelectedItem, DataRowView)
        If selectedRowView Is Nothing OrElse selectedRowView.Row Is Nothing Then
            Return Nothing
        End If

        If selectedRowView.Row.RowState = DataRowState.Deleted Then
            Return Nothing
        End If

        Return selectedRowView.Row
    End Function

    Private Sub RefreshCourseDetailsPanel()
        If CourseDetailsCodeTextBlock Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        Dim hasSelection As Boolean = selectedRow IsNot Nothing

        If EditSelectedCourseButton IsNot Nothing Then
            EditSelectedCourseButton.IsEnabled = hasSelection
        End If

        If DeleteSelectedCourseButton IsNot Nothing Then
            DeleteSelectedCourseButton.IsEnabled = hasSelection
        End If

        If Not hasSelection Then
            If CourseDetailsSubtitleTextBlock IsNot Nothing Then
                CourseDetailsSubtitleTextBlock.Text = "Select a course from the table."
            End If

            SetDetailsValue(CourseDetailsCodeTextBlock, String.Empty)
            SetDetailsValue(CourseDetailsTitleTextBlock, String.Empty)
            SetDetailsValue(CourseDetailsDepartmentTextBlock, String.Empty)
            SetDetailsValue(CourseDetailsUnitsTextBlock, String.Empty)
            Return
        End If

        If CourseDetailsSubtitleTextBlock IsNot Nothing Then
            CourseDetailsSubtitleTextBlock.Text = "Selected course record."
        End If

        SetDetailsValue(CourseDetailsCodeTextBlock, ReadRowValue(selectedRow, "Course Code"))
        SetDetailsValue(CourseDetailsTitleTextBlock, ReadRowValue(selectedRow, "Course Title"))
        SetDetailsValue(CourseDetailsDepartmentTextBlock, ReadRowValue(selectedRow, "Department"))
        SetDetailsValue(CourseDetailsUnitsTextBlock, ReadRowValue(selectedRow, "Units"))
    End Sub

    Private Sub SetDetailsValue(target As TextBlock, value As String)
        If target Is Nothing Then
            Return
        End If

        Dim normalizedValue As String = If(value, String.Empty).Trim()
        target.Text = If(String.IsNullOrWhiteSpace(normalizedValue), "--", normalizedValue)
    End Sub

    Private Sub ShowCourseForm()
        If CourseFormPanel IsNot Nothing Then
            CourseFormPanel.Visibility = Visibility.Visible
        End If

        If CourseDetailsPanel IsNot Nothing Then
            CourseDetailsPanel.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub HideCourseForm()
        _activeFormMode = CourseFormMode.Add
        _editingCourseOriginalCode = String.Empty
        CourseFormTitleTextBlock.Text = "Add Course"
        CourseFormSubtitleTextBlock.Text = "Create a new course record."
        SaveCourseFormButton.Content = "Add Course"
        CourseFormCodeTextBox.IsReadOnly = False
        ClearCourseFormInputs()

        If CourseFormPanel IsNot Nothing Then
            CourseFormPanel.Visibility = Visibility.Collapsed
        End If

        If CourseDetailsPanel IsNot Nothing Then
            CourseDetailsPanel.Visibility = Visibility.Visible
        End If

        RefreshCourseDetailsPanel()
    End Sub

    Private Sub PopulateCourseForm(row As DataRow)
        CourseFormCodeTextBox.Text = ReadRowValue(row, "Course Code")
        CourseFormTitleTextBox.Text = ReadRowValue(row, "Course Title")
        CourseFormDepartmentTextBox.Text = ReadRowValue(row, "Department")
        CourseFormUnitsTextBox.Text = ReadRowValue(row, "Units")
    End Sub

    Private Sub ClearCourseFormInputs()
        CourseFormCodeTextBox.Text = String.Empty
        CourseFormTitleTextBox.Text = String.Empty
        CourseFormDepartmentTextBox.Text = String.Empty
        CourseFormUnitsTextBox.Text = String.Empty
    End Sub

    Private Function TryReadCourseForm(ByRef values As CourseFormValues) As Boolean
        values.CourseCode = If(CourseFormCodeTextBox.Text, String.Empty).Trim()
        values.CourseTitle = If(CourseFormTitleTextBox.Text, String.Empty).Trim()
        values.Department = If(CourseFormDepartmentTextBox.Text, String.Empty).Trim()
        values.Units = If(CourseFormUnitsTextBox.Text, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(values.CourseCode) Then
            MessageBox.Show("Course Code is required.", "Course Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.CourseTitle) Then
            MessageBox.Show("Course Title is required.", "Course Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        Return True
    End Function

    Private Function FindCourseRowByCode(courseCode As String) As DataRow
        If _coursesTable Is Nothing Then
            Return Nothing
        End If

        Dim normalizedCourseCode As String = If(courseCode, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedCourseCode) Then
            Return Nothing
        End If

        For Each row As DataRow In _coursesTable.Rows
            If row.RowState = DataRowState.Deleted Then
                Continue For
            End If

            Dim candidateCode As String = ReadRowValue(row, "Course Code")
            If String.Equals(candidateCode, normalizedCourseCode, StringComparison.OrdinalIgnoreCase) Then
                Return row
            End If
        Next

        Return Nothing
    End Function

    Private Sub SelectCourseByCode(courseCode As String)
        If _coursesTable Is Nothing Then
            Return
        End If

        Dim normalizedCourseCode As String = If(courseCode, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedCourseCode) Then
            Return
        End If

        For Each rowView As DataRowView In _coursesTable.DefaultView
            Dim candidateCode As String = If(rowView("Course Code"), String.Empty).ToString().Trim()
            If String.Equals(candidateCode, normalizedCourseCode, StringComparison.OrdinalIgnoreCase) Then
                CoursesDataGrid.SelectedItem = rowView
                CoursesDataGrid.ScrollIntoView(rowView)
                RefreshCourseDetailsPanel()
                Return
            End If
        Next

        RefreshCourseDetailsPanel()
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
End Class
