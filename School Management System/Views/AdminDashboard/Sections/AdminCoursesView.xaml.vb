Imports System.Collections.Generic
Imports System.Data
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

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

    Private _coursesTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As CourseFormMode = CourseFormMode.Add
    Private _editingCourseOriginalCode As String = String.Empty
    Private ReadOnly _courseManagementService As New CourseManagementService()

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
        _coursesTable = If(table, CreateEmptyCoursesTable())
        CoursesDataGrid.ItemsSource = _coursesTable.DefaultView
        ApplyCoursesFilter()
        UpdateCoursesCount()
        EnsureSelectedCourseForDetails()
        RefreshCourseDetailsPanel()
    End Sub

    Private Sub LoadCoursesTable(Optional courseCodeToSelect As String = "")
        Dim result = _courseManagementService.GetCourses()

        If Not result.IsSuccess Then
            SetCoursesTable(CreateEmptyCoursesTable())
            MessageBox.Show(result.Message,
                            "Courses",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return
        End If

        SetCoursesTable(BuildCoursesTable(result.Data))

        If Not String.IsNullOrWhiteSpace(courseCodeToSelect) Then
            SelectCourseByCode(courseCodeToSelect)
        End If
    End Sub

    Private Function BuildCoursesTable(records As IEnumerable(Of CourseRecord)) As DataTable
        Dim table As DataTable = CreateEmptyCoursesTable()

        If records Is Nothing Then
            Return table
        End If

        For Each record As CourseRecord In records
            Dim row As DataRow = table.NewRow()
            row("Course Code") = record.CourseCode
            row("Course Title") = record.CourseName
            row("Department") = record.DepartmentDisplayName
            row("Units") = record.Units
            table.Rows.Add(row)
        Next

        Return table
    End Function

    Private Function CreateEmptyCoursesTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Course Code", GetType(String))
        table.Columns.Add("Course Title", GetType(String))
        table.Columns.Add("Department", GetType(String))
        table.Columns.Add("Units", GetType(String))
        Return table
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

        Dim request As New CourseSaveRequest() With {
            .OriginalCourseCode = _editingCourseOriginalCode,
            .CourseCode = formValues.CourseCode,
            .CourseName = formValues.CourseTitle,
            .DepartmentText = formValues.Department,
            .Units = formValues.Units
        }

        Dim isAddMode As Boolean = _activeFormMode = CourseFormMode.Add
        Dim result =
            If(isAddMode,
               _courseManagementService.CreateCourse(request),
               _courseManagementService.UpdateCourse(request))

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Courses",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        LoadCoursesTable(result.Data.CourseCode)
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

        Dim confirmation As MessageBoxResult =
            MessageBox.Show("Delete " & recordLabel & "?",
                            "Delete Course",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        Dim result = _courseManagementService.DeleteCourse(courseCode)
        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Courses",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        If _activeFormMode = CourseFormMode.Edit AndAlso
           String.Equals(_editingCourseOriginalCode,
                         courseCode,
                         StringComparison.OrdinalIgnoreCase) Then
            HideCourseForm()
        End If

        LoadCoursesTable()
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
