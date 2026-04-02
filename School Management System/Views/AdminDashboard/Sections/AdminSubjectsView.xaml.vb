Imports System.Collections.Generic
Imports System.Data
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class AdminSubjectsView
    Private Enum SubjectFormMode
        Add
        Edit
    End Enum

    Private Structure SubjectFormValues
        Public SubjectCode As String
        Public SubjectName As String
        Public Course As String
        Public YearLevel As String
        Public Units As String
    End Structure

    Private _subjectsTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As SubjectFormMode = SubjectFormMode.Add
    Private _editingSubjectOriginalCode As String = String.Empty
    Private ReadOnly _subjectManagementService As New SubjectManagementService()

    Public Sub New()
        InitializeComponent()
        LoadSubjectsTable()
        HideSubjectForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplySubjectsFilter()
    End Sub

    Public Sub RefreshData()
        LoadSubjectsTable(GetSelectedSubjectCode())
    End Sub

    Public Sub SetSubjectsTable(table As DataTable)
        _subjectsTable = If(table, CreateEmptySubjectsTable())
        SubjectsDataGrid.ItemsSource = _subjectsTable.DefaultView
        ApplySubjectsFilter()
        UpdateSubjectsCount()
        EnsureSelectedSubjectForDetails()
        RefreshSubjectDetailsPanel()
    End Sub

    Private Sub LoadSubjectsTable(Optional subjectCodeToSelect As String = "")
        Dim result = _subjectManagementService.GetSubjects()

        If Not result.IsSuccess Then
            SetSubjectsTable(CreateEmptySubjectsTable())
            MessageBox.Show(result.Message,
                            "Subjects",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return
        End If

        SetSubjectsTable(BuildSubjectsTable(result.Data))

        If Not String.IsNullOrWhiteSpace(subjectCodeToSelect) Then
            SelectSubjectByCode(subjectCodeToSelect)
        End If
    End Sub

    Private Function BuildSubjectsTable(records As IEnumerable(Of SubjectRecord)) As DataTable
        Dim table As DataTable = CreateEmptySubjectsTable()

        If records Is Nothing Then
            Return table
        End If

        For Each record As SubjectRecord In records
            Dim row As DataRow = table.NewRow()
            row("Subject Code") = record.SubjectCode
            row("Subject Name") = record.SubjectName
            row("Course") = record.CourseDisplayName
            row("Year Level") = record.YearLevel
            row("Units") = record.Units
            table.Rows.Add(row)
        Next

        Return table
    End Function

    Private Function CreateEmptySubjectsTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Subject Code", GetType(String))
        table.Columns.Add("Subject Name", GetType(String))
        table.Columns.Add("Course", GetType(String))
        table.Columns.Add("Year Level", GetType(String))
        table.Columns.Add("Units", GetType(String))
        Return table
    End Function

    Private Sub ApplySubjectsFilter()
        If _subjectsTable Is Nothing Then
            UpdateSubjectsCount()
            Return
        End If

        If String.IsNullOrWhiteSpace(_searchTerm) Then
            _subjectsTable.DefaultView.RowFilter = String.Empty
        Else
            Dim escapedTerm As String = EscapeLikeValue(_searchTerm)
            _subjectsTable.DefaultView.RowFilter =
                "[Subject Code] LIKE '*" & escapedTerm & "*' OR " &
                "[Subject Name] LIKE '*" & escapedTerm & "*' OR " &
                "[Course] LIKE '*" & escapedTerm & "*' OR " &
                "[Year Level] LIKE '*" & escapedTerm & "*' OR " &
                "[Units] LIKE '*" & escapedTerm & "*'"
        End If

        UpdateSubjectsCount()
        EnsureSelectedSubjectForDetails()
        RefreshSubjectDetailsPanel()
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

    Private Sub UpdateSubjectsCount()
        If SubjectsCountTextBlock Is Nothing Then
            Return
        End If

        If _subjectsTable Is Nothing Then
            SubjectsCountTextBlock.Text = "0 subjects"
            Return
        End If

        Dim visibleCount As Integer = _subjectsTable.DefaultView.Count
        Dim totalCount As Integer = _subjectsTable.Rows.Count
        Dim hasSearch As Boolean = Not String.IsNullOrWhiteSpace(_searchTerm)

        If hasSearch AndAlso visibleCount <> totalCount Then
            SubjectsCountTextBlock.Text = visibleCount.ToString() & " of " & totalCount.ToString() & " subjects"
        Else
            SubjectsCountTextBlock.Text = totalCount.ToString() & " subjects"
        End If
    End Sub

    Private Sub OpenAddSubjectButton_Click(sender As Object, e As RoutedEventArgs)
        BeginAddSubject()
    End Sub

    Private Sub EditSelectedSubjectButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        BeginEditSubject(selectedRow)
    End Sub

    Private Sub DeleteSelectedSubjectButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        DeleteSubjectRow(selectedRow)
    End Sub

    Private Sub SubjectsDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        RefreshSubjectDetailsPanel()
    End Sub

    Private Sub BeginAddSubject()
        OpenSubjectForm(SubjectFormMode.Add, "Add Subject", "Create a new subject record.", "Add Subject", Nothing)
        SubjectFormCodeTextBox.Focus()
    End Sub

    Private Sub BeginEditSubject(row As DataRow)
        If row Is Nothing Then
            Return
        End If

        OpenSubjectForm(SubjectFormMode.Edit, "Edit Subject", "Update subject details.", "Save Changes", row)
        SubjectFormNameTextBox.Focus()
    End Sub

    Private Sub SaveSubjectFormButton_Click(sender As Object, e As RoutedEventArgs)
        Dim formValues As New SubjectFormValues()
        If Not TryReadSubjectForm(formValues) Then
            Return
        End If

        Dim request As New SubjectSaveRequest() With {
            .OriginalSubjectCode = _editingSubjectOriginalCode,
            .SubjectCode = formValues.SubjectCode,
            .SubjectName = formValues.SubjectName,
            .CourseText = formValues.Course,
            .YearLevel = formValues.YearLevel,
            .Units = formValues.Units
        }

        Dim isAddMode As Boolean = _activeFormMode = SubjectFormMode.Add
        Dim result =
            If(isAddMode,
               _subjectManagementService.CreateSubject(request),
               _subjectManagementService.UpdateSubject(request))

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Subjects",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        LoadSubjectsTable(result.Data.SubjectCode)
        HideSubjectForm()
    End Sub

    Private Sub OpenSubjectForm(mode As SubjectFormMode,
                               title As String,
                               subtitle As String,
                               actionText As String,
                               row As DataRow)
        _activeFormMode = mode
        SubjectFormTitleTextBlock.Text = title
        SubjectFormSubtitleTextBlock.Text = subtitle
        SaveSubjectFormButton.Content = actionText
        SubjectFormCodeTextBox.IsReadOnly = False

        If row Is Nothing Then
            _editingSubjectOriginalCode = String.Empty
            ClearSubjectFormInputs()
        Else
            _editingSubjectOriginalCode = ReadRowValue(row, "Subject Code")
            PopulateSubjectForm(row)
        End If

        ShowSubjectForm()
    End Sub

    Private Sub CancelSubjectFormButton_Click(sender As Object, e As RoutedEventArgs)
        HideSubjectForm()
    End Sub

    Private Sub DeleteSubjectRow(row As DataRow)
        If row Is Nothing OrElse row.RowState = DataRowState.Deleted Then
            Return
        End If

        Dim subjectName As String = ReadRowValue(row, "Subject Name")
        Dim subjectCode As String = ReadRowValue(row, "Subject Code")
        Dim recordLabel As String = If(String.IsNullOrWhiteSpace(subjectName), subjectCode, subjectName)

        Dim confirmation As MessageBoxResult =
            MessageBox.Show("Delete " & recordLabel & "?",
                            "Delete Subject",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        Dim result = _subjectManagementService.DeleteSubject(subjectCode)
        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Subjects",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        If _activeFormMode = SubjectFormMode.Edit AndAlso
           String.Equals(_editingSubjectOriginalCode,
                         subjectCode,
                         StringComparison.OrdinalIgnoreCase) Then
            HideSubjectForm()
        End If

        LoadSubjectsTable()
    End Sub

    Private Sub EnsureSelectedSubjectForDetails()
        If SubjectsDataGrid Is Nothing OrElse _subjectsTable Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow IsNot Nothing Then
            Return
        End If

        If _subjectsTable.DefaultView.Count > 0 Then
            SubjectsDataGrid.SelectedItem = _subjectsTable.DefaultView(0)
        End If
    End Sub

    Private Function GetSelectedSubjectCode() As String
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return String.Empty
        End If

        Return ReadRowValue(selectedRow, "Subject Code")
    End Function

    Private Function TryGetSelectedGridRow() As DataRow
        Dim selectedRowView As DataRowView = TryCast(SubjectsDataGrid.SelectedItem, DataRowView)
        If selectedRowView Is Nothing OrElse selectedRowView.Row Is Nothing Then
            Return Nothing
        End If

        If selectedRowView.Row.RowState = DataRowState.Deleted Then
            Return Nothing
        End If

        Return selectedRowView.Row
    End Function

    Private Sub RefreshSubjectDetailsPanel()
        If SubjectDetailsCodeTextBlock Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        Dim hasSelection As Boolean = selectedRow IsNot Nothing

        If EditSelectedSubjectButton IsNot Nothing Then
            EditSelectedSubjectButton.IsEnabled = hasSelection
        End If

        If DeleteSelectedSubjectButton IsNot Nothing Then
            DeleteSelectedSubjectButton.IsEnabled = hasSelection
        End If

        If Not hasSelection Then
            If SubjectDetailsSubtitleTextBlock IsNot Nothing Then
                SubjectDetailsSubtitleTextBlock.Text = "Select a subject from the table."
            End If

            SetDetailsValue(SubjectDetailsCodeTextBlock, String.Empty)
            SetDetailsValue(SubjectDetailsNameTextBlock, String.Empty)
            SetDetailsValue(SubjectDetailsCourseTextBlock, String.Empty)
            SetDetailsValue(SubjectDetailsYearLevelTextBlock, String.Empty)
            SetDetailsValue(SubjectDetailsUnitsTextBlock, String.Empty)
            Return
        End If

        If SubjectDetailsSubtitleTextBlock IsNot Nothing Then
            SubjectDetailsSubtitleTextBlock.Text = "Selected subject record."
        End If

        SetDetailsValue(SubjectDetailsCodeTextBlock, ReadRowValue(selectedRow, "Subject Code"))
        SetDetailsValue(SubjectDetailsNameTextBlock, ReadRowValue(selectedRow, "Subject Name"))
        SetDetailsValue(SubjectDetailsCourseTextBlock, ReadRowValue(selectedRow, "Course"))
        SetDetailsValue(SubjectDetailsYearLevelTextBlock, ReadRowValue(selectedRow, "Year Level"))
        SetDetailsValue(SubjectDetailsUnitsTextBlock, ReadRowValue(selectedRow, "Units"))
    End Sub

    Private Sub SetDetailsValue(target As TextBlock, value As String)
        If target Is Nothing Then
            Return
        End If

        Dim normalizedValue As String = If(value, String.Empty).Trim()
        target.Text = If(String.IsNullOrWhiteSpace(normalizedValue), "--", normalizedValue)
    End Sub

    Private Sub ShowSubjectForm()
        If SubjectFormPanel IsNot Nothing Then
            SubjectFormPanel.Visibility = Visibility.Visible
        End If

        If SubjectDetailsPanel IsNot Nothing Then
            SubjectDetailsPanel.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub HideSubjectForm()
        _activeFormMode = SubjectFormMode.Add
        _editingSubjectOriginalCode = String.Empty
        SubjectFormTitleTextBlock.Text = "Add Subject"
        SubjectFormSubtitleTextBlock.Text = "Create a new subject record."
        SaveSubjectFormButton.Content = "Add Subject"
        SubjectFormCodeTextBox.IsReadOnly = False
        ClearSubjectFormInputs()

        If SubjectFormPanel IsNot Nothing Then
            SubjectFormPanel.Visibility = Visibility.Collapsed
        End If

        If SubjectDetailsPanel IsNot Nothing Then
            SubjectDetailsPanel.Visibility = Visibility.Visible
        End If

        RefreshSubjectDetailsPanel()
    End Sub

    Private Sub PopulateSubjectForm(row As DataRow)
        SubjectFormCodeTextBox.Text = ReadRowValue(row, "Subject Code")
        SubjectFormNameTextBox.Text = ReadRowValue(row, "Subject Name")
        SubjectFormCourseTextBox.Text = ReadRowValue(row, "Course")
        SubjectFormYearLevelTextBox.Text = ReadRowValue(row, "Year Level")
        SubjectFormUnitsTextBox.Text = ReadRowValue(row, "Units")
    End Sub

    Private Sub ClearSubjectFormInputs()
        SubjectFormCodeTextBox.Text = String.Empty
        SubjectFormNameTextBox.Text = String.Empty
        SubjectFormCourseTextBox.Text = String.Empty
        SubjectFormYearLevelTextBox.Text = String.Empty
        SubjectFormUnitsTextBox.Text = String.Empty
    End Sub

    Private Function TryReadSubjectForm(ByRef values As SubjectFormValues) As Boolean
        values.SubjectCode = If(SubjectFormCodeTextBox.Text, String.Empty).Trim()
        values.SubjectName = If(SubjectFormNameTextBox.Text, String.Empty).Trim()
        values.Course = If(SubjectFormCourseTextBox.Text, String.Empty).Trim()
        values.YearLevel = If(SubjectFormYearLevelTextBox.Text, String.Empty).Trim()
        values.Units = If(SubjectFormUnitsTextBox.Text, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(values.SubjectCode) Then
            MessageBox.Show("Subject Code is required.", "Subject Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.SubjectName) Then
            MessageBox.Show("Subject Name is required.", "Subject Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        Return True
    End Function

    Private Sub SelectSubjectByCode(subjectCode As String)
        If _subjectsTable Is Nothing Then
            Return
        End If

        Dim normalizedSubjectCode As String = If(subjectCode, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedSubjectCode) Then
            Return
        End If

        For Each rowView As DataRowView In _subjectsTable.DefaultView
            Dim candidateCode As String = If(rowView("Subject Code"), String.Empty).ToString().Trim()
            If String.Equals(candidateCode, normalizedSubjectCode, StringComparison.OrdinalIgnoreCase) Then
                SubjectsDataGrid.SelectedItem = rowView
                SubjectsDataGrid.ScrollIntoView(rowView)
                RefreshSubjectDetailsPanel()
                Return
            End If
        Next

        RefreshSubjectDetailsPanel()
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
