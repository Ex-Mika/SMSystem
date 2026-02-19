Imports System.Data
Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json

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

    Private Class SubjectStorageRecord
        Public Property SubjectCode As String
        Public Property SubjectName As String
        Public Property Course As String
        Public Property YearLevel As String
        Public Property Units As String
    End Class

    Private _subjectsTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As SubjectFormMode = SubjectFormMode.Add
    Private _editingSubjectOriginalCode As String = String.Empty
    Private ReadOnly _subjectsStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "subjects.json")
    Private ReadOnly _subjectStorageJsonOptions As New JsonSerializerOptions() With {
        .WriteIndented = True
    }

    Public Sub New()
        InitializeComponent()
        LoadSubjectsTable()
        HideSubjectForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplySubjectsFilter()
    End Sub

    Public Sub SetSubjectsTable(table As DataTable)
        _subjectsTable = NormalizeSubjectsTable(table)
        SubjectsDataGrid.ItemsSource = _subjectsTable.DefaultView
        ApplySubjectsFilter()
        UpdateSubjectsCount()
        EnsureSelectedSubjectForDetails()
        RefreshSubjectDetailsPanel()
    End Sub

    Private Sub LoadSubjectsTable()
        Dim subjectsTable As DataTable = FetchSubjectsTable()
        SetSubjectsTable(subjectsTable)
    End Sub

    Private Function FetchSubjectsTable() As DataTable
        Return ReadSubjectsFromStorage()
    End Function

    Private Function NormalizeSubjectsTable(source As DataTable) As DataTable
        Dim table As DataTable = If(source Is Nothing, CreateEmptySubjectsTable(), source.Copy())

        Dim subjectCodeColumn As DataColumn = EnsureSubjectsColumn(table, "Subject Code", "SubjectCode", "Code", "Subject_ID")
        Dim subjectNameColumn As DataColumn = EnsureSubjectsColumn(table, "Subject Name", "SubjectName", "Name", "Subject Name")
        Dim courseColumn As DataColumn = EnsureSubjectsColumn(table, "Course", "Program", "Course Name", "College")
        Dim yearLevelColumn As DataColumn = EnsureSubjectsColumn(table, "Year Level", "YearLevel", "Year")
        Dim unitsColumn As DataColumn = EnsureSubjectsColumn(table, "Units", "Unit", "Credits", "Credit")

        RemoveSubjectsColumns(table,
                             "Section",
                             "Instructor",
                             "Status",
                             "Photo",
                             "Photo Path",
                             "Image",
                             "ImagePath",
                             "Avatar")

        subjectCodeColumn.SetOrdinal(0)
        subjectNameColumn.SetOrdinal(1)
        courseColumn.SetOrdinal(2)
        yearLevelColumn.SetOrdinal(3)
        unitsColumn.SetOrdinal(4)

        For Each row As DataRow In table.Rows
            If row.IsNull("Subject Code") Then
                row("Subject Code") = String.Empty
            Else
                row("Subject Code") = row("Subject Code").ToString().Trim()
            End If

            If row.IsNull("Subject Name") Then
                row("Subject Name") = String.Empty
            Else
                row("Subject Name") = row("Subject Name").ToString().Trim()
            End If

            If row.IsNull("Course") Then
                row("Course") = String.Empty
            Else
                row("Course") = row("Course").ToString().Trim()
            End If

            If row.IsNull("Year Level") Then
                row("Year Level") = String.Empty
            Else
                row("Year Level") = row("Year Level").ToString().Trim()
            End If

            If row.IsNull("Units") Then
                row("Units") = String.Empty
            Else
                row("Units") = row("Units").ToString().Trim()
            End If
        Next

        Return table
    End Function

    Private Sub RemoveSubjectsColumns(table As DataTable, ParamArray columnNames() As String)
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

    Private Function EnsureSubjectsColumn(table As DataTable, targetColumnName As String, ParamArray aliases() As String) As DataColumn
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

    Private Function CreateEmptySubjectsTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Subject Code", GetType(String))
        table.Columns.Add("Subject Name", GetType(String))
        table.Columns.Add("Course", GetType(String))
        table.Columns.Add("Year Level", GetType(String))
        table.Columns.Add("Units", GetType(String))
        Return table
    End Function

    Private Function ReadSubjectsFromStorage() As DataTable
        Dim table As DataTable = CreateEmptySubjectsTable()
        If Not File.Exists(_subjectsStoragePath) Then
            Return table
        End If

        Try
            Dim json As String = File.ReadAllText(_subjectsStoragePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return table
            End If

            Dim records As List(Of SubjectStorageRecord) =
                JsonSerializer.Deserialize(Of List(Of SubjectStorageRecord))(json, _subjectStorageJsonOptions)
            If records Is Nothing Then
                Return table
            End If

            For Each record As SubjectStorageRecord In records
                Dim row As DataRow = table.NewRow()
                row("Subject Code") = If(record.SubjectCode, String.Empty).Trim()
                row("Subject Name") = If(record.SubjectName, String.Empty).Trim()
                row("Course") = If(record.Course, String.Empty).Trim()
                row("Year Level") = If(record.YearLevel, String.Empty).Trim()
                row("Units") = If(record.Units, String.Empty).Trim()
                table.Rows.Add(row)
            Next
        Catch ex As Exception
            MessageBox.Show("Unable to load saved subjects data." & Environment.NewLine & ex.Message,
                            "Subjects",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
        End Try

        Return table
    End Function

    Private Function PersistSubjectsToStorage() As Boolean
        If _subjectsTable Is Nothing Then
            Return True
        End If

        Try
            Dim storageDirectory As String = Path.GetDirectoryName(_subjectsStoragePath)
            If Not String.IsNullOrWhiteSpace(storageDirectory) Then
                Directory.CreateDirectory(storageDirectory)
            End If

            Dim records As New List(Of SubjectStorageRecord)()
            For Each row As DataRow In _subjectsTable.Rows
                If row.RowState = DataRowState.Deleted Then
                    Continue For
                End If

                records.Add(New SubjectStorageRecord With {
                    .SubjectCode = ReadRowValue(row, "Subject Code"),
                    .SubjectName = ReadRowValue(row, "Subject Name"),
                    .Course = ReadRowValue(row, "Course"),
                    .YearLevel = ReadRowValue(row, "Year Level"),
                    .Units = ReadRowValue(row, "Units")
                })
            Next

            Dim json As String = JsonSerializer.Serialize(records, _subjectStorageJsonOptions)
            File.WriteAllText(_subjectsStoragePath, json)
            Return True
        Catch ex As Exception
            MessageBox.Show("Unable to save subjects data." & Environment.NewLine & ex.Message,
                            "Subjects",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            Return False
        End Try
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

        EnsureSubjectsTableLoaded()
        Dim targetRow As DataRow = ResolveTargetRowForSave(formValues.SubjectCode)
        If targetRow Is Nothing Then
            Return
        End If

        WriteSubjectValues(targetRow, formValues)
        If _activeFormMode = SubjectFormMode.Add Then
            _subjectsTable.Rows.Add(targetRow)
        End If

        _subjectsTable.AcceptChanges()
        If Not PersistSubjectsToStorage() Then
            Return
        End If

        ApplySubjectsFilter()
        SelectSubjectByCode(formValues.SubjectCode)
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

    Private Sub EnsureSubjectsTableLoaded()
        If _subjectsTable Is Nothing Then
            SetSubjectsTable(Nothing)
        End If
    End Sub

    Private Function ResolveTargetRowForSave(subjectCode As String) As DataRow
        Select Case _activeFormMode
            Case SubjectFormMode.Add
                If FindSubjectRowByCode(subjectCode) IsNot Nothing Then
                    MessageBox.Show("Subject Code already exists.", "Duplicate Subject Code", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return _subjectsTable.NewRow()

            Case SubjectFormMode.Edit
                Dim targetRow As DataRow = FindSubjectRowByCode(_editingSubjectOriginalCode)
                If targetRow Is Nothing Then
                    MessageBox.Show("The selected subject no longer exists.", "Edit Subject", MessageBoxButton.OK, MessageBoxImage.Information)
                    HideSubjectForm()
                    Return Nothing
                End If

                If Not String.Equals(_editingSubjectOriginalCode, subjectCode, StringComparison.OrdinalIgnoreCase) AndAlso
                   FindSubjectRowByCode(subjectCode) IsNot Nothing Then
                    MessageBox.Show("Subject Code already exists.", "Duplicate Subject Code", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return targetRow
        End Select

        Return Nothing
    End Function

    Private Sub WriteSubjectValues(row As DataRow, values As SubjectFormValues)
        row("Subject Code") = values.SubjectCode
        row("Subject Name") = values.SubjectName
        row("Course") = values.Course
        row("Year Level") = values.YearLevel
        row("Units") = values.Units
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

        Dim confirmation As MessageBoxResult = MessageBox.Show("Delete " & recordLabel & "?", "Delete Subject", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        row.Delete()
        _subjectsTable.AcceptChanges()
        If Not PersistSubjectsToStorage() Then
            Return
        End If

        ApplySubjectsFilter()

        If _activeFormMode = SubjectFormMode.Edit AndAlso
           String.Equals(_editingSubjectOriginalCode, subjectCode, StringComparison.OrdinalIgnoreCase) Then
            HideSubjectForm()
        End If

        EnsureSelectedSubjectForDetails()
        RefreshSubjectDetailsPanel()
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

    Private Function FindSubjectRowByCode(subjectCode As String) As DataRow
        If _subjectsTable Is Nothing Then
            Return Nothing
        End If

        Dim normalizedSubjectCode As String = If(subjectCode, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedSubjectCode) Then
            Return Nothing
        End If

        For Each row As DataRow In _subjectsTable.Rows
            If row.RowState = DataRowState.Deleted Then
                Continue For
            End If

            Dim candidateCode As String = ReadRowValue(row, "Subject Code")
            If String.Equals(candidateCode, normalizedSubjectCode, StringComparison.OrdinalIgnoreCase) Then
                Return row
            End If
        Next

        Return Nothing
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

