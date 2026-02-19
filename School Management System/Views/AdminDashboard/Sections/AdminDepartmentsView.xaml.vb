Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Text.Json

Class AdminDepartmentsView
    Private Enum DepartmentFormMode
        Add
        Edit
    End Enum

    Private Structure DepartmentFormValues
        Public DepartmentId As String
        Public DepartmentName As String
        Public Head As String
    End Structure

    Private Class DepartmentStorageRecord
        Public Property DepartmentId As String
        Public Property DepartmentName As String
        Public Property Head As String
    End Class

    Private _departmentsTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As DepartmentFormMode = DepartmentFormMode.Add
    Private _editingDepartmentOriginalId As String = String.Empty
    Private ReadOnly _departmentsStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "departments.json")
    Private ReadOnly _departmentStorageJsonOptions As New JsonSerializerOptions() With {
        .WriteIndented = True
    }

    Public Sub New()
        InitializeComponent()
        LoadDepartmentsTable()
        HideDepartmentForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyDepartmentsFilter()
    End Sub

    Public Sub SetDepartmentsTable(table As DataTable)
        _departmentsTable = NormalizeDepartmentsTable(table)
        DepartmentsDataGrid.ItemsSource = _departmentsTable.DefaultView
        ApplyDepartmentsFilter()
        UpdateDepartmentsCount()
        EnsureSelectedDepartmentForDetails()
        RefreshDepartmentDetailsPanel()
    End Sub

    Private Sub LoadDepartmentsTable()
        Dim departmentsTable As DataTable = FetchDepartmentsTable()
        SetDepartmentsTable(departmentsTable)
    End Sub

    Private Function FetchDepartmentsTable() As DataTable
        Return ReadDepartmentsFromStorage()
    End Function

    Private Function NormalizeDepartmentsTable(source As DataTable) As DataTable
        Dim table As DataTable = If(source Is Nothing, CreateEmptyDepartmentsTable(), source.Copy())

        Dim departmentIdColumn As DataColumn = EnsureDepartmentsColumn(table, "Department ID", "DepartmentId", "ID", "Department Code")
        Dim departmentNameColumn As DataColumn = EnsureDepartmentsColumn(table, "Department Name", "DepartmentName", "Name", "Program")
        Dim headColumn As DataColumn = EnsureDepartmentsColumn(table, "Head", "HeadName", "Chair", "Chairman", "Coordinator")

        RemoveDepartmentsColumns(table,
                                 "Status",
                                 "State",
                                 "Photo",
                                 "Photo Path",
                                 "Image",
                                 "ImagePath",
                                 "Avatar")

        departmentIdColumn.SetOrdinal(0)
        departmentNameColumn.SetOrdinal(1)
        headColumn.SetOrdinal(2)

        For Each row As DataRow In table.Rows
            If row.IsNull("Department ID") Then
                row("Department ID") = String.Empty
            Else
                row("Department ID") = row("Department ID").ToString().Trim()
            End If

            If row.IsNull("Department Name") Then
                row("Department Name") = String.Empty
            Else
                row("Department Name") = row("Department Name").ToString().Trim()
            End If

            If row.IsNull("Head") Then
                row("Head") = String.Empty
            Else
                row("Head") = row("Head").ToString().Trim()
            End If
        Next

        Return table
    End Function

    Private Sub RemoveDepartmentsColumns(table As DataTable, ParamArray columnNames() As String)
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

    Private Function EnsureDepartmentsColumn(table As DataTable, targetColumnName As String, ParamArray aliases() As String) As DataColumn
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

    Private Function CreateEmptyDepartmentsTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Department ID", GetType(String))
        table.Columns.Add("Department Name", GetType(String))
        table.Columns.Add("Head", GetType(String))
        Return table
    End Function

    Private Function ReadDepartmentsFromStorage() As DataTable
        Dim table As DataTable = CreateEmptyDepartmentsTable()
        If Not File.Exists(_departmentsStoragePath) Then
            Return table
        End If

        Try
            Dim json As String = File.ReadAllText(_departmentsStoragePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return table
            End If

            Dim records As List(Of DepartmentStorageRecord) =
                JsonSerializer.Deserialize(Of List(Of DepartmentStorageRecord))(json, _departmentStorageJsonOptions)
            If records Is Nothing Then
                Return table
            End If

            For Each record As DepartmentStorageRecord In records
                Dim row As DataRow = table.NewRow()
                row("Department ID") = If(record.DepartmentId, String.Empty).Trim()
                row("Department Name") = If(record.DepartmentName, String.Empty).Trim()
                row("Head") = If(record.Head, String.Empty).Trim()
                table.Rows.Add(row)
            Next
        Catch ex As Exception
            MessageBox.Show("Unable to load saved departments data." & Environment.NewLine & ex.Message,
                            "Departments",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
        End Try

        Return table
    End Function

    Private Function PersistDepartmentsToStorage() As Boolean
        If _departmentsTable Is Nothing Then
            Return True
        End If

        Try
            Dim storageDirectory As String = Path.GetDirectoryName(_departmentsStoragePath)
            If Not String.IsNullOrWhiteSpace(storageDirectory) Then
                Directory.CreateDirectory(storageDirectory)
            End If

            Dim records As New List(Of DepartmentStorageRecord)()
            For Each row As DataRow In _departmentsTable.Rows
                If row.RowState = DataRowState.Deleted Then
                    Continue For
                End If

                records.Add(New DepartmentStorageRecord With {
                    .DepartmentId = ReadRowValue(row, "Department ID"),
                    .DepartmentName = ReadRowValue(row, "Department Name"),
                    .Head = ReadRowValue(row, "Head")
                })
            Next

            Dim json As String = JsonSerializer.Serialize(records, _departmentStorageJsonOptions)
            File.WriteAllText(_departmentsStoragePath, json)
            Return True
        Catch ex As Exception
            MessageBox.Show("Unable to save departments data." & Environment.NewLine & ex.Message,
                            "Departments",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            Return False
        End Try
    End Function

    Private Sub ApplyDepartmentsFilter()
        If _departmentsTable Is Nothing Then
            UpdateDepartmentsCount()
            Return
        End If

        If String.IsNullOrWhiteSpace(_searchTerm) Then
            _departmentsTable.DefaultView.RowFilter = String.Empty
        Else
            Dim escapedTerm As String = EscapeLikeValue(_searchTerm)
            _departmentsTable.DefaultView.RowFilter =
                "[Department ID] LIKE '*" & escapedTerm & "*' OR " &
                "[Department Name] LIKE '*" & escapedTerm & "*' OR " &
                "[Head] LIKE '*" & escapedTerm & "*'"
        End If

        UpdateDepartmentsCount()
        EnsureSelectedDepartmentForDetails()
        RefreshDepartmentDetailsPanel()
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

    Private Sub UpdateDepartmentsCount()
        If DepartmentsCountTextBlock Is Nothing Then
            Return
        End If

        If _departmentsTable Is Nothing Then
            DepartmentsCountTextBlock.Text = "0 departments"
            Return
        End If

        Dim visibleCount As Integer = _departmentsTable.DefaultView.Count
        Dim totalCount As Integer = _departmentsTable.Rows.Count
        Dim hasSearch As Boolean = Not String.IsNullOrWhiteSpace(_searchTerm)

        If hasSearch AndAlso visibleCount <> totalCount Then
            DepartmentsCountTextBlock.Text = visibleCount.ToString() & " of " & totalCount.ToString() & " departments"
        Else
            DepartmentsCountTextBlock.Text = totalCount.ToString() & " departments"
        End If
    End Sub

    Private Sub OpenAddDepartmentButton_Click(sender As Object, e As RoutedEventArgs)
        BeginAddDepartment()
    End Sub

    Private Sub EditSelectedDepartmentButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        BeginEditDepartment(selectedRow)
    End Sub

    Private Sub DeleteSelectedDepartmentButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        DeleteDepartmentRow(selectedRow)
    End Sub

    Private Sub DepartmentsDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        RefreshDepartmentDetailsPanel()
    End Sub

    Private Sub BeginAddDepartment()
        OpenDepartmentForm(DepartmentFormMode.Add, "Add Department", "Create a new department record.", "Add Department", Nothing)
        DepartmentFormIdTextBox.Focus()
    End Sub

    Private Sub BeginEditDepartment(row As DataRow)
        If row Is Nothing Then
            Return
        End If

        OpenDepartmentForm(DepartmentFormMode.Edit, "Edit Department", "Update department details.", "Save Changes", row)
        DepartmentFormNameTextBox.Focus()
    End Sub

    Private Sub SaveDepartmentFormButton_Click(sender As Object, e As RoutedEventArgs)
        Dim formValues As New DepartmentFormValues()
        If Not TryReadDepartmentForm(formValues) Then
            Return
        End If

        EnsureDepartmentsTableLoaded()
        Dim targetRow As DataRow = ResolveTargetRowForSave(formValues.DepartmentId)
        If targetRow Is Nothing Then
            Return
        End If

        WriteDepartmentValues(targetRow, formValues)
        If _activeFormMode = DepartmentFormMode.Add Then
            _departmentsTable.Rows.Add(targetRow)
        End If

        _departmentsTable.AcceptChanges()
        If Not PersistDepartmentsToStorage() Then
            Return
        End If

        ApplyDepartmentsFilter()
        SelectDepartmentById(formValues.DepartmentId)
        HideDepartmentForm()
    End Sub

    Private Sub OpenDepartmentForm(mode As DepartmentFormMode,
                                   title As String,
                                   subtitle As String,
                                   actionText As String,
                                   row As DataRow)
        _activeFormMode = mode
        DepartmentFormTitleTextBlock.Text = title
        DepartmentFormSubtitleTextBlock.Text = subtitle
        SaveDepartmentFormButton.Content = actionText
        DepartmentFormIdTextBox.IsReadOnly = False

        If row Is Nothing Then
            _editingDepartmentOriginalId = String.Empty
            ClearDepartmentFormInputs()
        Else
            _editingDepartmentOriginalId = ReadRowValue(row, "Department ID")
            PopulateDepartmentForm(row)
        End If

        ShowDepartmentForm()
    End Sub

    Private Sub EnsureDepartmentsTableLoaded()
        If _departmentsTable Is Nothing Then
            SetDepartmentsTable(Nothing)
        End If
    End Sub

    Private Function ResolveTargetRowForSave(departmentId As String) As DataRow
        Select Case _activeFormMode
            Case DepartmentFormMode.Add
                If FindDepartmentRowById(departmentId) IsNot Nothing Then
                    MessageBox.Show("Department ID already exists.", "Duplicate Department ID", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return _departmentsTable.NewRow()

            Case DepartmentFormMode.Edit
                Dim targetRow As DataRow = FindDepartmentRowById(_editingDepartmentOriginalId)
                If targetRow Is Nothing Then
                    MessageBox.Show("The selected department no longer exists.", "Edit Department", MessageBoxButton.OK, MessageBoxImage.Information)
                    HideDepartmentForm()
                    Return Nothing
                End If

                If Not String.Equals(_editingDepartmentOriginalId, departmentId, StringComparison.OrdinalIgnoreCase) AndAlso
                   FindDepartmentRowById(departmentId) IsNot Nothing Then
                    MessageBox.Show("Department ID already exists.", "Duplicate Department ID", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return targetRow
        End Select

        Return Nothing
    End Function

    Private Sub WriteDepartmentValues(row As DataRow, values As DepartmentFormValues)
        row("Department ID") = values.DepartmentId
        row("Department Name") = values.DepartmentName
        row("Head") = values.Head
    End Sub

    Private Sub CancelDepartmentFormButton_Click(sender As Object, e As RoutedEventArgs)
        HideDepartmentForm()
    End Sub

    Private Sub DeleteDepartmentRow(row As DataRow)
        If row Is Nothing OrElse row.RowState = DataRowState.Deleted Then
            Return
        End If

        Dim departmentName As String = ReadRowValue(row, "Department Name")
        Dim departmentId As String = ReadRowValue(row, "Department ID")
        Dim recordLabel As String = If(String.IsNullOrWhiteSpace(departmentName), departmentId, departmentName)

        Dim confirmation As MessageBoxResult = MessageBox.Show("Delete " & recordLabel & "?", "Delete Department", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        row.Delete()
        _departmentsTable.AcceptChanges()
        If Not PersistDepartmentsToStorage() Then
            Return
        End If

        ApplyDepartmentsFilter()

        If _activeFormMode = DepartmentFormMode.Edit AndAlso
           String.Equals(_editingDepartmentOriginalId, departmentId, StringComparison.OrdinalIgnoreCase) Then
            HideDepartmentForm()
        End If

        EnsureSelectedDepartmentForDetails()
        RefreshDepartmentDetailsPanel()
    End Sub

    Private Sub EnsureSelectedDepartmentForDetails()
        If DepartmentsDataGrid Is Nothing OrElse _departmentsTable Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow IsNot Nothing Then
            Return
        End If

        If _departmentsTable.DefaultView.Count > 0 Then
            DepartmentsDataGrid.SelectedItem = _departmentsTable.DefaultView(0)
        End If
    End Sub

    Private Function TryGetSelectedGridRow() As DataRow
        Dim selectedRowView As DataRowView = TryCast(DepartmentsDataGrid.SelectedItem, DataRowView)
        If selectedRowView Is Nothing OrElse selectedRowView.Row Is Nothing Then
            Return Nothing
        End If

        If selectedRowView.Row.RowState = DataRowState.Deleted Then
            Return Nothing
        End If

        Return selectedRowView.Row
    End Function

    Private Sub RefreshDepartmentDetailsPanel()
        If DepartmentDetailsIdTextBlock Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        Dim hasSelection As Boolean = selectedRow IsNot Nothing

        If EditSelectedDepartmentButton IsNot Nothing Then
            EditSelectedDepartmentButton.IsEnabled = hasSelection
        End If

        If DeleteSelectedDepartmentButton IsNot Nothing Then
            DeleteSelectedDepartmentButton.IsEnabled = hasSelection
        End If

        If Not hasSelection Then
            If DepartmentDetailsSubtitleTextBlock IsNot Nothing Then
                DepartmentDetailsSubtitleTextBlock.Text = "Select a department from the table."
            End If

            SetDetailsValue(DepartmentDetailsIdTextBlock, String.Empty)
            SetDetailsValue(DepartmentDetailsNameTextBlock, String.Empty)
            SetDetailsValue(DepartmentDetailsHeadTextBlock, String.Empty)
            Return
        End If

        If DepartmentDetailsSubtitleTextBlock IsNot Nothing Then
            DepartmentDetailsSubtitleTextBlock.Text = "Selected department record."
        End If

        SetDetailsValue(DepartmentDetailsIdTextBlock, ReadRowValue(selectedRow, "Department ID"))
        SetDetailsValue(DepartmentDetailsNameTextBlock, ReadRowValue(selectedRow, "Department Name"))
        SetDetailsValue(DepartmentDetailsHeadTextBlock, ReadRowValue(selectedRow, "Head"))
    End Sub

    Private Sub SetDetailsValue(target As TextBlock, value As String)
        If target Is Nothing Then
            Return
        End If

        Dim normalizedValue As String = If(value, String.Empty).Trim()
        target.Text = If(String.IsNullOrWhiteSpace(normalizedValue), "--", normalizedValue)
    End Sub

    Private Sub ShowDepartmentForm()
        If DepartmentFormPanel IsNot Nothing Then
            DepartmentFormPanel.Visibility = Visibility.Visible
        End If

        If DepartmentDetailsPanel IsNot Nothing Then
            DepartmentDetailsPanel.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub HideDepartmentForm()
        _activeFormMode = DepartmentFormMode.Add
        _editingDepartmentOriginalId = String.Empty
        DepartmentFormTitleTextBlock.Text = "Add Department"
        DepartmentFormSubtitleTextBlock.Text = "Create a new department record."
        SaveDepartmentFormButton.Content = "Add Department"
        DepartmentFormIdTextBox.IsReadOnly = False
        ClearDepartmentFormInputs()

        If DepartmentFormPanel IsNot Nothing Then
            DepartmentFormPanel.Visibility = Visibility.Collapsed
        End If

        If DepartmentDetailsPanel IsNot Nothing Then
            DepartmentDetailsPanel.Visibility = Visibility.Visible
        End If

        RefreshDepartmentDetailsPanel()
    End Sub

    Private Sub PopulateDepartmentForm(row As DataRow)
        DepartmentFormIdTextBox.Text = ReadRowValue(row, "Department ID")
        DepartmentFormNameTextBox.Text = ReadRowValue(row, "Department Name")
        DepartmentFormHeadTextBox.Text = ReadRowValue(row, "Head")
    End Sub

    Private Sub ClearDepartmentFormInputs()
        DepartmentFormIdTextBox.Text = String.Empty
        DepartmentFormNameTextBox.Text = String.Empty
        DepartmentFormHeadTextBox.Text = String.Empty
    End Sub

    Private Function TryReadDepartmentForm(ByRef values As DepartmentFormValues) As Boolean
        values.DepartmentId = If(DepartmentFormIdTextBox.Text, String.Empty).Trim()
        values.DepartmentName = If(DepartmentFormNameTextBox.Text, String.Empty).Trim()
        values.Head = If(DepartmentFormHeadTextBox.Text, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(values.DepartmentId) Then
            MessageBox.Show("Department ID is required.", "Department Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.DepartmentName) Then
            MessageBox.Show("Department Name is required.", "Department Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        Return True
    End Function

    Private Function FindDepartmentRowById(departmentId As String) As DataRow
        If _departmentsTable Is Nothing Then
            Return Nothing
        End If

        Dim normalizedDepartmentId As String = If(departmentId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedDepartmentId) Then
            Return Nothing
        End If

        For Each row As DataRow In _departmentsTable.Rows
            If row.RowState = DataRowState.Deleted Then
                Continue For
            End If

            Dim candidateId As String = ReadRowValue(row, "Department ID")
            If String.Equals(candidateId, normalizedDepartmentId, StringComparison.OrdinalIgnoreCase) Then
                Return row
            End If
        Next

        Return Nothing
    End Function

    Private Sub SelectDepartmentById(departmentId As String)
        If _departmentsTable Is Nothing Then
            Return
        End If

        Dim normalizedDepartmentId As String = If(departmentId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedDepartmentId) Then
            Return
        End If

        For Each rowView As DataRowView In _departmentsTable.DefaultView
            Dim candidateId As String = If(rowView("Department ID"), String.Empty).ToString().Trim()
            If String.Equals(candidateId, normalizedDepartmentId, StringComparison.OrdinalIgnoreCase) Then
                DepartmentsDataGrid.SelectedItem = rowView
                DepartmentsDataGrid.ScrollIntoView(rowView)
                RefreshDepartmentDetailsPanel()
                Return
            End If
        Next

        RefreshDepartmentDetailsPanel()
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
