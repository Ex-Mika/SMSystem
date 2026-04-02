Imports System.Collections.Generic
Imports System.Data
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

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

    Private _departmentsTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As DepartmentFormMode = DepartmentFormMode.Add
    Private _editingDepartmentOriginalId As String = String.Empty
    Private ReadOnly _departmentManagementService As New DepartmentManagementService()

    Public Sub New()
        InitializeComponent()
        LoadDepartmentsTable()
        HideDepartmentForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyDepartmentsFilter()
    End Sub

    Public Sub RefreshData()
        LoadDepartmentsTable(GetSelectedDepartmentId())
    End Sub

    Public Sub SetDepartmentsTable(table As DataTable)
        _departmentsTable = If(table, CreateEmptyDepartmentsTable())
        DepartmentsDataGrid.ItemsSource = _departmentsTable.DefaultView
        ApplyDepartmentsFilter()
        UpdateDepartmentsCount()
        EnsureSelectedDepartmentForDetails()
        RefreshDepartmentDetailsPanel()
    End Sub

    Private Sub LoadDepartmentsTable(Optional departmentIdToSelect As String = "")
        Dim result = _departmentManagementService.GetDepartments()

        If Not result.IsSuccess Then
            SetDepartmentsTable(CreateEmptyDepartmentsTable())
            MessageBox.Show(result.Message,
                            "Departments",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return
        End If

        SetDepartmentsTable(BuildDepartmentsTable(result.Data))

        If Not String.IsNullOrWhiteSpace(departmentIdToSelect) Then
            SelectDepartmentById(departmentIdToSelect)
        End If
    End Sub

    Private Function BuildDepartmentsTable(records As IEnumerable(Of DepartmentRecord)) As DataTable
        Dim table As DataTable = CreateEmptyDepartmentsTable()

        If records Is Nothing Then
            Return table
        End If

        For Each record As DepartmentRecord In records
            Dim row As DataRow = table.NewRow()
            row("Department ID") = record.DepartmentCode
            row("Department Name") = record.DepartmentName
            row("Head") = record.HeadName
            table.Rows.Add(row)
        Next

        Return table
    End Function

    Private Function CreateEmptyDepartmentsTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Department ID", GetType(String))
        table.Columns.Add("Department Name", GetType(String))
        table.Columns.Add("Head", GetType(String))
        Return table
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
            DepartmentsCountTextBlock.Text =
                visibleCount.ToString() & " of " & totalCount.ToString() & " departments"
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
        OpenDepartmentForm(DepartmentFormMode.Add,
                           "Add Department",
                           "Create a new department record.",
                           "Add Department",
                           Nothing)
        DepartmentFormIdTextBox.Focus()
    End Sub

    Private Sub BeginEditDepartment(row As DataRow)
        If row Is Nothing Then
            Return
        End If

        OpenDepartmentForm(DepartmentFormMode.Edit,
                           "Edit Department",
                           "Update department details.",
                           "Save Changes",
                           row)
        DepartmentFormNameTextBox.Focus()
    End Sub

    Private Sub SaveDepartmentFormButton_Click(sender As Object, e As RoutedEventArgs)
        Dim formValues As New DepartmentFormValues()
        If Not TryReadDepartmentForm(formValues) Then
            Return
        End If

        Dim request As New DepartmentSaveRequest() With {
            .OriginalDepartmentCode = _editingDepartmentOriginalId,
            .DepartmentCode = formValues.DepartmentId,
            .DepartmentName = formValues.DepartmentName,
            .HeadName = formValues.Head
        }

        Dim isAddMode As Boolean = _activeFormMode = DepartmentFormMode.Add
        Dim result =
            If(isAddMode,
               _departmentManagementService.CreateDepartment(request),
               _departmentManagementService.UpdateDepartment(request))

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Departments",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        LoadDepartmentsTable(result.Data.DepartmentCode)
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

    Private Sub CancelDepartmentFormButton_Click(sender As Object, e As RoutedEventArgs)
        HideDepartmentForm()
    End Sub

    Private Sub DeleteDepartmentRow(row As DataRow)
        If row Is Nothing OrElse row.RowState = DataRowState.Deleted Then
            Return
        End If

        Dim departmentName As String = ReadRowValue(row, "Department Name")
        Dim departmentId As String = ReadRowValue(row, "Department ID")
        Dim recordLabel As String =
            If(String.IsNullOrWhiteSpace(departmentName), departmentId, departmentName)

        Dim confirmation As MessageBoxResult =
            MessageBox.Show("Delete " & recordLabel & "?",
                            "Delete Department",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        Dim result = _departmentManagementService.DeleteDepartment(departmentId)
        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Departments",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        If _activeFormMode = DepartmentFormMode.Edit AndAlso
           String.Equals(_editingDepartmentOriginalId,
                         departmentId,
                         StringComparison.OrdinalIgnoreCase) Then
            HideDepartmentForm()
        End If

        LoadDepartmentsTable()
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

    Private Function GetSelectedDepartmentId() As String
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return String.Empty
        End If

        Return ReadRowValue(selectedRow, "Department ID")
    End Function

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
            MessageBox.Show("Department ID is required.",
                            "Department Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.DepartmentName) Then
            MessageBox.Show("Department Name is required.",
                            "Department Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        Return True
    End Function

    Private Sub SelectDepartmentById(departmentId As String)
        If _departmentsTable Is Nothing Then
            Return
        End If

        Dim normalizedDepartmentId As String = If(departmentId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedDepartmentId) Then
            RefreshDepartmentDetailsPanel()
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
