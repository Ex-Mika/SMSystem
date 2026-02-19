Imports System.Data
Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json
Imports System.Windows.Media.Imaging
Imports Microsoft.Win32

Class AdminAdministratorsView
    Private Enum AdministratorFormMode
        Add
        Edit
    End Enum

    Private Structure AdministratorFormValues
        Public AdministratorId As String
        Public FullName As String
        Public Role As String
        Public Email As String
        Public PhotoPath As String
    End Structure

    Private Class AdministratorStorageRecord
        Public Property AdministratorId As String
        Public Property FullName As String
        Public Property Role As String
        Public Property Email As String
        Public Property PhotoPath As String
    End Class

    Private _administratorsTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As AdministratorFormMode = AdministratorFormMode.Add
    Private _editingAdministratorOriginalId As String = String.Empty
    Private ReadOnly _administratorsStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "administrators.json")
    Private ReadOnly _administratorStorageJsonOptions As New JsonSerializerOptions() With {
        .WriteIndented = True
    }

    Public Sub New()
        InitializeComponent()
        LoadAdministratorsTable()
        HideAdministratorForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyAdministratorsFilter()
    End Sub

    Public Sub SetAdministratorsTable(table As DataTable)
        _administratorsTable = NormalizeAdministratorsTable(table)
        AdministratorsDataGrid.ItemsSource = _administratorsTable.DefaultView
        ApplyAdministratorsFilter()
        UpdateAdministratorsCount()
        EnsureSelectedAdministratorForDetails()
        RefreshAdministratorDetailsPanel()
    End Sub

    Private Sub LoadAdministratorsTable()
        Dim administratorsTable As DataTable = FetchAdministratorsTable()
        SetAdministratorsTable(administratorsTable)
    End Sub

    Private Function FetchAdministratorsTable() As DataTable
        Return ReadAdministratorsFromStorage()
    End Function

    Private Function NormalizeAdministratorsTable(source As DataTable) As DataTable
        Dim table As DataTable = If(source Is Nothing, CreateEmptyAdministratorsTable(), source.Copy())

        Dim administratorIdColumn As DataColumn = EnsureAdministratorsColumn(table, "Administrator ID", "AdministratorId", "Administrator_ID", "ID", "Administrator Number")
        Dim fullNameColumn As DataColumn = EnsureAdministratorsColumn(table, "Full Name", "FullName", "Name", "Administrator Name")
        Dim roleColumn As DataColumn = EnsureAdministratorsColumn(table, "Role", "Program", "Role Name")
        Dim emailColumn As DataColumn = EnsureAdministratorsColumn(table, "Email", "Class", "Section", "Block")
        Dim photoPathColumn As DataColumn = EnsureAdministratorsColumn(table, "Photo Path", "PhotoPath", "Photo", "Image", "ImagePath", "Avatar")

        RemoveAdministratorsColumns(table,
                              "Status",
                              "Employment Status",
                              "State",
                              "Enrollment Status",
                              "Photo",
                              "Image",
                              "ImagePath",
                              "Avatar")

        administratorIdColumn.SetOrdinal(0)
        fullNameColumn.SetOrdinal(1)
        roleColumn.SetOrdinal(2)
        emailColumn.SetOrdinal(3)
        photoPathColumn.SetOrdinal(4)

        For Each row As DataRow In table.Rows
            If row.IsNull("Administrator ID") Then
                row("Administrator ID") = String.Empty
            Else
                row("Administrator ID") = row("Administrator ID").ToString().Trim()
            End If

            If row.IsNull("Full Name") Then
                row("Full Name") = String.Empty
            Else
                row("Full Name") = row("Full Name").ToString().Trim()
            End If

            If row.IsNull("Role") Then
                row("Role") = String.Empty
            Else
                row("Role") = row("Role").ToString().Trim()
            End If

            If row.IsNull("Email") Then
                row("Email") = String.Empty
            Else
                row("Email") = row("Email").ToString().Trim()
            End If

            If row.IsNull("Photo Path") Then
                row("Photo Path") = String.Empty
            Else
                row("Photo Path") = row("Photo Path").ToString().Trim()
            End If
        Next

        Return table
    End Function

    Private Sub RemoveAdministratorsColumns(table As DataTable, ParamArray columnNames() As String)
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

    Private Function EnsureAdministratorsColumn(table As DataTable, targetColumnName As String, ParamArray aliases() As String) As DataColumn
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

    Private Function CreateEmptyAdministratorsTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Administrator ID", GetType(String))
        table.Columns.Add("Full Name", GetType(String))
        table.Columns.Add("Role", GetType(String))
        table.Columns.Add("Email", GetType(String))
        table.Columns.Add("Photo Path", GetType(String))
        Return table
    End Function

    Private Function ReadAdministratorsFromStorage() As DataTable
        Dim table As DataTable = CreateEmptyAdministratorsTable()
        If Not File.Exists(_administratorsStoragePath) Then
            Return table
        End If

        Try
            Dim json As String = File.ReadAllText(_administratorsStoragePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return table
            End If

            Dim records As List(Of AdministratorStorageRecord) =
                JsonSerializer.Deserialize(Of List(Of AdministratorStorageRecord))(json, _administratorStorageJsonOptions)
            If records Is Nothing Then
                Return table
            End If

            For Each record As AdministratorStorageRecord In records
                Dim row As DataRow = table.NewRow()
                row("Administrator ID") = If(record.AdministratorId, String.Empty).Trim()
                row("Full Name") = If(record.FullName, String.Empty).Trim()
                row("Role") = If(record.Role, String.Empty).Trim()
                row("Email") = If(record.Email, String.Empty).Trim()
                row("Photo Path") = If(record.PhotoPath, String.Empty).Trim()
                table.Rows.Add(row)
            Next
        Catch ex As Exception
            MessageBox.Show("Unable to load saved administrators data." & Environment.NewLine & ex.Message,
                            "Administrators",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
        End Try

        Return table
    End Function

    Private Function PersistAdministratorsToStorage() As Boolean
        If _administratorsTable Is Nothing Then
            Return True
        End If

        Try
            Dim storageDirectory As String = Path.GetDirectoryName(_administratorsStoragePath)
            If Not String.IsNullOrWhiteSpace(storageDirectory) Then
                Directory.CreateDirectory(storageDirectory)
            End If

            Dim records As New List(Of AdministratorStorageRecord)()
            For Each row As DataRow In _administratorsTable.Rows
                If row.RowState = DataRowState.Deleted Then
                    Continue For
                End If

                records.Add(New AdministratorStorageRecord With {
                    .AdministratorId = ReadRowValue(row, "Administrator ID"),
                    .FullName = ReadRowValue(row, "Full Name"),
                    .Role = ReadRowValue(row, "Role"),
                    .Email = ReadRowValue(row, "Email"),
                    .PhotoPath = ReadRowValue(row, "Photo Path")
                })
            Next

            Dim json As String = JsonSerializer.Serialize(records, _administratorStorageJsonOptions)
            File.WriteAllText(_administratorsStoragePath, json)
            Return True
        Catch ex As Exception
            MessageBox.Show("Unable to save administrators data." & Environment.NewLine & ex.Message,
                            "Administrators",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            Return False
        End Try
    End Function

    Private Sub ApplyAdministratorsFilter()
        If _administratorsTable Is Nothing Then
            UpdateAdministratorsCount()
            Return
        End If

        If String.IsNullOrWhiteSpace(_searchTerm) Then
            _administratorsTable.DefaultView.RowFilter = String.Empty
        Else
            Dim escapedTerm As String = EscapeLikeValue(_searchTerm)
            _administratorsTable.DefaultView.RowFilter =
                "[Administrator ID] LIKE '*" & escapedTerm & "*' OR " &
                "[Full Name] LIKE '*" & escapedTerm & "*' OR " &
                "[Role] LIKE '*" & escapedTerm & "*' OR " &
                "[Email] LIKE '*" & escapedTerm & "*'"
        End If

        UpdateAdministratorsCount()
        EnsureSelectedAdministratorForDetails()
        RefreshAdministratorDetailsPanel()
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

    Private Sub UpdateAdministratorsCount()
        If AdministratorsCountTextBlock Is Nothing Then
            Return
        End If

        If _administratorsTable Is Nothing Then
            AdministratorsCountTextBlock.Text = "0 administrators"
            Return
        End If

        Dim visibleCount As Integer = _administratorsTable.DefaultView.Count
        Dim totalCount As Integer = _administratorsTable.Rows.Count
        Dim hasSearch As Boolean = Not String.IsNullOrWhiteSpace(_searchTerm)

        If hasSearch AndAlso visibleCount <> totalCount Then
            AdministratorsCountTextBlock.Text = visibleCount.ToString() & " of " & totalCount.ToString() & " administrators"
        Else
            AdministratorsCountTextBlock.Text = totalCount.ToString() & " administrators"
        End If
    End Sub

    Private Sub OpenAddAdministratorButton_Click(sender As Object, e As RoutedEventArgs)
        BeginAddAdministrator()
    End Sub

    Private Sub EditSelectedAdministratorButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        BeginEditAdministrator(selectedRow)
    End Sub

    Private Sub DeleteSelectedAdministratorButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        DeleteAdministratorRow(selectedRow)
    End Sub

    Private Sub AdministratorsDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        RefreshAdministratorDetailsPanel()
    End Sub

    Private Sub BrowseAdministratorPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dialog As New OpenFileDialog() With {
            .Title = "Select Administrator Photo",
            .Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp|All Files|*.*",
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        If dialog.ShowDialog() = True Then
            AdministratorFormPhotoPathTextBox.Text = If(dialog.FileName, String.Empty).Trim()
        End If
    End Sub

    Private Sub ClearAdministratorPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        AdministratorFormPhotoPathTextBox.Text = String.Empty
    End Sub

    Private Sub AdministratorFormPhotoPathTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        UpdateImageControlSource(AdministratorFormPhotoPreviewImage, If(AdministratorFormPhotoPathTextBox.Text, String.Empty))
    End Sub

    Private Sub BeginAddAdministrator()
        OpenAdministratorForm(AdministratorFormMode.Add, "Add Administrator", "Create a new administrator record.", "Add Administrator", Nothing)
        AdministratorFormAdministratorIdTextBox.Focus()
    End Sub

    Private Sub BeginEditAdministrator(row As DataRow)
        If row Is Nothing Then
            Return
        End If

        OpenAdministratorForm(AdministratorFormMode.Edit, "Edit Administrator", "Update administrator details.", "Save Changes", row)
        AdministratorFormFirstNameTextBox.Focus()
    End Sub

    Private Sub SaveAdministratorFormButton_Click(sender As Object, e As RoutedEventArgs)
        Dim formValues As New AdministratorFormValues()
        If Not TryReadAdministratorForm(formValues) Then
            Return
        End If

        EnsureAdministratorsTableLoaded()
        Dim targetRow As DataRow = ResolveTargetRowForSave(formValues.AdministratorId)
        If targetRow Is Nothing Then
            Return
        End If

        WriteAdministratorValues(targetRow, formValues)
        If _activeFormMode = AdministratorFormMode.Add Then
            _administratorsTable.Rows.Add(targetRow)
        End If

        _administratorsTable.AcceptChanges()
        If Not PersistAdministratorsToStorage() Then
            Return
        End If

        ApplyAdministratorsFilter()
        SelectAdministratorById(formValues.AdministratorId)
        HideAdministratorForm()
    End Sub

    Private Sub OpenAdministratorForm(mode As AdministratorFormMode,
                                title As String,
                                subtitle As String,
                                actionText As String,
                                row As DataRow)
        _activeFormMode = mode
        AdministratorFormTitleTextBlock.Text = title
        AdministratorFormSubtitleTextBlock.Text = subtitle
        SaveAdministratorFormButton.Content = actionText
        AdministratorFormAdministratorIdTextBox.IsReadOnly = False

        If row Is Nothing Then
            _editingAdministratorOriginalId = String.Empty
            ClearAdministratorFormInputs()
        Else
            _editingAdministratorOriginalId = ReadRowValue(row, "Administrator ID")
            PopulateAdministratorForm(row)
        End If

        ShowAdministratorForm()
    End Sub

    Private Sub EnsureAdministratorsTableLoaded()
        If _administratorsTable Is Nothing Then
            SetAdministratorsTable(Nothing)
        End If
    End Sub

    Private Function ResolveTargetRowForSave(administratorId As String) As DataRow
        Select Case _activeFormMode
            Case AdministratorFormMode.Add
                If FindAdministratorRowById(administratorId) IsNot Nothing Then
                    MessageBox.Show("Administrator ID already exists.", "Duplicate Administrator ID", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return _administratorsTable.NewRow()

            Case AdministratorFormMode.Edit
                Dim targetRow As DataRow = FindAdministratorRowById(_editingAdministratorOriginalId)
                If targetRow Is Nothing Then
                    MessageBox.Show("The selected administrator no longer exists.", "Edit Administrator", MessageBoxButton.OK, MessageBoxImage.Information)
                    HideAdministratorForm()
                    Return Nothing
                End If

                If Not String.Equals(_editingAdministratorOriginalId, administratorId, StringComparison.OrdinalIgnoreCase) AndAlso
                   FindAdministratorRowById(administratorId) IsNot Nothing Then
                    MessageBox.Show("Administrator ID already exists.", "Duplicate Administrator ID", MessageBoxButton.OK, MessageBoxImage.Information)
                    Return Nothing
                End If

                Return targetRow
        End Select

        Return Nothing
    End Function

    Private Sub WriteAdministratorValues(row As DataRow, values As AdministratorFormValues)
        row("Administrator ID") = values.AdministratorId
        row("Full Name") = values.FullName
        row("Role") = values.Role
        row("Email") = values.Email
        row("Photo Path") = values.PhotoPath
    End Sub

    Private Sub CancelAdministratorFormButton_Click(sender As Object, e As RoutedEventArgs)
        HideAdministratorForm()
    End Sub

    Private Sub DeleteAdministratorRow(row As DataRow)
        If row Is Nothing OrElse row.RowState = DataRowState.Deleted Then
            Return
        End If

        Dim fullName As String = ReadRowValue(row, "Full Name")
        Dim administratorId As String = ReadRowValue(row, "Administrator ID")
        Dim recordLabel As String = If(String.IsNullOrWhiteSpace(fullName), administratorId, fullName)

        Dim confirmation As MessageBoxResult = MessageBox.Show("Delete " & recordLabel & "?", "Delete Administrator", MessageBoxButton.YesNo, MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        row.Delete()
        _administratorsTable.AcceptChanges()
        If Not PersistAdministratorsToStorage() Then
            Return
        End If

        ApplyAdministratorsFilter()

        If _activeFormMode = AdministratorFormMode.Edit AndAlso
           String.Equals(_editingAdministratorOriginalId, administratorId, StringComparison.OrdinalIgnoreCase) Then
            HideAdministratorForm()
        End If

        EnsureSelectedAdministratorForDetails()
        RefreshAdministratorDetailsPanel()
    End Sub

    Private Sub EnsureSelectedAdministratorForDetails()
        If AdministratorsDataGrid Is Nothing OrElse _administratorsTable Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow IsNot Nothing Then
            Return
        End If

        If _administratorsTable.DefaultView.Count > 0 Then
            AdministratorsDataGrid.SelectedItem = _administratorsTable.DefaultView(0)
        End If
    End Sub

    Private Function TryGetSelectedGridRow() As DataRow
        Dim selectedRowView As DataRowView = TryCast(AdministratorsDataGrid.SelectedItem, DataRowView)
        If selectedRowView Is Nothing OrElse selectedRowView.Row Is Nothing Then
            Return Nothing
        End If

        If selectedRowView.Row.RowState = DataRowState.Deleted Then
            Return Nothing
        End If

        Return selectedRowView.Row
    End Function

    Private Sub RefreshAdministratorDetailsPanel()
        If AdministratorDetailsAdministratorIdTextBlock Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        Dim hasSelection As Boolean = selectedRow IsNot Nothing

        If EditSelectedAdministratorButton IsNot Nothing Then
            EditSelectedAdministratorButton.IsEnabled = hasSelection
        End If

        If DeleteSelectedAdministratorButton IsNot Nothing Then
            DeleteSelectedAdministratorButton.IsEnabled = hasSelection
        End If

        If Not hasSelection Then
            If AdministratorDetailsSubtitleTextBlock IsNot Nothing Then
                AdministratorDetailsSubtitleTextBlock.Text = "Select an administrator from the table."
            End If

            SetDetailsValue(AdministratorDetailsAdministratorIdTextBlock, String.Empty)
            SetDetailsValue(AdministratorDetailsFullNameTextBlock, String.Empty)
            SetDetailsValue(AdministratorDetailsRoleTextBlock, String.Empty)
            SetDetailsValue(AdministratorDetailsEmailTextBlock, String.Empty)
            UpdateImageControlSource(AdministratorDetailsPhotoImage, String.Empty)
            Return
        End If

        If AdministratorDetailsSubtitleTextBlock IsNot Nothing Then
            AdministratorDetailsSubtitleTextBlock.Text = "Selected administrator record."
        End If

        SetDetailsValue(AdministratorDetailsAdministratorIdTextBlock, ReadRowValue(selectedRow, "Administrator ID"))
        SetDetailsValue(AdministratorDetailsFullNameTextBlock, ReadRowValue(selectedRow, "Full Name"))
        SetDetailsValue(AdministratorDetailsRoleTextBlock, ReadRowValue(selectedRow, "Role"))
        SetDetailsValue(AdministratorDetailsEmailTextBlock, ReadRowValue(selectedRow, "Email"))
        UpdateImageControlSource(AdministratorDetailsPhotoImage, ReadRowValue(selectedRow, "Photo Path"))
    End Sub

    Private Sub SetDetailsValue(target As TextBlock, value As String)
        If target Is Nothing Then
            Return
        End If

        Dim normalizedValue As String = If(value, String.Empty).Trim()
        target.Text = If(String.IsNullOrWhiteSpace(normalizedValue), "--", normalizedValue)
    End Sub

    Private Sub ShowAdministratorForm()
        If AdministratorFormPanel IsNot Nothing Then
            AdministratorFormPanel.Visibility = Visibility.Visible
        End If

        If AdministratorDetailsPanel IsNot Nothing Then
            AdministratorDetailsPanel.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub HideAdministratorForm()
        _activeFormMode = AdministratorFormMode.Add
        _editingAdministratorOriginalId = String.Empty
        AdministratorFormTitleTextBlock.Text = "Add Administrator"
        AdministratorFormSubtitleTextBlock.Text = "Create a new administrator record."
        SaveAdministratorFormButton.Content = "Add Administrator"
        AdministratorFormAdministratorIdTextBox.IsReadOnly = False
        ClearAdministratorFormInputs()

        If AdministratorFormPanel IsNot Nothing Then
            AdministratorFormPanel.Visibility = Visibility.Collapsed
        End If

        If AdministratorDetailsPanel IsNot Nothing Then
            AdministratorDetailsPanel.Visibility = Visibility.Visible
        End If

        RefreshAdministratorDetailsPanel()
    End Sub

    Private Sub PopulateAdministratorForm(row As DataRow)
        AdministratorFormAdministratorIdTextBox.Text = ReadRowValue(row, "Administrator ID")

        Dim fullName As String = ReadRowValue(row, "Full Name")
        Dim firstName As String = String.Empty
        Dim lastName As String = String.Empty
        Dim middleName As String = String.Empty

        SplitFullName(fullName, firstName, lastName, middleName)

        AdministratorFormFirstNameTextBox.Text = firstName
        AdministratorFormLastNameTextBox.Text = lastName
        AdministratorFormMiddleNameTextBox.Text = middleName
        AdministratorFormRoleTextBox.Text = ReadRowValue(row, "Role")
        AdministratorFormEmailTextBox.Text = ReadRowValue(row, "Email")
        AdministratorFormPhotoPathTextBox.Text = ReadRowValue(row, "Photo Path")
        UpdateImageControlSource(AdministratorFormPhotoPreviewImage, AdministratorFormPhotoPathTextBox.Text)
    End Sub

    Private Sub ClearAdministratorFormInputs()
        AdministratorFormAdministratorIdTextBox.Text = String.Empty
        AdministratorFormFirstNameTextBox.Text = String.Empty
        AdministratorFormLastNameTextBox.Text = String.Empty
        AdministratorFormMiddleNameTextBox.Text = String.Empty
        AdministratorFormRoleTextBox.Text = String.Empty
        AdministratorFormEmailTextBox.Text = String.Empty
        AdministratorFormPhotoPathTextBox.Text = String.Empty
        UpdateImageControlSource(AdministratorFormPhotoPreviewImage, String.Empty)
    End Sub

    Private Function TryReadAdministratorForm(ByRef values As AdministratorFormValues) As Boolean
        values.AdministratorId = If(AdministratorFormAdministratorIdTextBox.Text, String.Empty).Trim()

        Dim firstName As String = If(AdministratorFormFirstNameTextBox.Text, String.Empty).Trim()
        Dim lastName As String = If(AdministratorFormLastNameTextBox.Text, String.Empty).Trim()
        Dim middleName As String = If(AdministratorFormMiddleNameTextBox.Text, String.Empty).Trim()

        values.FullName = BuildFullName(firstName, lastName, middleName)
        values.Role = If(AdministratorFormRoleTextBox.Text, String.Empty).Trim()
        values.Email = If(AdministratorFormEmailTextBox.Text, String.Empty).Trim()
        values.PhotoPath = If(AdministratorFormPhotoPathTextBox.Text, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(values.AdministratorId) Then
            MessageBox.Show("Administrator ID is required.", "Administrator Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(firstName) Then
            MessageBox.Show("First Name is required.", "Administrator Form", MessageBoxButton.OK, MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(lastName) Then
            MessageBox.Show("Last Name is required.", "Administrator Form", MessageBoxButton.OK, MessageBoxImage.Information)
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

    Private Function FindAdministratorRowById(administratorId As String) As DataRow
        If _administratorsTable Is Nothing Then
            Return Nothing
        End If

        Dim normalizedAdministratorId As String = If(administratorId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedAdministratorId) Then
            Return Nothing
        End If

        For Each row As DataRow In _administratorsTable.Rows
            If row.RowState = DataRowState.Deleted Then
                Continue For
            End If

            Dim candidateId As String = ReadRowValue(row, "Administrator ID")
            If String.Equals(candidateId, normalizedAdministratorId, StringComparison.OrdinalIgnoreCase) Then
                Return row
            End If
        Next

        Return Nothing
    End Function

    Private Sub SelectAdministratorById(administratorId As String)
        If _administratorsTable Is Nothing Then
            Return
        End If

        Dim normalizedAdministratorId As String = If(administratorId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedAdministratorId) Then
            Return
        End If

        For Each rowView As DataRowView In _administratorsTable.DefaultView
            Dim candidateId As String = If(rowView("Administrator ID"), String.Empty).ToString().Trim()
            If String.Equals(candidateId, normalizedAdministratorId, StringComparison.OrdinalIgnoreCase) Then
                AdministratorsDataGrid.SelectedItem = rowView
                AdministratorsDataGrid.ScrollIntoView(rowView)
                RefreshAdministratorDetailsPanel()
                Return
            End If
        Next

        RefreshAdministratorDetailsPanel()
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

