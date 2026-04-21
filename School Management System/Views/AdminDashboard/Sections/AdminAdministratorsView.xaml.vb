Imports System.Data
Imports System.IO
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Microsoft.Win32
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class AdminAdministratorsView
    Private Enum AdministratorFormMode
        Add
        Edit
    End Enum

    Private Structure AdministratorFormValues
        Public AdministratorId As String
        Public FirstName As String
        Public MiddleName As String
        Public LastName As String
        Public RoleTitle As String
        Public Email As String
        Public Password As String
        Public PhotoPath As String
    End Structure

    Private _administratorsTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As AdministratorFormMode = AdministratorFormMode.Add
    Private _editingAdministratorOriginalId As String = String.Empty
    Private ReadOnly _administratorManagementService As New AdministratorManagementService()

    Public Sub New()
        InitializeComponent()
        LoadAdministratorsTable()
        HideAdministratorForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyAdministratorsFilter()
    End Sub

    Public Sub RefreshData()
        LoadAdministratorsTable(GetSelectedAdministratorId())
    End Sub

    Private Sub LoadAdministratorsTable(Optional administratorIdToSelect As String = "")
        Dim result = _administratorManagementService.GetAdministrators()

        If Not result.IsSuccess Then
            SetAdministratorsTable(CreateEmptyAdministratorsTable())
            MessageBox.Show(result.Message,
                            "Administrators",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return
        End If

        SetAdministratorsTable(BuildAdministratorsTable(result.Data))

        If Not String.IsNullOrWhiteSpace(administratorIdToSelect) Then
            SelectAdministratorById(administratorIdToSelect)
        End If
    End Sub

    Public Sub SetAdministratorsTable(table As DataTable)
        _administratorsTable = If(table, CreateEmptyAdministratorsTable())
        AdministratorsDataGrid.ItemsSource = _administratorsTable.DefaultView
        ApplyAdministratorsFilter()
        UpdateAdministratorsCount()
        EnsureSelectedAdministratorForDetails()
        RefreshAdministratorDetailsPanel()
    End Sub

    Private Function BuildAdministratorsTable(records As IEnumerable(Of AdministratorRecord)) As DataTable
        Dim table As DataTable = CreateEmptyAdministratorsTable()

        If records Is Nothing Then
            Return table
        End If

        For Each record As AdministratorRecord In records
            Dim row As DataRow = table.NewRow()
            row("Administrator ID") = record.AdministratorCode
            row("Full Name") = record.FullName
            row("First Name") = record.FirstName
            row("Middle Name") = record.MiddleName
            row("Last Name") = record.LastName
            row("Role") = record.RoleTitle
            row("Email") = record.Email
            row("Photo Path") = record.PhotoPath
            table.Rows.Add(row)
        Next

        Return table
    End Function

    Private Function CreateEmptyAdministratorsTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Administrator ID", GetType(String))
        table.Columns.Add("Full Name", GetType(String))
        table.Columns.Add("First Name", GetType(String))
        table.Columns.Add("Middle Name", GetType(String))
        table.Columns.Add("Last Name", GetType(String))
        table.Columns.Add("Role", GetType(String))
        table.Columns.Add("Email", GetType(String))
        table.Columns.Add("Photo Path", GetType(String))
        Return table
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
            AdministratorsCountTextBlock.Text =
                visibleCount.ToString() & " of " & totalCount.ToString() & " administrators"
            Return
        End If

        AdministratorsCountTextBlock.Text = totalCount.ToString() & " administrators"
    End Sub

    Private Sub OpenAddAdministratorButton_Click(sender As Object, e As RoutedEventArgs)
        OpenAdministratorForm(AdministratorFormMode.Add,
                              "Add Administrator",
                              "Create a new administrator record.",
                              "Add Administrator",
                              Nothing)
        AdministratorFormAdministratorIdTextBox.Focus()
    End Sub

    Private Sub EditSelectedAdministratorButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        OpenAdministratorForm(AdministratorFormMode.Edit,
                              "Edit Administrator",
                              "Update administrator details.",
                              "Save Changes",
                              selectedRow)
        AdministratorFormFirstNameTextBox.Focus()
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

    Private Sub NestedPanelScrollViewer_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)
        Dim activeScrollViewer As ScrollViewer = TryCast(sender, ScrollViewer)
        If activeScrollViewer Is Nothing OrElse CanScrollViewerConsumeWheel(activeScrollViewer, e.Delta) Then
            Return
        End If

        ForwardWheelToRootScrollViewer(sender, e)
    End Sub

    Private Sub AdministratorsDataGrid_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)
        Dim activeScrollViewer As ScrollViewer =
            FindDescendantScrollViewer(TryCast(sender, DependencyObject))

        If activeScrollViewer IsNot Nothing AndAlso
           CanScrollViewerConsumeWheel(activeScrollViewer, e.Delta) Then
            Return
        End If

        ForwardWheelToRootScrollViewer(sender, e)
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
        UpdateImageControlSource(AdministratorFormPhotoPreviewImage, AdministratorFormPhotoPathTextBox.Text)
    End Sub

    Private Sub SaveAdministratorFormButton_Click(sender As Object, e As RoutedEventArgs)
        Dim formValues As New AdministratorFormValues()
        If Not TryReadAdministratorForm(formValues) Then
            Return
        End If

        Dim request As New AdministratorSaveRequest() With {
            .OriginalAdministratorCode = _editingAdministratorOriginalId,
            .AdministratorCode = formValues.AdministratorId,
            .FirstName = formValues.FirstName,
            .MiddleName = formValues.MiddleName,
            .LastName = formValues.LastName,
            .RoleTitle = formValues.RoleTitle,
            .Email = formValues.Email,
            .Password = formValues.Password,
            .PhotoPath = formValues.PhotoPath
        }

        Dim isAddMode As Boolean = _activeFormMode = AdministratorFormMode.Add
        Dim result =
            If(isAddMode,
               _administratorManagementService.CreateAdministrator(request),
               _administratorManagementService.UpdateAdministrator(request))

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Administrators",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        LoadAdministratorsTable(result.Data.AdministratorCode)
        HideAdministratorForm()

        If isAddMode AndAlso Not String.IsNullOrWhiteSpace(result.Message) Then
            MessageBox.Show(result.Message,
                            "Administrators",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
        End If
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

        Dim confirmation As MessageBoxResult =
            MessageBox.Show("Delete " & recordLabel & "?",
                            "Delete Administrator",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        Dim result = _administratorManagementService.DeleteAdministrator(administratorId)
        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Administrators",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        If _activeFormMode = AdministratorFormMode.Edit AndAlso
           String.Equals(_editingAdministratorOriginalId,
                         administratorId,
                         StringComparison.OrdinalIgnoreCase) Then
            HideAdministratorForm()
        End If

        LoadAdministratorsTable()
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

        If row Is Nothing Then
            _editingAdministratorOriginalId = String.Empty
            ClearAdministratorFormInputs()
        Else
            _editingAdministratorOriginalId = ReadRowValue(row, "Administrator ID")
            PopulateAdministratorForm(row)
        End If

        ShowAdministratorForm()
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

    Private Function GetSelectedAdministratorId() As String
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return String.Empty
        End If

        Return ReadRowValue(selectedRow, "Administrator ID")
    End Function

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

    Private Function CanScrollViewerConsumeWheel(scrollViewer As ScrollViewer,
                                                 delta As Integer) As Boolean
        If scrollViewer Is Nothing OrElse delta = 0 Then
            Return False
        End If

        If scrollViewer.ScrollableHeight <= 0 Then
            Return False
        End If

        If delta > 0 Then
            Return scrollViewer.VerticalOffset > 0
        End If

        Return scrollViewer.VerticalOffset < scrollViewer.ScrollableHeight
    End Function

    Private Sub ForwardWheelToRootScrollViewer(source As Object, e As MouseWheelEventArgs)
        If e Is Nothing OrElse AdministratorsRootScrollViewer Is Nothing Then
            Return
        End If

        e.Handled = True

        Dim forwardedEventArgs As New MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
        forwardedEventArgs.RoutedEvent = UIElement.MouseWheelEvent
        forwardedEventArgs.Source = source
        AdministratorsRootScrollViewer.RaiseEvent(forwardedEventArgs)
    End Sub

    Private Function FindDescendantScrollViewer(root As DependencyObject) As ScrollViewer
        If root Is Nothing Then
            Return Nothing
        End If

        If TypeOf root Is ScrollViewer Then
            Return DirectCast(root, ScrollViewer)
        End If

        Dim childCount As Integer = VisualTreeHelper.GetChildrenCount(root)
        For index As Integer = 0 To childCount - 1
            Dim child As DependencyObject = VisualTreeHelper.GetChild(root, index)
            Dim scrollViewer As ScrollViewer = FindDescendantScrollViewer(child)

            If scrollViewer IsNot Nothing Then
                Return scrollViewer
            End If
        Next

        Return Nothing
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
            AdministratorDetailsSubtitleTextBlock.Text = "Select an administrator from the table."
            SetDetailsValue(AdministratorDetailsAdministratorIdTextBlock, String.Empty)
            SetDetailsValue(AdministratorDetailsFullNameTextBlock, String.Empty)
            SetDetailsValue(AdministratorDetailsRoleTextBlock, String.Empty)
            SetDetailsValue(AdministratorDetailsEmailTextBlock, String.Empty)
            UpdateImageControlSource(AdministratorDetailsPhotoImage, String.Empty)
            Return
        End If

        AdministratorDetailsSubtitleTextBlock.Text = "Selected administrator record."
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
        AdministratorFormPanel.Visibility = Visibility.Visible
        AdministratorDetailsPanel.Visibility = Visibility.Collapsed
    End Sub

    Private Sub HideAdministratorForm()
        _activeFormMode = AdministratorFormMode.Add
        _editingAdministratorOriginalId = String.Empty
        AdministratorFormTitleTextBlock.Text = "Add Administrator"
        AdministratorFormSubtitleTextBlock.Text = "Create a new administrator record."
        SaveAdministratorFormButton.Content = "Add Administrator"
        ClearAdministratorFormInputs()
        AdministratorFormPanel.Visibility = Visibility.Collapsed
        AdministratorDetailsPanel.Visibility = Visibility.Visible
        RefreshAdministratorDetailsPanel()
    End Sub

    Private Sub PopulateAdministratorForm(row As DataRow)
        AdministratorFormAdministratorIdTextBox.Text = ReadRowValue(row, "Administrator ID")
        AdministratorFormFirstNameTextBox.Text = ReadRowValue(row, "First Name")
        AdministratorFormMiddleNameTextBox.Text = ReadRowValue(row, "Middle Name")
        AdministratorFormLastNameTextBox.Text = ReadRowValue(row, "Last Name")
        AdministratorFormRoleTextBox.Text = ReadRowValue(row, "Role")
        AdministratorFormEmailTextBox.Text = ReadRowValue(row, "Email")
        AdministratorFormPasswordTextBox.Text = String.Empty
        AdministratorFormPhotoPathTextBox.Text = ReadRowValue(row, "Photo Path")
        UpdateImageControlSource(AdministratorFormPhotoPreviewImage, AdministratorFormPhotoPathTextBox.Text)
    End Sub

    Private Sub ClearAdministratorFormInputs()
        AdministratorFormAdministratorIdTextBox.Text = String.Empty
        AdministratorFormFirstNameTextBox.Text = String.Empty
        AdministratorFormMiddleNameTextBox.Text = String.Empty
        AdministratorFormLastNameTextBox.Text = String.Empty
        AdministratorFormRoleTextBox.Text = String.Empty
        AdministratorFormEmailTextBox.Text = String.Empty
        AdministratorFormPasswordTextBox.Text = String.Empty
        AdministratorFormPhotoPathTextBox.Text = String.Empty
        UpdateImageControlSource(AdministratorFormPhotoPreviewImage, String.Empty)
    End Sub

    Private Function TryReadAdministratorForm(ByRef values As AdministratorFormValues) As Boolean
        values.AdministratorId = If(AdministratorFormAdministratorIdTextBox.Text, String.Empty).Trim()
        values.FirstName = If(AdministratorFormFirstNameTextBox.Text, String.Empty).Trim()
        values.MiddleName = If(AdministratorFormMiddleNameTextBox.Text, String.Empty).Trim()
        values.LastName = If(AdministratorFormLastNameTextBox.Text, String.Empty).Trim()
        values.RoleTitle = If(AdministratorFormRoleTextBox.Text, String.Empty).Trim()
        values.Email = If(AdministratorFormEmailTextBox.Text, String.Empty).Trim()
        values.Password = If(AdministratorFormPasswordTextBox.Text, String.Empty).Trim()
        values.PhotoPath = If(AdministratorFormPhotoPathTextBox.Text, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(values.AdministratorId) Then
            MessageBox.Show("Administrator ID is required.",
                            "Administrator Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.FirstName) Then
            MessageBox.Show("First Name is required.",
                            "Administrator Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.LastName) Then
            MessageBox.Show("Last Name is required.",
                            "Administrator Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.Email) Then
            MessageBox.Show("Email is required.",
                            "Administrator Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        Return True
    End Function

    Private Sub SelectAdministratorById(administratorId As String)
        If _administratorsTable Is Nothing Then
            Return
        End If

        Dim normalizedAdministratorId As String = If(administratorId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedAdministratorId) Then
            RefreshAdministratorDetailsPanel()
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

    Private Sub UpdateImageControlSource(targetImage As System.Windows.Controls.Image,
                                         imagePath As String)
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
