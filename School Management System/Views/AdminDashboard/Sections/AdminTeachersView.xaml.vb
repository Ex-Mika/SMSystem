Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Windows.Media.Imaging
Imports Microsoft.Win32
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class AdminTeachersView
    Private Enum TeacherFormMode
        Add
        Edit
    End Enum

    Private Structure TeacherFormValues
        Public TeacherId As String
        Public FirstName As String
        Public MiddleName As String
        Public LastName As String
        Public Department As String
        Public Advisory As String
        Public PhotoPath As String
    End Structure

    Private _teachersTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As TeacherFormMode = TeacherFormMode.Add
    Private _editingTeacherOriginalId As String = String.Empty
    Private ReadOnly _teacherManagementService As New TeacherManagementService()

    Public Sub New()
        InitializeComponent()
        LoadTeachersTable()
        HideTeacherForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyTeachersFilter()
    End Sub

    Public Sub RefreshData()
        LoadTeachersTable(GetSelectedTeacherId())
    End Sub

    Private Sub LoadTeachersTable(Optional teacherIdToSelect As String = "")
        Dim result = _teacherManagementService.GetTeachers()

        If Not result.IsSuccess Then
            SetTeachersTable(CreateEmptyTeachersTable())
            MessageBox.Show(result.Message,
                            "Teachers",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return
        End If

        SetTeachersTable(BuildTeachersTable(result.Data))

        If Not String.IsNullOrWhiteSpace(teacherIdToSelect) Then
            SelectTeacherById(teacherIdToSelect)
        End If
    End Sub

    Public Sub SetTeachersTable(table As DataTable)
        _teachersTable = If(table, CreateEmptyTeachersTable())
        TeachersDataGrid.ItemsSource = _teachersTable.DefaultView
        ApplyTeachersFilter()
        UpdateTeachersCount()
        EnsureSelectedTeacherForDetails()
        RefreshTeacherDetailsPanel()
    End Sub

    Private Function BuildTeachersTable(records As IEnumerable(Of TeacherRecord)) As DataTable
        Dim table As DataTable = CreateEmptyTeachersTable()

        If records Is Nothing Then
            Return table
        End If

        For Each record As TeacherRecord In records
            Dim row As DataRow = table.NewRow()
            row("Teacher ID") = record.EmployeeNumber
            row("Full Name") = record.FullName
            row("First Name") = record.FirstName
            row("Middle Name") = record.MiddleName
            row("Last Name") = record.LastName
            row("Department") = record.DepartmentDisplayName
            row("Advisory") = record.AdvisorySection
            row("Photo Path") = record.PhotoPath
            row("Email") = record.Email
            table.Rows.Add(row)
        Next

        Return table
    End Function

    Private Function CreateEmptyTeachersTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Teacher ID", GetType(String))
        table.Columns.Add("Full Name", GetType(String))
        table.Columns.Add("First Name", GetType(String))
        table.Columns.Add("Middle Name", GetType(String))
        table.Columns.Add("Last Name", GetType(String))
        table.Columns.Add("Department", GetType(String))
        table.Columns.Add("Advisory", GetType(String))
        table.Columns.Add("Photo Path", GetType(String))
        table.Columns.Add("Email", GetType(String))
        Return table
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
        OpenTeacherForm(TeacherFormMode.Add,
                        "Add Teacher",
                        "Create a new teacher record.",
                        "Add Teacher",
                        Nothing)
        TeacherFormTeacherIdTextBox.Focus()
    End Sub

    Private Sub EditSelectedTeacherButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        OpenTeacherForm(TeacherFormMode.Edit,
                        "Edit Teacher",
                        "Update teacher details.",
                        "Save Changes",
                        selectedRow)
        TeacherFormFirstNameTextBox.Focus()
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

    Private Sub NestedPanelScrollViewer_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)
        Dim activeScrollViewer As ScrollViewer = TryCast(sender, ScrollViewer)
        If activeScrollViewer Is Nothing OrElse CanScrollViewerConsumeWheel(activeScrollViewer, e.Delta) Then
            Return
        End If

        ForwardWheelToRootScrollViewer(sender, e)
    End Sub

    Private Sub TeachersDataGrid_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)
        Dim activeScrollViewer As ScrollViewer =
            FindDescendantScrollViewer(TryCast(sender, DependencyObject))

        If activeScrollViewer IsNot Nothing AndAlso
           CanScrollViewerConsumeWheel(activeScrollViewer, e.Delta) Then
            Return
        End If

        ForwardWheelToRootScrollViewer(sender, e)
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
        UpdateImageControlSource(TeacherFormPhotoPreviewImage,
                                 If(TeacherFormPhotoPathTextBox.Text, String.Empty))
    End Sub

    Private Sub SaveTeacherFormButton_Click(sender As Object, e As RoutedEventArgs)
        Dim formValues As New TeacherFormValues()
        If Not TryReadTeacherForm(formValues) Then
            Return
        End If

        Dim request As New TeacherSaveRequest() With {
            .OriginalEmployeeNumber = _editingTeacherOriginalId,
            .EmployeeNumber = formValues.TeacherId,
            .FirstName = formValues.FirstName,
            .MiddleName = formValues.MiddleName,
            .LastName = formValues.LastName,
            .DepartmentText = formValues.Department,
            .AdvisorySection = formValues.Advisory,
            .PhotoPath = formValues.PhotoPath
        }

        Dim isAddMode As Boolean = _activeFormMode = TeacherFormMode.Add
        Dim result =
            If(isAddMode,
               _teacherManagementService.CreateTeacher(request),
               _teacherManagementService.UpdateTeacher(request))

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Teachers",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        LoadTeachersTable(result.Data.EmployeeNumber)
        HideTeacherForm()

        If isAddMode AndAlso Not String.IsNullOrWhiteSpace(result.Message) Then
            MessageBox.Show(result.Message,
                            "Teachers",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
        End If
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
        Dim confirmation As MessageBoxResult =
            MessageBox.Show("Delete " & recordLabel & "?",
                            "Delete Teacher",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        Dim result = _teacherManagementService.DeleteTeacher(teacherId)
        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Teachers",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        If _activeFormMode = TeacherFormMode.Edit AndAlso
           String.Equals(_editingTeacherOriginalId,
                         teacherId,
                         StringComparison.OrdinalIgnoreCase) Then
            HideTeacherForm()
        End If

        LoadTeachersTable()
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

    Private Function GetSelectedTeacherId() As String
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return String.Empty
        End If

        Return ReadRowValue(selectedRow, "Teacher ID")
    End Function

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
        If e Is Nothing OrElse TeachersRootScrollViewer Is Nothing Then
            Return
        End If

        e.Handled = True

        Dim forwardedEventArgs As New MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
        forwardedEventArgs.RoutedEvent = UIElement.MouseWheelEvent
        forwardedEventArgs.Source = source
        TeachersRootScrollViewer.RaiseEvent(forwardedEventArgs)
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

    Private Sub SetDetailsValue(target As TextBlock,
                                value As String,
                                Optional placeholder As String = "--")
        If target Is Nothing Then
            Return
        End If

        Dim normalizedValue As String = If(value, String.Empty).Trim()
        target.Text = If(String.IsNullOrWhiteSpace(normalizedValue), placeholder, normalizedValue)
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
        TeacherFormFirstNameTextBox.Text = ReadRowValue(row, "First Name")
        TeacherFormMiddleNameTextBox.Text = ReadRowValue(row, "Middle Name")
        TeacherFormLastNameTextBox.Text = ReadRowValue(row, "Last Name")
        TeacherFormDepartmentTextBox.Text = ReadRowValue(row, "Department")
        TeacherFormAdvisoryTextBox.Text = ReadRowValue(row, "Advisory")
        TeacherFormPhotoPathTextBox.Text = ReadRowValue(row, "Photo Path")
        UpdateImageControlSource(TeacherFormPhotoPreviewImage, TeacherFormPhotoPathTextBox.Text)
    End Sub

    Private Sub ClearTeacherFormInputs()
        TeacherFormTeacherIdTextBox.Text = String.Empty
        TeacherFormFirstNameTextBox.Text = String.Empty
        TeacherFormMiddleNameTextBox.Text = String.Empty
        TeacherFormLastNameTextBox.Text = String.Empty
        TeacherFormDepartmentTextBox.Text = String.Empty
        TeacherFormAdvisoryTextBox.Text = String.Empty
        TeacherFormPhotoPathTextBox.Text = String.Empty
        UpdateImageControlSource(TeacherFormPhotoPreviewImage, String.Empty)
    End Sub

    Private Function TryReadTeacherForm(ByRef values As TeacherFormValues) As Boolean
        values.TeacherId = If(TeacherFormTeacherIdTextBox.Text, String.Empty).Trim()
        values.FirstName = If(TeacherFormFirstNameTextBox.Text, String.Empty).Trim()
        values.MiddleName = If(TeacherFormMiddleNameTextBox.Text, String.Empty).Trim()
        values.LastName = If(TeacherFormLastNameTextBox.Text, String.Empty).Trim()
        values.Department = If(TeacherFormDepartmentTextBox.Text, String.Empty).Trim()
        values.Advisory = If(TeacherFormAdvisoryTextBox.Text, String.Empty).Trim()
        values.PhotoPath = If(TeacherFormPhotoPathTextBox.Text, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(values.TeacherId) Then
            MessageBox.Show("Teacher ID is required.",
                            "Teacher Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.FirstName) Then
            MessageBox.Show("First Name is required.",
                            "Teacher Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.LastName) Then
            MessageBox.Show("Last Name is required.",
                            "Teacher Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        Return True
    End Function

    Private Sub SelectTeacherById(teacherId As String)
        If _teachersTable Is Nothing Then
            Return
        End If

        Dim normalizedTeacherId As String = If(teacherId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedTeacherId) Then
            RefreshTeacherDetailsPanel()
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
