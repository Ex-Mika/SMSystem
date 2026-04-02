Imports System.Data
Imports System.IO
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Microsoft.Win32
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class AdminStudentsView
    Private Enum StudentFormMode
        Add
        Edit
    End Enum

    Private Structure StudentFormValues
        Public StudentId As String
        Public FirstName As String
        Public MiddleName As String
        Public LastName As String
        Public YearLevel As Integer?
        Public Course As String
        Public Section As String
        Public Password As String
        Public PhotoPath As String
    End Structure

    Private _studentsTable As DataTable
    Private _searchTerm As String = String.Empty
    Private _activeFormMode As StudentFormMode = StudentFormMode.Add
    Private _editingStudentOriginalId As String = String.Empty
    Private _studentFormPhotoPath As String = String.Empty
    Private ReadOnly _studentManagementService As New StudentManagementService()

    Public Sub New()
        InitializeComponent()
        LoadStudentsTable()
        HideStudentForm()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyStudentsFilter()
    End Sub

    Public Sub RefreshData()
        LoadStudentsTable(GetSelectedStudentId())
    End Sub

    Public Sub OpenAddStudentFormFromDashboard()
        OpenStudentForm(StudentFormMode.Add,
                        "Add Student",
                        "Create a new student record.",
                        "Add Student",
                        Nothing)
        StudentFormStudentIdTextBox.Focus()
    End Sub

    Private Sub LoadStudentsTable(Optional studentIdToSelect As String = "")
        Dim result = _studentManagementService.GetStudents()

        If Not result.IsSuccess Then
            SetStudentsTable(CreateEmptyStudentsTable())
            MessageBox.Show(result.Message,
                            "Students",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return
        End If

        SetStudentsTable(BuildStudentsTable(result.Data))

        If Not String.IsNullOrWhiteSpace(studentIdToSelect) Then
            SelectStudentById(studentIdToSelect)
        End If
    End Sub

    Public Sub SetStudentsTable(table As DataTable)
        _studentsTable = If(table, CreateEmptyStudentsTable())
        StudentsDataGrid.ItemsSource = _studentsTable.DefaultView
        ApplyStudentsFilter()
        UpdateStudentsCount()
        EnsureSelectedRowForDetails()
        RefreshStudentDetailsPanel()
    End Sub

    Private Function BuildStudentsTable(records As IEnumerable(Of StudentRecord)) As DataTable
        Dim table As DataTable = CreateEmptyStudentsTable()

        If records Is Nothing Then
            Return table
        End If

        For Each record As StudentRecord In records
            Dim row As DataRow = table.NewRow()
            row("Student ID") = record.StudentNumber
            row("Full Name") = record.FullName
            row("First Name") = record.FirstName
            row("Middle Name") = record.MiddleName
            row("Last Name") = record.LastName
            row("Year Level") = record.YearLevelLabel
            row("Course") = record.CourseDisplayName
            row("Section") = record.SectionName
            row("Photo Path") = record.PhotoPath
            row("Email") = record.Email
            table.Rows.Add(row)
        Next

        Return table
    End Function

    Private Function CreateEmptyStudentsTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Student ID", GetType(String))
        table.Columns.Add("Full Name", GetType(String))
        table.Columns.Add("First Name", GetType(String))
        table.Columns.Add("Middle Name", GetType(String))
        table.Columns.Add("Last Name", GetType(String))
        table.Columns.Add("Year Level", GetType(String))
        table.Columns.Add("Course", GetType(String))
        table.Columns.Add("Section", GetType(String))
        table.Columns.Add("Photo Path", GetType(String))
        table.Columns.Add("Email", GetType(String))
        Return table
    End Function

    Private Sub ApplyStudentsFilter()
        If _studentsTable Is Nothing Then
            UpdateStudentsCount()
            Return
        End If

        If String.IsNullOrWhiteSpace(_searchTerm) Then
            _studentsTable.DefaultView.RowFilter = String.Empty
        Else
            Dim escapedTerm As String = EscapeLikeValue(_searchTerm)
            _studentsTable.DefaultView.RowFilter =
                "[Student ID] LIKE '*" & escapedTerm & "*' OR " &
                "[Full Name] LIKE '*" & escapedTerm & "*' OR " &
                "[Year Level] LIKE '*" & escapedTerm & "*' OR " &
                "[Course] LIKE '*" & escapedTerm & "*' OR " &
                "[Section] LIKE '*" & escapedTerm & "*'"
        End If

        UpdateStudentsCount()
        EnsureSelectedRowForDetails()
        RefreshStudentDetailsPanel()
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

    Private Sub UpdateStudentsCount()
        If StudentsCountTextBlock Is Nothing Then
            Return
        End If

        Dim visibleCount As Integer = 0
        Dim totalCount As Integer = 0
        Dim hasSearch As Boolean = Not String.IsNullOrWhiteSpace(_searchTerm)

        If _studentsTable IsNot Nothing Then
            visibleCount = _studentsTable.DefaultView.Count
            totalCount = _studentsTable.Rows.Count
        End If

        If hasSearch AndAlso visibleCount <> totalCount Then
            StudentsCountTextBlock.Text =
                visibleCount.ToString() & " of " & totalCount.ToString()
        Else
            StudentsCountTextBlock.Text = totalCount.ToString()
        End If

    End Sub

    Private Sub OpenAddStudentButton_Click(sender As Object, e As RoutedEventArgs)
        OpenAddStudentFormFromDashboard()
    End Sub

    Private Sub EditSelectedStudentButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        OpenStudentForm(StudentFormMode.Edit,
                        "Edit Student",
                        "Update student details.",
                        "Save Changes",
                        selectedRow)
        StudentFormFirstNameTextBox.Focus()
    End Sub

    Private Sub DeleteSelectedStudentButton_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return
        End If

        DeleteStudentRow(selectedRow)
    End Sub

    Private Sub StudentsDataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        RefreshStudentDetailsPanel()
    End Sub

    Private Sub NestedPanelScrollViewer_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)
        Dim activeScrollViewer As ScrollViewer = TryCast(sender, ScrollViewer)
        If activeScrollViewer Is Nothing OrElse CanScrollViewerConsumeWheel(activeScrollViewer, e.Delta) Then
            Return
        End If

        ForwardWheelToRootScrollViewer(sender, e)
    End Sub

    Private Sub StudentsDataGrid_PreviewMouseWheel(sender As Object, e As MouseWheelEventArgs)
        Dim activeScrollViewer As ScrollViewer =
            FindDescendantScrollViewer(TryCast(sender, DependencyObject))

        If activeScrollViewer IsNot Nothing AndAlso
           CanScrollViewerConsumeWheel(activeScrollViewer, e.Delta) Then
            Return
        End If

        ForwardWheelToRootScrollViewer(sender, e)
    End Sub

    Private Sub BrowseStudentPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dialog As New OpenFileDialog() With {
            .Title = "Select Student Photo",
            .Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp|All Files|*.*",
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        If dialog.ShowDialog() = True Then
            SetStudentFormPhotoPath(dialog.FileName)
        End If
    End Sub

    Private Sub ClearStudentPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        SetStudentFormPhotoPath(String.Empty)
    End Sub

    Private Sub SaveStudentFormButton_Click(sender As Object, e As RoutedEventArgs)
        Dim formValues As New StudentFormValues()
        If Not TryReadStudentForm(formValues) Then
            Return
        End If

        Dim request As New StudentSaveRequest() With {
            .OriginalStudentNumber = _editingStudentOriginalId,
            .StudentNumber = formValues.StudentId,
            .FirstName = formValues.FirstName,
            .MiddleName = formValues.MiddleName,
            .LastName = formValues.LastName,
            .YearLevel = formValues.YearLevel,
            .CourseText = formValues.Course,
            .SectionName = formValues.Section,
            .Password = formValues.Password,
            .PhotoPath = formValues.PhotoPath
        }

        Dim isAddMode As Boolean = _activeFormMode = StudentFormMode.Add
        Dim result =
            If(isAddMode,
               _studentManagementService.CreateStudent(request),
               _studentManagementService.UpdateStudent(request))

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Students",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        LoadStudentsTable(result.Data.StudentNumber)
        HideStudentForm()

        If isAddMode AndAlso Not String.IsNullOrWhiteSpace(result.Message) Then
            MessageBox.Show(result.Message,
                            "Students",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
        End If
    End Sub

    Private Sub CancelStudentFormButton_Click(sender As Object, e As RoutedEventArgs)
        HideStudentForm()
    End Sub

    Private Sub DeleteStudentRow(row As DataRow)
        If row Is Nothing OrElse row.RowState = DataRowState.Deleted Then
            Return
        End If

        Dim fullName As String = ReadRowValue(row, "Full Name")
        Dim studentId As String = ReadRowValue(row, "Student ID")
        Dim recordLabel As String = If(String.IsNullOrWhiteSpace(fullName), studentId, fullName)

        Dim confirmation As MessageBoxResult =
            MessageBox.Show("Delete " & recordLabel & "?",
                            "Delete Student",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question)
        If confirmation <> MessageBoxResult.Yes Then
            Return
        End If

        Dim result = _studentManagementService.DeleteStudent(studentId)
        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Students",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        If _activeFormMode = StudentFormMode.Edit AndAlso
           String.Equals(_editingStudentOriginalId,
                         studentId,
                         StringComparison.OrdinalIgnoreCase) Then
            HideStudentForm()
        End If

        LoadStudentsTable()
    End Sub

    Private Sub OpenStudentForm(mode As StudentFormMode,
                                title As String,
                                subtitle As String,
                                actionText As String,
                                row As DataRow)
        _activeFormMode = mode
        StudentFormTitleTextBlock.Text = title
        StudentFormSubtitleTextBlock.Text = subtitle
        SaveStudentFormButton.Content = actionText

        If row Is Nothing Then
            _editingStudentOriginalId = String.Empty
            ClearStudentFormInputs()
        Else
            _editingStudentOriginalId = ReadRowValue(row, "Student ID")
            PopulateStudentForm(row)
        End If

        ShowStudentForm()
    End Sub

    Private Sub EnsureSelectedRowForDetails()
        If StudentsDataGrid Is Nothing OrElse _studentsTable Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow IsNot Nothing Then
            Return
        End If

        If _studentsTable.DefaultView.Count > 0 Then
            StudentsDataGrid.SelectedItem = _studentsTable.DefaultView(0)
        End If
    End Sub

    Private Function GetSelectedStudentId() As String
        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        If selectedRow Is Nothing Then
            Return String.Empty
        End If

        Return ReadRowValue(selectedRow, "Student ID")
    End Function

    Private Function TryGetSelectedGridRow() As DataRow
        Dim selectedRowView As DataRowView = TryCast(StudentsDataGrid.SelectedItem, DataRowView)
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
        If e Is Nothing OrElse StudentsRootScrollViewer Is Nothing Then
            Return
        End If

        e.Handled = True

        Dim forwardedEventArgs As New MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
        forwardedEventArgs.RoutedEvent = UIElement.MouseWheelEvent
        forwardedEventArgs.Source = source
        StudentsRootScrollViewer.RaiseEvent(forwardedEventArgs)
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

    Private Sub RefreshStudentDetailsPanel()
        If StudentDetailsStudentIdTextBlock Is Nothing Then
            Return
        End If

        Dim selectedRow As DataRow = TryGetSelectedGridRow()
        Dim hasSelection As Boolean = selectedRow IsNot Nothing

        If EditSelectedStudentButton IsNot Nothing Then
            EditSelectedStudentButton.IsEnabled = hasSelection
        End If

        If DeleteSelectedStudentButton IsNot Nothing Then
            DeleteSelectedStudentButton.IsEnabled = hasSelection
        End If

        If Not hasSelection Then
            StudentDetailsSubtitleTextBlock.Text =
                "Choose a student from the roster to preview details."
            SetDetailsValue(StudentDetailsHeroNameTextBlock,
                            String.Empty,
                            "No student selected")
            SetDetailsValue(StudentDetailsStudentIdTextBlock, String.Empty)
            SetDetailsValue(StudentDetailsEmailTextBlock, String.Empty)
            SetDetailsValue(StudentDetailsFullNameTextBlock, String.Empty)
            SetDetailsValue(StudentDetailsYearLevelTextBlock, String.Empty)
            SetDetailsValue(StudentDetailsCourseTextBlock, String.Empty)
            SetDetailsValue(StudentDetailsSectionTextBlock, String.Empty)
            UpdateImageControlSource(StudentDetailsPhotoImage, String.Empty)
            Return
        End If

        Dim fullName As String = ReadRowValue(selectedRow, "Full Name")

        StudentDetailsSubtitleTextBlock.Text =
            "Selected student record. Use Edit to update details."
        SetDetailsValue(StudentDetailsHeroNameTextBlock, fullName, "Unnamed student")
        SetDetailsValue(StudentDetailsStudentIdTextBlock, ReadRowValue(selectedRow, "Student ID"))
        SetDetailsValue(StudentDetailsEmailTextBlock, ReadRowValue(selectedRow, "Email"))
        SetDetailsValue(StudentDetailsFullNameTextBlock, fullName)
        SetDetailsValue(StudentDetailsYearLevelTextBlock, ReadRowValue(selectedRow, "Year Level"))
        SetDetailsValue(StudentDetailsCourseTextBlock, ReadRowValue(selectedRow, "Course"))
        SetDetailsValue(StudentDetailsSectionTextBlock, ReadRowValue(selectedRow, "Section"))
        UpdateImageControlSource(StudentDetailsPhotoImage, ReadRowValue(selectedRow, "Photo Path"))
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

    Private Sub ShowStudentForm()
        StudentFormPanel.Visibility = Visibility.Visible
        StudentDetailsPanel.Visibility = Visibility.Collapsed
    End Sub

    Private Sub HideStudentForm()
        _activeFormMode = StudentFormMode.Add
        _editingStudentOriginalId = String.Empty
        StudentFormTitleTextBlock.Text = "Add Student"
        StudentFormSubtitleTextBlock.Text = "Create a new student record."
        SaveStudentFormButton.Content = "Add Student"
        ClearStudentFormInputs()
        StudentFormPanel.Visibility = Visibility.Collapsed
        StudentDetailsPanel.Visibility = Visibility.Visible
        RefreshStudentDetailsPanel()
    End Sub

    Private Sub PopulateStudentForm(row As DataRow)
        StudentFormStudentIdTextBox.Text = ReadRowValue(row, "Student ID")
        StudentFormFirstNameTextBox.Text = ReadRowValue(row, "First Name")
        StudentFormMiddleNameTextBox.Text = ReadRowValue(row, "Middle Name")
        StudentFormLastNameTextBox.Text = ReadRowValue(row, "Last Name")
        SetComboBoxValue(StudentFormYearLevelComboBox, ReadRowValue(row, "Year Level"))
        StudentFormCourseTextBox.Text = ReadRowValue(row, "Course")
        StudentFormSectionTextBox.Text = ReadRowValue(row, "Section")
        StudentFormPasswordTextBox.Text = String.Empty
        SetStudentFormPhotoPath(ReadRowValue(row, "Photo Path"))
    End Sub

    Private Sub ClearStudentFormInputs()
        StudentFormStudentIdTextBox.Text = String.Empty
        StudentFormFirstNameTextBox.Text = String.Empty
        StudentFormMiddleNameTextBox.Text = String.Empty
        StudentFormLastNameTextBox.Text = String.Empty
        StudentFormYearLevelComboBox.SelectedIndex = -1
        StudentFormCourseTextBox.Text = String.Empty
        StudentFormSectionTextBox.Text = String.Empty
        StudentFormPasswordTextBox.Text = String.Empty
        SetStudentFormPhotoPath(String.Empty)
    End Sub

    Private Function TryReadStudentForm(ByRef values As StudentFormValues) As Boolean
        values.StudentId = If(StudentFormStudentIdTextBox.Text, String.Empty).Trim()
        values.FirstName = If(StudentFormFirstNameTextBox.Text, String.Empty).Trim()
        values.MiddleName = If(StudentFormMiddleNameTextBox.Text, String.Empty).Trim()
        values.LastName = If(StudentFormLastNameTextBox.Text, String.Empty).Trim()
        values.YearLevel = ParseYearLevelValue(ReadComboBoxValue(StudentFormYearLevelComboBox))
        values.Course = If(StudentFormCourseTextBox.Text, String.Empty).Trim()
        values.Section = If(StudentFormSectionTextBox.Text, String.Empty).Trim()
        values.Password = If(StudentFormPasswordTextBox.Text, String.Empty).Trim()
        values.PhotoPath = _studentFormPhotoPath

        If String.IsNullOrWhiteSpace(values.StudentId) Then
            MessageBox.Show("Student ID is required.",
                            "Student Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.FirstName) Then
            MessageBox.Show("First Name is required.",
                            "Student Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        If String.IsNullOrWhiteSpace(values.LastName) Then
            MessageBox.Show("Last Name is required.",
                            "Student Form",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return False
        End If

        Return True
    End Function

    Private Sub SetStudentFormPhotoPath(photoPath As String)
        _studentFormPhotoPath = If(photoPath, String.Empty).Trim()
        UpdateImageControlSource(StudentFormPhotoPreviewImage, _studentFormPhotoPath)
    End Sub

    Private Function ParseYearLevelValue(displayValue As String) As Integer?
        Dim normalizedValue As String = If(displayValue, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedValue) Then
            Return Nothing
        End If

        Select Case normalizedValue.ToLowerInvariant()
            Case "1", "1st year"
                Return 1
            Case "2", "2nd year"
                Return 2
            Case "3", "3rd year"
                Return 3
            Case "4", "4th year"
                Return 4
        End Select

        Dim numericValue As Integer
        If Integer.TryParse(normalizedValue, numericValue) Then
            Return numericValue
        End If

        Return Nothing
    End Function

    Private Function ReadComboBoxValue(comboBox As ComboBox) As String
        If comboBox Is Nothing Then
            Return String.Empty
        End If

        Dim selectedItem As ComboBoxItem = TryCast(comboBox.SelectedItem, ComboBoxItem)
        If selectedItem IsNot Nothing Then
            Return If(selectedItem.Content, String.Empty).ToString().Trim()
        End If

        Return If(comboBox.Text, String.Empty).Trim()
    End Function

    Private Sub SetComboBoxValue(comboBox As ComboBox, value As String)
        If comboBox Is Nothing Then
            Return
        End If

        Dim requestedValue As String = If(value, String.Empty).Trim()
        comboBox.SelectedIndex = -1

        For Each rawItem As Object In comboBox.Items
            Dim comboItem As ComboBoxItem = TryCast(rawItem, ComboBoxItem)
            If comboItem Is Nothing Then
                Continue For
            End If

            Dim itemValue As String = If(comboItem.Content, String.Empty).ToString().Trim()
            If String.Equals(itemValue, requestedValue, StringComparison.OrdinalIgnoreCase) Then
                comboBox.SelectedItem = comboItem
                Return
            End If
        Next
    End Sub

    Private Sub SelectStudentById(studentId As String)
        If _studentsTable Is Nothing Then
            Return
        End If

        Dim normalizedStudentId As String = If(studentId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedStudentId) Then
            RefreshStudentDetailsPanel()
            Return
        End If

        For Each rowView As DataRowView In _studentsTable.DefaultView
            Dim candidateId As String = If(rowView("Student ID"), String.Empty).ToString().Trim()
            If String.Equals(candidateId, normalizedStudentId, StringComparison.OrdinalIgnoreCase) Then
                StudentsDataGrid.SelectedItem = rowView
                StudentsDataGrid.ScrollIntoView(rowView)
                RefreshStudentDetailsPanel()
                Return
            End If
        Next

        RefreshStudentDetailsPanel()
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
