Imports System.Collections.Generic
Imports System.Data
Imports System.Windows.Data

Class StudentDashboardWindow
    Private Enum StudentDashboardSection
        Dashboard
        MySubjects
        ClassSchedule
        Grades
        Profile
    End Enum

    Public Property LoggedInStudentId As String = String.Empty
    Public Property LoggedInStudentName As String = String.Empty

    Public Sub New()
        InitializeComponent()
        UpdateMaximizeRestoreIcon()
        ApplyLoggedInStudentProfile()
        LoadDashboardTimetable()
        SetActiveSection(StudentDashboardSection.Dashboard)
    End Sub

    Public Sub SetLoggedInStudent(studentId As String, Optional studentName As String = "")
        LoggedInStudentId = If(studentId, String.Empty).Trim()
        LoggedInStudentName = If(studentName, String.Empty).Trim()
        ApplyLoggedInStudentProfile()
        LoadDashboardTimetable()
    End Sub

    Private Sub TitleBar_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
        If e.ChangedButton <> MouseButton.Left Then
            Return
        End If

        Dim sourceElement As DependencyObject = TryCast(e.OriginalSource, DependencyObject)
        While sourceElement IsNot Nothing
            If TypeOf sourceElement Is Button Then
                Return
            End If

            sourceElement = VisualTreeHelper.GetParent(sourceElement)
        End While

        If e.ClickCount = 2 Then
            ToggleWindowState()
            Return
        End If

        Try
            DragMove()
        Catch
            ' Ignore drag exceptions from rapid state changes.
        End Try
    End Sub

    Private Sub MinimizeButton_Click(sender As Object, e As RoutedEventArgs)
        WindowState = WindowState.Minimized
    End Sub

    Private Sub MaximizeRestoreButton_Click(sender As Object, e As RoutedEventArgs)
        ToggleWindowState()
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As RoutedEventArgs)
        Close()
    End Sub

    Private Sub StudentDashboardWindow_StateChanged(sender As Object, e As EventArgs)
        UpdateMaximizeRestoreIcon()
    End Sub

    Private Sub DashboardNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentDashboardSection.Dashboard)
    End Sub

    Private Sub ProfileNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentDashboardSection.Profile)
    End Sub

    Private Sub MySubjectsNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentDashboardSection.MySubjects)
    End Sub

    Private Sub ClassScheduleNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentDashboardSection.ClassSchedule)
    End Sub

    Private Sub GradesNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentDashboardSection.Grades)
    End Sub

    Private Sub SignOutButton_Click(sender As Object, e As RoutedEventArgs)
        Dim loginWindow As New LoginWindow()
        loginWindow.Show()
        Close()
    End Sub

    Private Sub ToggleWindowState()
        If ResizeMode = ResizeMode.NoResize Then
            Return
        End If

        WindowState = If(WindowState = WindowState.Maximized, WindowState.Normal, WindowState.Maximized)
        UpdateMaximizeRestoreIcon()
    End Sub

    Private Sub SetActiveSection(section As StudentDashboardSection)
        Dim isDashboardSelected As Boolean = section = StudentDashboardSection.Dashboard
        Dim isMySubjectsSelected As Boolean = section = StudentDashboardSection.MySubjects
        Dim isClassScheduleSelected As Boolean = section = StudentDashboardSection.ClassSchedule
        Dim isGradesSelected As Boolean = section = StudentDashboardSection.Grades
        Dim isProfileSelected As Boolean = section = StudentDashboardSection.Profile

        DashboardContentScrollViewer.Visibility = If(isDashboardSelected, Visibility.Visible, Visibility.Collapsed)
        MySubjectsContentView.Visibility = If(isMySubjectsSelected, Visibility.Visible, Visibility.Collapsed)
        ClassScheduleContentView.Visibility = If(isClassScheduleSelected, Visibility.Visible, Visibility.Collapsed)
        GradesContentView.Visibility = If(isGradesSelected, Visibility.Visible, Visibility.Collapsed)
        ProfileContentView.Visibility = If(isProfileSelected, Visibility.Visible, Visibility.Collapsed)

        If isDashboardSelected Then
            ContentTitleTextBlock.Text = "Overview"
        ElseIf isMySubjectsSelected Then
            ContentTitleTextBlock.Text = "My Subjects"
        ElseIf isClassScheduleSelected Then
            ContentTitleTextBlock.Text = "Class Schedule"
        ElseIf isGradesSelected Then
            ContentTitleTextBlock.Text = "Grades"
        Else
            ContentTitleTextBlock.Text = "My Profile"
        End If

        SearchDashboardTextBox.Visibility = If(isProfileSelected, Visibility.Collapsed, Visibility.Visible)
        ApplySidebarSelectionStyles(section)
    End Sub

    Private Sub ApplySidebarSelectionStyles(activeSection As StudentDashboardSection)
        ApplySidebarButtonState(DashboardNavButton, DashboardNavIconBorder, DashboardNavIconText, DashboardNavText, activeSection = StudentDashboardSection.Dashboard)
        ApplySidebarButtonState(MySubjectsNavButton, MySubjectsNavIconBorder, MySubjectsNavIconText, MySubjectsNavText, activeSection = StudentDashboardSection.MySubjects)
        ApplySidebarButtonState(ClassScheduleNavButton, ClassScheduleNavIconBorder, ClassScheduleNavIconText, ClassScheduleNavText, activeSection = StudentDashboardSection.ClassSchedule)
        ApplySidebarButtonState(GradesNavButton, GradesNavIconBorder, GradesNavIconText, GradesNavText, activeSection = StudentDashboardSection.Grades)
        ApplySidebarButtonState(ProfileNavButton, ProfileNavIconBorder, ProfileNavIconText, ProfileNavText, activeSection = StudentDashboardSection.Profile)
    End Sub

    Private Sub ApplySidebarButtonState(navButton As Button, navIconBorder As Border, navIconText As TextBlock, navText As TextBlock, isSelected As Boolean)
        navButton.Style = CType(FindResource(If(isSelected, "DashboardSidebarNavSelectedButtonStyle", "DashboardSidebarNavButtonStyle")), Style)
        navIconBorder.Style = CType(FindResource(If(isSelected, "DashboardSidebarIconSelectedBadgeStyle", "DashboardSidebarIconBadgeStyle")), Style)
        navIconText.Style = CType(FindResource(If(isSelected, "DashboardSidebarIconGlyphSelectedStyle", "DashboardSidebarIconGlyphStyle")), Style)
        navText.Style = CType(FindResource(If(isSelected, "DashboardSidebarNavTextSelectedStyle", "DashboardSidebarNavTextStyle")), Style)
    End Sub

    Private Sub ApplyLoggedInStudentProfile()
        Dim studentId As String = If(LoggedInStudentId, String.Empty).Trim()
        Dim studentName As String = If(LoggedInStudentName, String.Empty).Trim()

        HeaderStudentNameTextBlock.Text = If(String.IsNullOrWhiteSpace(studentName), "Student", studentName)
        HeaderStudentIdTextBlock.Text = If(String.IsNullOrWhiteSpace(studentId), " ", studentId)
        SidebarStudentSubtitleTextBlock.Text = If(String.IsNullOrWhiteSpace(studentId), "Student Dashboard", studentId)
        MySubjectsContentView.SetStudentContext(studentId, studentName)
        ClassScheduleContentView.SetStudentContext(studentId, studentName)
        GradesContentView.SetStudentContext(studentId, studentName)
        ProfileContentView.SetStudentContext(studentId, studentName)
    End Sub

    Public Sub SetDashboardTimetable(table As DataTable)
        Dim resolvedTable As DataTable = table
        If resolvedTable Is Nothing OrElse resolvedTable.Columns.Count = 0 Then
            resolvedTable = CreateDefaultDashboardTimetable()
        End If

        resolvedTable = EnsureSessionColumn(resolvedTable)
        BuildDashboardTimetableColumns(resolvedTable)
        DashboardTimetableDataGrid.ItemsSource = resolvedTable.DefaultView
    End Sub

    Private Sub LoadDashboardTimetable()
        Dim timetableTable As DataTable = FetchDashboardTimetableForStudent(LoggedInStudentId)
        If timetableTable Is Nothing Then
            timetableTable = CreateDefaultDashboardTimetable()
        End If

        SetDashboardTimetable(timetableTable)
    End Sub

    Private Function FetchDashboardTimetableForStudent(studentId As String) As DataTable
        ' Hook your database query here and return the result as a DataTable.
        Return Nothing
    End Function

    Private Function CreateDefaultDashboardTimetable() As DataTable
        Return CreateEmptyDashboardTimetable(New String() {"Mon", "Tue", "Wed", "Thu", "Fri"}, 2)
    End Function

    Private Function EnsureSessionColumn(source As DataTable) As DataTable
        If source Is Nothing Then
            Return Nothing
        End If

        Dim table As DataTable = source.Copy()
        Dim sessionColumn As DataColumn = FindTableColumn(table, "Session")

        If sessionColumn Is Nothing Then
            Dim aliasColumn As DataColumn = FindTableColumn(table, "Time", "Time Slot", "Timeslot", "Period", "Slot", "Schedule")
            If aliasColumn IsNot Nothing Then
                aliasColumn.ColumnName = "Session"
                sessionColumn = aliasColumn
            End If
        End If

        If sessionColumn Is Nothing Then
            sessionColumn = New DataColumn("Session", GetType(String))
            table.Columns.Add(sessionColumn)
            For Each row As DataRow In table.Rows
                row("Session") = "--"
            Next
        End If

        For Each row As DataRow In table.Rows
            Dim sessionValue As String = If(row("Session"), String.Empty).ToString().Trim()
            If String.IsNullOrWhiteSpace(sessionValue) Then
                row("Session") = "--"
            End If
        Next

        sessionColumn.SetOrdinal(0)

        Return table
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

    Private Function CreateEmptyDashboardTimetable(dayHeaders As IEnumerable(Of String), rowCount As Integer) As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Session", GetType(String))

        Dim normalizedDays As New List(Of String)()
        If dayHeaders IsNot Nothing Then
            For Each dayHeader As String In dayHeaders
                Dim normalizedDay As String = If(dayHeader, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalizedDay) Then
                    Continue For
                End If

                Dim exists As Boolean = False
                For Each existingDay As String In normalizedDays
                    If String.Equals(existingDay, normalizedDay, StringComparison.OrdinalIgnoreCase) Then
                        exists = True
                        Exit For
                    End If
                Next

                If Not exists Then
                    normalizedDays.Add(normalizedDay)
                    table.Columns.Add(normalizedDay, GetType(String))
                End If
            Next
        End If

        If normalizedDays.Count = 0 Then
            For Each fallbackDay As String In New String() {"Mon", "Tue", "Wed", "Thu", "Fri"}
                normalizedDays.Add(fallbackDay)
                table.Columns.Add(fallbackDay, GetType(String))
            Next
        End If

        Dim safeRowCount As Integer = Math.Max(1, rowCount)
        For rowIndex As Integer = 0 To safeRowCount - 1
            Dim row As DataRow = table.NewRow()
            row("Session") = "--"

            For Each day As String In normalizedDays
                row(day) = "--"
            Next

            table.Rows.Add(row)
        Next

        Return table
    End Function

    Private Sub BuildDashboardTimetableColumns(sourceTable As DataTable)
        If DashboardTimetableDataGrid Is Nothing OrElse sourceTable Is Nothing Then
            Return
        End If

        DashboardTimetableDataGrid.Columns.Clear()

        Dim sessionColumn As New DataGridTextColumn()
        sessionColumn.Header = "Session"
        sessionColumn.Binding = New Binding(BuildDataTableBindingPath("Session"))
        sessionColumn.IsReadOnly = True
        sessionColumn.CanUserSort = False
        sessionColumn.MinWidth = 130
        sessionColumn.Width = New DataGridLength(1, DataGridLengthUnitType.SizeToCells)
        DashboardTimetableDataGrid.Columns.Add(sessionColumn)

        For Each tableColumn As DataColumn In sourceTable.Columns
            If String.Equals(tableColumn.ColumnName, "Session", StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim dynamicColumn As New DataGridTextColumn()
            dynamicColumn.Header = tableColumn.ColumnName
            dynamicColumn.Binding = New Binding(BuildDataTableBindingPath(tableColumn.ColumnName))
            dynamicColumn.IsReadOnly = True
            dynamicColumn.CanUserSort = False
            dynamicColumn.Width = New DataGridLength(1, DataGridLengthUnitType.Star)
            DashboardTimetableDataGrid.Columns.Add(dynamicColumn)
        Next
    End Sub

    Private Function BuildDataTableBindingPath(columnName As String) As String
        Dim safeColumnName As String = If(columnName, String.Empty)
        Return "[" & safeColumnName.Replace("]", "]]") & "]"
    End Function

    Private Sub UpdateMaximizeRestoreIcon()
        If MaximizeRestoreIcon Is Nothing Then
            Return
        End If

        MaximizeRestoreIcon.Text = If(WindowState = WindowState.Maximized, ChrW(&HE923), ChrW(&HE922))
    End Sub
End Class
