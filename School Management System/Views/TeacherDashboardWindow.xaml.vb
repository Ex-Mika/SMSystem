Imports System.Collections.Generic
Imports System.Data
Imports System.Windows.Data

Class TeacherDashboardWindow
    Private Enum TeacherDashboardSection
        Dashboard
        Classes
        Schedule
        Grading
        Profile
    End Enum

    Public Property LoggedInTeacherId As String = String.Empty
    Public Property LoggedInTeacherName As String = String.Empty

    Public Sub New()
        InitializeComponent()
        UpdateMaximizeRestoreIcon()
        ApplyLoggedInTeacherProfile()
        LoadDashboardTimetable()
        SetActiveSection(TeacherDashboardSection.Dashboard)
    End Sub

    Public Sub SetLoggedInTeacher(teacherId As String, Optional teacherName As String = "")
        LoggedInTeacherId = If(teacherId, String.Empty).Trim()
        LoggedInTeacherName = If(teacherName, String.Empty).Trim()
        ApplyLoggedInTeacherProfile()
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

    Private Sub TeacherDashboardWindow_StateChanged(sender As Object, e As EventArgs)
        UpdateMaximizeRestoreIcon()
    End Sub

    Private Sub DashboardNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(TeacherDashboardSection.Dashboard)
    End Sub

    Private Sub ProfileNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(TeacherDashboardSection.Profile)
    End Sub

    Private Sub ClassesNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(TeacherDashboardSection.Classes)
    End Sub

    Private Sub ScheduleNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(TeacherDashboardSection.Schedule)
    End Sub

    Private Sub GradingNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(TeacherDashboardSection.Grading)
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

    Private Sub SetActiveSection(section As TeacherDashboardSection)
        Dim isDashboardSelected As Boolean = section = TeacherDashboardSection.Dashboard
        Dim isClassesSelected As Boolean = section = TeacherDashboardSection.Classes
        Dim isScheduleSelected As Boolean = section = TeacherDashboardSection.Schedule
        Dim isGradingSelected As Boolean = section = TeacherDashboardSection.Grading
        Dim isProfileSelected As Boolean = section = TeacherDashboardSection.Profile

        DashboardContentScrollViewer.Visibility = If(isDashboardSelected, Visibility.Visible, Visibility.Collapsed)
        ClassesContentView.Visibility = If(isClassesSelected, Visibility.Visible, Visibility.Collapsed)
        ScheduleContentView.Visibility = If(isScheduleSelected, Visibility.Visible, Visibility.Collapsed)
        GradingContentView.Visibility = If(isGradingSelected, Visibility.Visible, Visibility.Collapsed)
        ProfileContentView.Visibility = If(isProfileSelected, Visibility.Visible, Visibility.Collapsed)

        If isDashboardSelected Then
            ContentTitleTextBlock.Text = "Overview"
        ElseIf isClassesSelected Then
            ContentTitleTextBlock.Text = "My Classes"
        ElseIf isScheduleSelected Then
            ContentTitleTextBlock.Text = "Schedule"
        ElseIf isGradingSelected Then
            ContentTitleTextBlock.Text = "Grading"
        Else
            ContentTitleTextBlock.Text = "Profile"
        End If

        SearchDashboardTextBox.Visibility = If(isProfileSelected, Visibility.Collapsed, Visibility.Visible)
        ApplySidebarSelectionStyles(section)
    End Sub

    Private Sub ApplySidebarSelectionStyles(activeSection As TeacherDashboardSection)
        ApplySidebarButtonState(DashboardNavButton, DashboardNavIconBorder, DashboardNavIconText, DashboardNavText, activeSection = TeacherDashboardSection.Dashboard)
        ApplySidebarButtonState(ClassesNavButton, ClassesNavIconBorder, ClassesNavIconText, ClassesNavText, activeSection = TeacherDashboardSection.Classes)
        ApplySidebarButtonState(ScheduleNavButton, ScheduleNavIconBorder, ScheduleNavIconText, ScheduleNavText, activeSection = TeacherDashboardSection.Schedule)
        ApplySidebarButtonState(GradingNavButton, GradingNavIconBorder, GradingNavIconText, GradingNavText, activeSection = TeacherDashboardSection.Grading)
        ApplySidebarButtonState(ProfileNavButton, ProfileNavIconBorder, ProfileNavIconText, ProfileNavText, activeSection = TeacherDashboardSection.Profile)
    End Sub

    Private Sub ApplySidebarButtonState(navButton As Button, navIconBorder As Border, navIconText As TextBlock, navText As TextBlock, isSelected As Boolean)
        navButton.Style = CType(FindResource(If(isSelected, "DashboardSidebarNavSelectedButtonStyle", "DashboardSidebarNavButtonStyle")), Style)
        navIconBorder.Style = CType(FindResource(If(isSelected, "DashboardSidebarIconSelectedBadgeStyle", "DashboardSidebarIconBadgeStyle")), Style)
        navIconText.Style = CType(FindResource(If(isSelected, "DashboardSidebarIconGlyphSelectedStyle", "DashboardSidebarIconGlyphStyle")), Style)
        navText.Style = CType(FindResource(If(isSelected, "DashboardSidebarNavTextSelectedStyle", "DashboardSidebarNavTextStyle")), Style)
    End Sub

    Private Sub ApplyLoggedInTeacherProfile()
        Dim teacherId As String = If(LoggedInTeacherId, String.Empty).Trim()
        Dim teacherName As String = If(LoggedInTeacherName, String.Empty).Trim()

        HeaderTeacherNameTextBlock.Text = If(String.IsNullOrWhiteSpace(teacherName), "Teacher", teacherName)
        HeaderTeacherIdTextBlock.Text = If(String.IsNullOrWhiteSpace(teacherId), " ", teacherId)
        SidebarTeacherSubtitleTextBlock.Text = If(String.IsNullOrWhiteSpace(teacherId), "Teacher Dashboard", teacherId)
        ClassesContentView.SetTeacherContext(teacherId, teacherName)
        ScheduleContentView.SetTeacherContext(teacherId, teacherName)
        GradingContentView.SetTeacherContext(teacherId, teacherName)
        ProfileContentView.SetTeacherContext(teacherId, teacherName)
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
        Dim timetableTable As DataTable = FetchDashboardTimetableForTeacher(LoggedInTeacherId)
        If timetableTable Is Nothing Then
            timetableTable = CreateDefaultDashboardTimetable()
        End If

        SetDashboardTimetable(timetableTable)
    End Sub

    Private Function FetchDashboardTimetableForTeacher(teacherId As String) As DataTable
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

