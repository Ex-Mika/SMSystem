Class AdminDashboardWindow
    Private Enum AdminDashboardSection
        Dashboard
        Students
        Teachers
        Administrators
        Courses
        Departments
        Subjects
        Reports
        Scheduling
    End Enum

    Private _activeSection As AdminDashboardSection = AdminDashboardSection.Dashboard

    Public Sub New()
        InitializeComponent()
        UpdateMaximizeRestoreIcon()
        SetActiveSection(AdminDashboardSection.Dashboard)
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

    Private Sub AdminDashboardWindow_StateChanged(sender As Object, e As EventArgs)
        UpdateMaximizeRestoreIcon()
    End Sub

    Private Sub DashboardNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(AdminDashboardSection.Dashboard)
    End Sub

    Private Sub StudentsNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(AdminDashboardSection.Students)
    End Sub

    Private Sub TeachersNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(AdminDashboardSection.Teachers)
    End Sub

    Private Sub AdministratorsNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(AdminDashboardSection.Administrators)
    End Sub

    Private Sub CoursesNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(AdminDashboardSection.Courses)
    End Sub

    Private Sub DepartmentsNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(AdminDashboardSection.Departments)
    End Sub

    Private Sub SubjectsNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(AdminDashboardSection.Subjects)
    End Sub

    Private Sub ReportsNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(AdminDashboardSection.Reports)
    End Sub

    Private Sub SchedulingNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(AdminDashboardSection.Scheduling)
    End Sub

    Private Sub SetActiveSection(section As AdminDashboardSection)
        _activeSection = section

        DashboardContentView.Visibility = If(section = AdminDashboardSection.Dashboard, Visibility.Visible, Visibility.Collapsed)
        StudentsContentView.Visibility = If(section = AdminDashboardSection.Students, Visibility.Visible, Visibility.Collapsed)
        TeachersContentView.Visibility = If(section = AdminDashboardSection.Teachers, Visibility.Visible, Visibility.Collapsed)
        AdministratorsContentView.Visibility = If(section = AdminDashboardSection.Administrators, Visibility.Visible, Visibility.Collapsed)
        CoursesContentView.Visibility = If(section = AdminDashboardSection.Courses, Visibility.Visible, Visibility.Collapsed)
        DepartmentsContentView.Visibility = If(section = AdminDashboardSection.Departments, Visibility.Visible, Visibility.Collapsed)
        SubjectsContentView.Visibility = If(section = AdminDashboardSection.Subjects, Visibility.Visible, Visibility.Collapsed)
        ReportsContentView.Visibility = If(section = AdminDashboardSection.Reports, Visibility.Visible, Visibility.Collapsed)
        SchedulingContentView.Visibility = If(section = AdminDashboardSection.Scheduling, Visibility.Visible, Visibility.Collapsed)

        ContentTitleTextBlock.Text = GetSectionTitle(section)
        SearchPlaceholderTextBlock.Text = GetSectionSearchPlaceholder(section)
        UpdateHeaderSectionAction(section)

        ApplySidebarSelectionStyles(section)

        RefreshSectionData(section)
        ApplySearchFilterToSection(section, SearchActionsTextBox.Text)
    End Sub

    Private Sub HeaderSectionActionButton_Click(sender As Object, e As RoutedEventArgs)
        If _activeSection = AdminDashboardSection.Students Then
            StudentsContentView.OpenAddStudentFormFromDashboard()
        End If
    End Sub

    Private Sub SearchActionsTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        ApplySearchFilterToSection(_activeSection, SearchActionsTextBox.Text)
    End Sub

    Private Sub ApplySidebarSelectionStyles(activeSection As AdminDashboardSection)
        ApplySidebarButtonState(DashboardNavButton, DashboardNavIconBorder, DashboardNavIconText, DashboardNavText, activeSection = AdminDashboardSection.Dashboard)
        ApplySidebarButtonState(StudentsNavButton, StudentsNavIconBorder, StudentsNavIconText, StudentsNavText, activeSection = AdminDashboardSection.Students)
        ApplySidebarButtonState(TeachersNavButton, TeachersNavIconBorder, TeachersNavIconText, TeachersNavText, activeSection = AdminDashboardSection.Teachers)
        ApplySidebarButtonState(AdministratorsNavButton, AdministratorsNavIconBorder, AdministratorsNavIconText, AdministratorsNavText, activeSection = AdminDashboardSection.Administrators)
        ApplySidebarButtonState(CoursesNavButton, CoursesNavIconBorder, CoursesNavIconText, CoursesNavText, activeSection = AdminDashboardSection.Courses)
        ApplySidebarButtonState(DepartmentsNavButton, DepartmentsNavIconBorder, DepartmentsNavIconText, DepartmentsNavText, activeSection = AdminDashboardSection.Departments)
        ApplySidebarButtonState(SubjectsNavButton, SubjectsNavIconBorder, SubjectsNavIconText, SubjectsNavText, activeSection = AdminDashboardSection.Subjects)
        ApplySidebarButtonState(ReportsNavButton, ReportsNavIconBorder, ReportsNavIconText, ReportsNavText, activeSection = AdminDashboardSection.Reports)
        ApplySidebarButtonState(SchedulingNavButton, SchedulingNavIconBorder, SchedulingNavIconText, SchedulingNavText, activeSection = AdminDashboardSection.Scheduling)
    End Sub

    Private Sub UpdateHeaderSectionAction(section As AdminDashboardSection)
        If HeaderSectionActionButton Is Nothing Then
            Return
        End If

        HeaderSectionActionButton.Visibility = Visibility.Collapsed
    End Sub

    Private Sub RefreshSectionData(section As AdminDashboardSection)
        Select Case section
            Case AdminDashboardSection.Dashboard
                DashboardContentView.RefreshData()
            Case AdminDashboardSection.Students
                StudentsContentView.RefreshData()
            Case AdminDashboardSection.Teachers
                TeachersContentView.RefreshData()
            Case AdminDashboardSection.Administrators
                AdministratorsContentView.RefreshData()
            Case AdminDashboardSection.Courses
                CoursesContentView.RefreshData()
            Case AdminDashboardSection.Departments
                DepartmentsContentView.RefreshData()
            Case AdminDashboardSection.Subjects
                SubjectsContentView.RefreshData()
            Case AdminDashboardSection.Reports
                ReportsContentView.RefreshData()
            Case AdminDashboardSection.Scheduling
                SchedulingContentView.RefreshData()
        End Select
    End Sub

    Private Sub ApplySearchFilterToSection(section As AdminDashboardSection, searchTerm As String)
        Select Case section
            Case AdminDashboardSection.Students
                StudentsContentView.ApplySearchFilter(searchTerm)
            Case AdminDashboardSection.Teachers
                TeachersContentView.ApplySearchFilter(searchTerm)
            Case AdminDashboardSection.Administrators
                AdministratorsContentView.ApplySearchFilter(searchTerm)
            Case AdminDashboardSection.Courses
                CoursesContentView.ApplySearchFilter(searchTerm)
            Case AdminDashboardSection.Departments
                DepartmentsContentView.ApplySearchFilter(searchTerm)
            Case AdminDashboardSection.Subjects
                SubjectsContentView.ApplySearchFilter(searchTerm)
            Case AdminDashboardSection.Reports
                ReportsContentView.ApplySearchFilter(searchTerm)
            Case AdminDashboardSection.Scheduling
                SchedulingContentView.ApplySearchFilter(searchTerm)
        End Select
    End Sub

    Private Sub ApplySidebarButtonState(navButton As Button, navIconBorder As Border, navIconText As TextBlock, navText As TextBlock, isSelected As Boolean)
        navButton.Style = CType(FindResource(If(isSelected, "DashboardSidebarNavSelectedButtonStyle", "DashboardSidebarNavButtonStyle")), Style)
        navIconBorder.Style = CType(FindResource(If(isSelected, "DashboardSidebarIconSelectedBadgeStyle", "DashboardSidebarIconBadgeStyle")), Style)
        navIconText.Style = CType(FindResource(If(isSelected, "DashboardSidebarIconGlyphSelectedStyle", "DashboardSidebarIconGlyphStyle")), Style)
        navText.Style = CType(FindResource(If(isSelected, "DashboardSidebarNavTextSelectedStyle", "DashboardSidebarNavTextStyle")), Style)
    End Sub

    Private Function GetSectionTitle(section As AdminDashboardSection) As String
        Select Case section
            Case AdminDashboardSection.Students
                Return "Students"
            Case AdminDashboardSection.Teachers
                Return "Teachers"
            Case AdminDashboardSection.Administrators
                Return "Administrators"
            Case AdminDashboardSection.Courses
                Return "Courses"
            Case AdminDashboardSection.Departments
                Return "Departments"
            Case AdminDashboardSection.Subjects
                Return "Subjects"
            Case AdminDashboardSection.Reports
                Return "Reports"
            Case AdminDashboardSection.Scheduling
                Return "Scheduling"
            Case Else
                Return "Overview"
        End Select
    End Function

    Private Function GetSectionSearchPlaceholder(section As AdminDashboardSection) As String
        Select Case section
            Case AdminDashboardSection.Students
                Return "Search students..."
            Case AdminDashboardSection.Teachers
                Return "Search teachers..."
            Case AdminDashboardSection.Administrators
                Return "Search administrators..."
            Case AdminDashboardSection.Courses
                Return "Search courses..."
            Case AdminDashboardSection.Departments
                Return "Search departments..."
            Case AdminDashboardSection.Subjects
                Return "Search subjects..."
            Case AdminDashboardSection.Reports
                Return "Search reports..."
            Case AdminDashboardSection.Scheduling
                Return "Search schedules..."
            Case Else
                Return "Search students, staff, or actions..."
        End Select
    End Function

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

    Private Sub UpdateMaximizeRestoreIcon()
        If MaximizeRestoreIcon Is Nothing Then
            Return
        End If

        MaximizeRestoreIcon.Text = If(WindowState = WindowState.Maximized, ChrW(&HE923), ChrW(&HE922))
    End Sub
End Class
