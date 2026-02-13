Class AdminDashboardWindow
    Private _currentNavButton As Button

    ' Cached page instances to preserve state across navigation
    Private _dashboardPage As DashboardPage
    Private _coursesPage As CoursesPage
    Private _subjectsPage As SubjectsPage
    Private _sectionsPage As SectionsPage
    Private _roomsPage As RoomsPage
    Private _studentsPage As StudentsPage
    Private _teachersPage As TeachersPage
    Private _schedulingPage As SchedulingPage

    Public Sub New()
        InitializeComponent()
        UpdateMaximizeRestoreIcon()
        _currentNavButton = NavDashboardButton
        _dashboardPage = New DashboardPage()
        MainContentHost.Content = _dashboardPage
    End Sub

    Private Sub NavigateTo(page As UserControl, selectedButton As Button, pageTitle As String)
        If selectedButton Is _currentNavButton Then Return

        MainContentHost.Content = page
        PageTitleText.Text = pageTitle

        ' Reset previous button to unselected style
        If _currentNavButton IsNot Nothing Then
            _currentNavButton.Style = CType(FindResource("DashboardSidebarNavButtonStyle"), Style)
        End If

        ' Set new button to selected style
        selectedButton.Style = CType(FindResource("DashboardSidebarNavSelectedButtonStyle"), Style)
        _currentNavButton = selectedButton
    End Sub

    Private Sub NavDashboard_Click(sender As Object, e As RoutedEventArgs)
        NavigateTo(_dashboardPage, NavDashboardButton, "Overview")
    End Sub

    Private Sub NavCourses_Click(sender As Object, e As RoutedEventArgs)
        If _coursesPage Is Nothing Then _coursesPage = New CoursesPage()
        NavigateTo(_coursesPage, NavCoursesButton, "Courses")
    End Sub

    Private Sub NavSubjects_Click(sender As Object, e As RoutedEventArgs)
        If _subjectsPage Is Nothing Then _subjectsPage = New SubjectsPage()
        NavigateTo(_subjectsPage, NavSubjectsButton, "Subjects")
    End Sub

    Private Sub NavSections_Click(sender As Object, e As RoutedEventArgs)
        If _sectionsPage Is Nothing Then _sectionsPage = New SectionsPage()
        NavigateTo(_sectionsPage, NavSectionsButton, "Sections")
    End Sub

    Private Sub NavRooms_Click(sender As Object, e As RoutedEventArgs)
        If _roomsPage Is Nothing Then _roomsPage = New RoomsPage()
        NavigateTo(_roomsPage, NavRoomsButton, "Rooms")
    End Sub

    Private Sub NavStudents_Click(sender As Object, e As RoutedEventArgs)
        If _studentsPage Is Nothing Then _studentsPage = New StudentsPage()
        NavigateTo(_studentsPage, NavStudentsButton, "Students")
    End Sub

    Private Sub NavTeachers_Click(sender As Object, e As RoutedEventArgs)
        If _teachersPage Is Nothing Then _teachersPage = New TeachersPage()
        NavigateTo(_teachersPage, NavTeachersButton, "Teachers")
    End Sub

    Private Sub NavScheduling_Click(sender As Object, e As RoutedEventArgs)
        If _schedulingPage Is Nothing Then _schedulingPage = New SchedulingPage()
        NavigateTo(_schedulingPage, NavSchedulingButton, "Scheduling")
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
