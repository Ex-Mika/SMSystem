Imports System.Collections.Generic
Imports System.Data
Imports System.Windows.Media
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class TeacherDashboardWindow
    Private Enum TeacherDashboardSection
        Dashboard
        Schedule
        Grading
        Profile
    End Enum

    Private ReadOnly _teacherManagementService As New TeacherManagementService()
    Private ReadOnly _headerTeacherAvatarFallbackBrush As Brush

    Public Property LoggedInTeacherId As String = String.Empty
    Public Property LoggedInTeacherName As String = String.Empty
    Public Property LoggedInTeacherPhotoPath As String = String.Empty

    Public Sub New()
        InitializeComponent()
        _headerTeacherAvatarFallbackBrush = HeaderTeacherAvatarBorder.Background
        AddHandler ProfileContentView.ProfileUpdated, AddressOf ProfileContentView_ProfileUpdated
        UpdateMaximizeRestoreIcon()
        ApplyLoggedInTeacherProfile()
        LoadDashboardTimetable()
        SetActiveSection(TeacherDashboardSection.Dashboard)
    End Sub

    Public Sub SetLoggedInTeacher(teacherId As String,
                                  Optional teacherName As String = "",
                                  Optional teacherPhotoPath As String = "")
        LoggedInTeacherId = If(teacherId, String.Empty).Trim()
        LoggedInTeacherName = If(teacherName, String.Empty).Trim()
        LoggedInTeacherPhotoPath = If(teacherPhotoPath, String.Empty).Trim()
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

    Private Sub ScheduleNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(TeacherDashboardSection.Schedule)
    End Sub

    Private Sub GradingNavButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(TeacherDashboardSection.Grading)
    End Sub

    Private Sub ProfileContentView_ProfileUpdated(sender As Object,
                                                  e As TeacherProfileView.TeacherProfileUpdatedEventArgs)
        If e Is Nothing OrElse e.Teacher Is Nothing Then
            Return
        End If

        LoggedInTeacherId = ResolveTeacherId(e.Teacher, LoggedInTeacherId)
        LoggedInTeacherName = ResolveDisplayName(e.Teacher, LoggedInTeacherName)
        LoggedInTeacherPhotoPath = ResolveTeacherPhotoPath(e.Teacher, LoggedInTeacherPhotoPath)

        ApplyTeacherShellDetails(e.Teacher,
                                 LoggedInTeacherId,
                                 LoggedInTeacherName,
                                 LoggedInTeacherPhotoPath)
        ScheduleContentView.SetTeacherContext(LoggedInTeacherId, LoggedInTeacherName)
        GradingContentView.SetTeacherContext(LoggedInTeacherId, LoggedInTeacherName)
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
        Dim isScheduleSelected As Boolean = section = TeacherDashboardSection.Schedule
        Dim isGradingSelected As Boolean = section = TeacherDashboardSection.Grading
        Dim isProfileSelected As Boolean = section = TeacherDashboardSection.Profile

        DashboardContentScrollViewer.Visibility = If(isDashboardSelected, Visibility.Visible, Visibility.Collapsed)
        ScheduleContentView.Visibility = If(isScheduleSelected, Visibility.Visible, Visibility.Collapsed)
        GradingContentView.Visibility = If(isGradingSelected, Visibility.Visible, Visibility.Collapsed)
        ProfileContentView.Visibility = If(isProfileSelected, Visibility.Visible, Visibility.Collapsed)

        If isDashboardSelected Then
            ContentTitleTextBlock.Text = "Overview"
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
        Dim teacherRecord As TeacherRecord = LoadTeacherRecord(teacherId)
        Dim resolvedTeacherId As String = ResolveTeacherId(teacherRecord, teacherId)
        Dim displayName As String = ResolveDisplayName(teacherRecord, teacherName)
        Dim photoPath As String = ResolveTeacherPhotoPath(teacherRecord, LoggedInTeacherPhotoPath)

        LoggedInTeacherId = resolvedTeacherId
        LoggedInTeacherName = displayName
        LoggedInTeacherPhotoPath = photoPath

        ApplyTeacherShellDetails(teacherRecord, resolvedTeacherId, displayName, photoPath)
        ScheduleContentView.SetTeacherContext(resolvedTeacherId, displayName)
        GradingContentView.SetTeacherContext(resolvedTeacherId, displayName)
        ProfileContentView.SetTeacherContext(resolvedTeacherId, displayName, photoPath)
    End Sub

    Private Sub ApplyTeacherShellDetails(teacher As TeacherRecord,
                                         fallbackTeacherId As String,
                                         displayName As String,
                                         photoPath As String)
        HeaderTeacherNameTextBlock.Text = displayName
        HeaderTeacherIdTextBlock.Text = ResolveHeaderSummary(teacher, fallbackTeacherId)
        SidebarTeacherSubtitleTextBlock.Text = ResolveSidebarSubtitle(teacher, fallbackTeacherId)
        HeaderTeacherInitialTextBlock.Text = BuildTeacherInitial(displayName, fallbackTeacherId)
        DashboardProfileImageHelper.ApplyProfilePhoto(HeaderTeacherAvatarBorder,
                                                      HeaderTeacherInitialTextBlock,
                                                      photoPath,
                                                      _headerTeacherAvatarFallbackBrush)
    End Sub

    Public Sub SetDashboardTimetable(table As DataTable)
        Dim resolvedTable As DataTable = table
        If resolvedTable Is Nothing OrElse resolvedTable.Columns.Count = 0 Then
            resolvedTable = StudentTimetablePresenter.CreateEmptyTable()
        End If

        StudentTimetablePresenter.ConfigureDataGrid(DashboardTimetableDataGrid, resolvedTable)
        DashboardTimetableDataGrid.ItemsSource = resolvedTable.DefaultView
    End Sub

    Private Sub LoadDashboardTimetable()
        Dim timetableTable As DataTable = FetchDashboardTimetableForTeacher(LoggedInTeacherId)
        If timetableTable Is Nothing Then
            timetableTable = StudentTimetablePresenter.CreateEmptyTable()
        End If

        SetDashboardTimetable(timetableTable)
    End Sub

    Private Function FetchDashboardTimetableForTeacher(teacherId As String) As DataTable
        If ScheduleContentView Is Nothing Then
            Return StudentTimetablePresenter.CreateEmptyTable()
        End If

        Dim timetableSnapshot As DataTable = ScheduleContentView.GetTimetableSnapshot()
        If timetableSnapshot Is Nothing OrElse timetableSnapshot.Columns.Count = 0 Then
            Return StudentTimetablePresenter.CreateEmptyTable()
        End If

        Return timetableSnapshot
    End Function

    Private Sub UpdateMaximizeRestoreIcon()
        If MaximizeRestoreIcon Is Nothing Then
            Return
        End If

        MaximizeRestoreIcon.Text = If(WindowState = WindowState.Maximized, ChrW(&HE923), ChrW(&HE922))
    End Sub

    Private Function LoadTeacherRecord(teacherId As String) As TeacherRecord
        If String.IsNullOrWhiteSpace(teacherId) Then
            Return Nothing
        End If

        Dim result = _teacherManagementService.GetTeacherByEmployeeNumber(teacherId)
        If result Is Nothing OrElse Not result.IsSuccess Then
            Return Nothing
        End If

        Return result.Data
    End Function

    Private Function ResolveDisplayName(teacher As TeacherRecord,
                                        fallbackName As String) As String
        If teacher IsNot Nothing Then
            Dim fullName As String = If(teacher.FullName, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(fullName) Then
                Return fullName
            End If
        End If

        Dim normalizedFallbackName As String = If(fallbackName, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedFallbackName) Then
            Return "Teacher"
        End If

        Return normalizedFallbackName
    End Function

    Private Function ResolveTeacherPhotoPath(teacher As TeacherRecord,
                                             fallbackPhotoPath As String) As String
        If teacher IsNot Nothing Then
            Dim normalizedPhotoPath As String = If(teacher.PhotoPath, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(normalizedPhotoPath) Then
                Return normalizedPhotoPath
            End If
        End If

        Return If(fallbackPhotoPath, String.Empty).Trim()
    End Function

    Private Function ResolveHeaderSummary(teacher As TeacherRecord,
                                          fallbackTeacherId As String) As String
        Dim summaryParts As New List(Of String)()
        Dim teacherId As String = ResolveTeacherId(teacher, fallbackTeacherId)
        Dim department As String = ResolveDepartment(teacher)
        Dim position As String = ResolvePosition(teacher)

        If Not String.IsNullOrWhiteSpace(teacherId) Then
            summaryParts.Add(teacherId)
        End If

        If Not String.IsNullOrWhiteSpace(department) Then
            summaryParts.Add(department)
        End If

        If Not String.IsNullOrWhiteSpace(position) Then
            summaryParts.Add(position)
        End If

        If summaryParts.Count = 0 Then
            Return " "
        End If

        Return String.Join(" | ", summaryParts)
    End Function

    Private Function ResolveSidebarSubtitle(teacher As TeacherRecord,
                                            fallbackTeacherId As String) As String
        Dim teacherId As String = ResolveTeacherId(teacher, fallbackTeacherId)
        If Not String.IsNullOrWhiteSpace(teacherId) Then
            Return teacherId
        End If

        Dim department As String = ResolveDepartment(teacher)
        If Not String.IsNullOrWhiteSpace(department) Then
            Return department
        End If

        Return "Teacher Dashboard"
    End Function

    Private Function ResolveTeacherId(teacher As TeacherRecord,
                                      fallbackTeacherId As String) As String
        If teacher IsNot Nothing Then
            Dim employeeNumber As String = If(teacher.EmployeeNumber, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(employeeNumber) Then
                Return employeeNumber
            End If
        End If

        Return If(fallbackTeacherId, String.Empty).Trim()
    End Function

    Private Function ResolveDepartment(teacher As TeacherRecord) As String
        If teacher Is Nothing Then
            Return String.Empty
        End If

        Return If(teacher.DepartmentDisplayName, String.Empty).Trim()
    End Function

    Private Function ResolvePosition(teacher As TeacherRecord) As String
        If teacher Is Nothing Then
            Return String.Empty
        End If

        Return If(teacher.PositionTitle, String.Empty).Trim()
    End Function

    Private Function BuildTeacherInitial(displayName As String,
                                         fallbackTeacherId As String) As String
        Dim sourceText As String = If(displayName, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(sourceText) Then
            sourceText = If(fallbackTeacherId, String.Empty).Trim()
        End If

        If String.IsNullOrWhiteSpace(sourceText) Then
            Return "T"
        End If

        Return sourceText.Substring(0, 1).ToUpperInvariant()
    End Function
End Class

