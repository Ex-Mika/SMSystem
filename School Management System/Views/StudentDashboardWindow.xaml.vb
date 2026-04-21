Imports System.Collections.Generic
Imports System.Data
Imports System.Windows.Media
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class StudentDashboardWindow
    Private Enum StudentDashboardSection
        Dashboard
        MySubjects
        ClassSchedule
        Grades
        Profile
    End Enum

    Private ReadOnly _studentManagementService As New StudentManagementService()
    Private ReadOnly _headerStudentAvatarFallbackBrush As Brush

    Public Property LoggedInStudentId As String = String.Empty
    Public Property LoggedInStudentName As String = String.Empty
    Public Property LoggedInStudentPhotoPath As String = String.Empty

    Public Sub New()
        InitializeComponent()
        _headerStudentAvatarFallbackBrush = HeaderStudentAvatarBorder.Background
        AddHandler ClassScheduleContentView.TimetableChanged, AddressOf ClassScheduleContentView_TimetableChanged
        AddHandler ProfileContentView.ProfileUpdated, AddressOf ProfileContentView_ProfileUpdated
        UpdateMaximizeRestoreIcon()
        ApplyLoggedInStudentProfile()
        LoadDashboardTimetable()
        SetActiveSection(StudentDashboardSection.Dashboard)
    End Sub

    Public Sub SetLoggedInStudent(studentId As String,
                                  Optional studentName As String = "",
                                  Optional studentPhotoPath As String = "")
        LoggedInStudentId = If(studentId, String.Empty).Trim()
        LoggedInStudentName = If(studentName, String.Empty).Trim()
        LoggedInStudentPhotoPath = If(studentPhotoPath, String.Empty).Trim()
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

        If isGradesSelected Then
            GradesContentView.SetStudentContext(LoggedInStudentId, LoggedInStudentName)
        End If

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

    Private Sub ClassScheduleContentView_TimetableChanged(sender As Object, e As EventArgs)
        Dim studentId As String = If(LoggedInStudentId, String.Empty).Trim()
        Dim studentName As String = If(LoggedInStudentName, String.Empty).Trim()
        Dim studentRecord As StudentRecord = LoadStudentRecord(studentId)
        Dim displayName As String = ResolveDisplayName(studentRecord, studentName)
        Dim photoPath As String = ResolveStudentPhotoPath(studentRecord, LoggedInStudentPhotoPath)

        ApplyStudentShellDetails(studentRecord, studentId, displayName, photoPath)
        SetDashboardTimetable(ClassScheduleContentView.GetCurrentTimetableSnapshot())
        MySubjectsContentView.SetStudentContext(studentId, displayName)
        ProfileContentView.SetStudentContext(studentId, displayName, photoPath)
    End Sub

    Private Sub ProfileContentView_ProfileUpdated(sender As Object,
                                                  e As StudentProfileView.StudentProfileUpdatedEventArgs)
        If e Is Nothing OrElse e.Student Is Nothing Then
            Return
        End If

        LoggedInStudentId = If(e.Student.StudentNumber, LoggedInStudentId).Trim()
        LoggedInStudentName = ResolveDisplayName(e.Student, LoggedInStudentName)
        LoggedInStudentPhotoPath = ResolveStudentPhotoPath(e.Student, LoggedInStudentPhotoPath)

        ApplyStudentShellDetails(e.Student,
                                 LoggedInStudentId,
                                 LoggedInStudentName,
                                 LoggedInStudentPhotoPath)
        MySubjectsContentView.SetStudentContext(LoggedInStudentId, LoggedInStudentName)
        ClassScheduleContentView.SetStudentContext(LoggedInStudentId, LoggedInStudentName)
        GradesContentView.SetStudentContext(LoggedInStudentId, LoggedInStudentName)
    End Sub

    Private Sub ApplySidebarSelectionStyles(activeSection As StudentDashboardSection)
        ApplySidebarButtonState(DashboardNavButton, DashboardNavIconBorder, DashboardNavIconText, DashboardNavText, activeSection = StudentDashboardSection.Dashboard)
        ApplySidebarButtonState(MySubjectsNavButton, MySubjectsNavIconBorder, MySubjectsNavIconText, MySubjectsNavText, activeSection = StudentDashboardSection.MySubjects)
        ApplySidebarButtonState(ClassScheduleNavButton, ClassScheduleNavIconBorder, ClassScheduleNavIconText, ClassScheduleNavText, activeSection = StudentDashboardSection.ClassSchedule)
        ApplySidebarButtonState(GradesNavButton, GradesNavIconBorder, GradesNavIconText, GradesNavText, activeSection = StudentDashboardSection.Grades)
        ApplySidebarButtonState(ProfileNavButton, ProfileNavIconBorder, ProfileNavIconText, ProfileNavText, activeSection = StudentDashboardSection.Profile)
    End Sub

    Private Sub ApplySidebarButtonState(navButton As Button,
                                        navIconBorder As Border,
                                        navIconText As TextBlock,
                                        navText As TextBlock,
                                        isSelected As Boolean)
        navButton.Style = CType(FindResource(If(isSelected,
                                                "DashboardSidebarNavSelectedButtonStyle",
                                                "DashboardSidebarNavButtonStyle")), Style)
        navIconBorder.Style = CType(FindResource(If(isSelected,
                                                    "DashboardSidebarIconSelectedBadgeStyle",
                                                    "DashboardSidebarIconBadgeStyle")), Style)
        navIconText.Style = CType(FindResource(If(isSelected,
                                                  "DashboardSidebarIconGlyphSelectedStyle",
                                                  "DashboardSidebarIconGlyphStyle")), Style)
        navText.Style = CType(FindResource(If(isSelected,
                                              "DashboardSidebarNavTextSelectedStyle",
                                              "DashboardSidebarNavTextStyle")), Style)
    End Sub

    Private Sub ApplyLoggedInStudentProfile()
        Dim studentId As String = If(LoggedInStudentId, String.Empty).Trim()
        Dim studentName As String = If(LoggedInStudentName, String.Empty).Trim()
        Dim studentRecord As StudentRecord = LoadStudentRecord(studentId)
        Dim displayName As String = ResolveDisplayName(studentRecord, studentName)
        Dim photoPath As String = ResolveStudentPhotoPath(studentRecord, LoggedInStudentPhotoPath)

        ApplyStudentShellDetails(studentRecord, studentId, displayName, photoPath)
        MySubjectsContentView.SetStudentContext(studentId, displayName)
        ClassScheduleContentView.SetStudentContext(studentId, displayName)
        GradesContentView.SetStudentContext(studentId, displayName)
        ProfileContentView.SetStudentContext(studentId, displayName, photoPath)
    End Sub

    Private Sub ApplyStudentShellDetails(student As StudentRecord,
                                         fallbackStudentId As String,
                                         displayName As String,
                                         photoPath As String)
        HeaderStudentNameTextBlock.Text = displayName
        HeaderStudentIdTextBlock.Text = ResolveHeaderAcademicSummary(student, fallbackStudentId)
        SidebarStudentSubtitleTextBlock.Text = ResolveSidebarSubtitle(student, fallbackStudentId)
        HeaderStudentInitialTextBlock.Text = BuildStudentInitial(displayName, fallbackStudentId)
        DashboardProfileImageHelper.ApplyProfilePhoto(HeaderStudentAvatarBorder,
                                                      HeaderStudentInitialTextBlock,
                                                      photoPath,
                                                      _headerStudentAvatarFallbackBrush)
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
        Dim timetableTable As DataTable = FetchDashboardTimetableForStudent(LoggedInStudentId)
        If timetableTable Is Nothing Then
            timetableTable = StudentTimetablePresenter.CreateEmptyTable()
        End If

        SetDashboardTimetable(timetableTable)
    End Sub

    Private Function FetchDashboardTimetableForStudent(studentId As String) As DataTable
        If ClassScheduleContentView Is Nothing Then
            Return StudentTimetablePresenter.CreateEmptyTable()
        End If

        Return ClassScheduleContentView.GetCurrentTimetableSnapshot()
    End Function

    Private Sub UpdateMaximizeRestoreIcon()
        If MaximizeRestoreIcon Is Nothing Then
            Return
        End If

        MaximizeRestoreIcon.Text = If(WindowState = WindowState.Maximized, ChrW(&HE923), ChrW(&HE922))
    End Sub

    Private Function LoadStudentRecord(studentId As String) As StudentRecord
        If String.IsNullOrWhiteSpace(studentId) Then
            Return Nothing
        End If

        Dim result = _studentManagementService.GetStudentByStudentNumber(studentId)
        If result Is Nothing OrElse Not result.IsSuccess Then
            Return Nothing
        End If

        Return result.Data
    End Function

    Private Function ResolveDisplayName(student As StudentRecord,
                                        fallbackName As String) As String
        If student IsNot Nothing Then
            Dim fullName As String = If(student.FullName, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(fullName) Then
                Return fullName
            End If
        End If

        Dim normalizedFallbackName As String = If(fallbackName, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedFallbackName) Then
            Return "Student"
        End If

        Return normalizedFallbackName
    End Function

    Private Function ResolveStudentPhotoPath(student As StudentRecord,
                                             fallbackPhotoPath As String) As String
        If student IsNot Nothing Then
            Dim normalizedPhotoPath As String = If(student.PhotoPath, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(normalizedPhotoPath) Then
                Return normalizedPhotoPath
            End If
        End If

        Return If(fallbackPhotoPath, String.Empty).Trim()
    End Function

    Private Function BuildStudentInitial(displayName As String,
                                         fallbackStudentId As String) As String
        Dim sourceText As String = If(displayName, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(sourceText) Then
            sourceText = If(fallbackStudentId, String.Empty).Trim()
        End If

        If String.IsNullOrWhiteSpace(sourceText) Then
            Return "S"
        End If

        Return sourceText.Substring(0, 1).ToUpperInvariant()
    End Function

    Private Function ResolveHeaderAcademicSummary(student As StudentRecord,
                                                  fallbackStudentId As String) As String
        Dim summaryParts As New List(Of String)()

        If student IsNot Nothing Then
            Dim courseName As String = If(student.CourseDisplayName, String.Empty).Trim()
            Dim yearLabel As String = If(student.YearLevelLabel, String.Empty).Trim()
            Dim sectionLabel As String = BuildSectionDisplayLabel(student)

            If Not String.IsNullOrWhiteSpace(courseName) Then
                summaryParts.Add(courseName)
            End If

            If Not String.IsNullOrWhiteSpace(yearLabel) Then
                summaryParts.Add(yearLabel)
            End If

            If Not String.IsNullOrWhiteSpace(sectionLabel) Then
                summaryParts.Add(sectionLabel)
            End If
        End If

        If summaryParts.Count = 0 Then
            Return If(String.IsNullOrWhiteSpace(fallbackStudentId), " ", fallbackStudentId)
        End If

        Return String.Join(" | ", summaryParts)
    End Function

    Private Function ResolveSidebarSubtitle(student As StudentRecord,
                                            fallbackStudentId As String) As String
        If student IsNot Nothing Then
            Dim yearLabel As String = If(student.YearLevelLabel, String.Empty).Trim()
            Dim courseName As String = If(student.CourseDisplayName, String.Empty).Trim()

            If Not String.IsNullOrWhiteSpace(courseName) AndAlso
               Not String.IsNullOrWhiteSpace(yearLabel) Then
                Return courseName & " | " & yearLabel
            End If

            If Not String.IsNullOrWhiteSpace(courseName) Then
                Return courseName
            End If

            If Not String.IsNullOrWhiteSpace(yearLabel) Then
                Return yearLabel
            End If
        End If

        If String.IsNullOrWhiteSpace(fallbackStudentId) Then
            Return "Student Dashboard"
        End If

        Return fallbackStudentId
    End Function

    Private Function BuildSectionDisplayLabel(student As StudentRecord) As String
        Return StudentScheduleHelper.BuildStudentSectionValue(student, String.Empty)
    End Function
End Class
