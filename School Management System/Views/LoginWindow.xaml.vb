Imports System.Threading.Tasks
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class LoginWindow
    Private ReadOnly _authenticationService As New AuthenticationService()
    Private _isPasswordVisible As Boolean
    Private _isSyncingPassword As Boolean
    Private _isAuthenticating As Boolean

    Public Sub New()
        InitializeComponent()
        UpdateMaximizeRestoreIcon()
        SetPasswordVisibility(False, False)
        UpdateIdentifierLabel()
        UpdatePasswordPlaceholderVisibility()
    End Sub

    Private Sub RoleSelector_Checked(sender As Object, e As RoutedEventArgs)
        UpdateIdentifierLabel()
        ClearLoginStatus()
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

    Private Sub LoginWindow_StateChanged(sender As Object, e As EventArgs)
        UpdateMaximizeRestoreIcon()
    End Sub

    Private Sub TogglePasswordVisibility_Click(sender As Object, e As RoutedEventArgs)
        SetPasswordVisibility(Not _isPasswordVisible)
    End Sub

    Private Sub PasswordInput_PasswordChanged(sender As Object, e As RoutedEventArgs)
        If _isSyncingPassword Then
            Return
        End If

        ClearLoginStatus()

        If Not _isPasswordVisible Then
            SyncVisibleTextFromPasswordBox()
        End If

        UpdatePasswordPlaceholderVisibility()
    End Sub

    Private Sub PasswordVisibleTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        If _isSyncingPassword Then
            Return
        End If

        ClearLoginStatus()

        If _isPasswordVisible Then
            SyncPasswordBoxFromVisibleText()
        End If

        UpdatePasswordPlaceholderVisibility()
    End Sub

    Private Sub UsernameTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        ClearLoginStatus()
    End Sub

    Private Sub PasswordControl_GotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        UpdatePasswordPlaceholderVisibility()
    End Sub

    Private Sub PasswordControl_LostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        UpdatePasswordPlaceholderVisibility()
    End Sub

    Private Sub SetPasswordVisibility(isVisible As Boolean, Optional applyFocus As Boolean = True)
        _isPasswordVisible = isVisible

        If _isPasswordVisible Then
            SyncVisibleTextFromPasswordBox()
            PasswordInput.Visibility = Visibility.Collapsed
            PasswordVisibleTextBox.Visibility = Visibility.Visible
            If applyFocus Then
                PasswordVisibleTextBox.Focus()
                PasswordVisibleTextBox.CaretIndex = PasswordVisibleTextBox.Text.Length
            End If
        Else
            SyncPasswordBoxFromVisibleText()
            PasswordVisibleTextBox.Visibility = Visibility.Collapsed
            PasswordInput.Visibility = Visibility.Visible
            If applyFocus Then
                PasswordInput.Focus()
                PasswordInput.SelectAll()
            End If
        End If
        UpdatePasswordEyeIcon()
        UpdatePasswordPlaceholderVisibility()
    End Sub

    Private Sub SyncVisibleTextFromPasswordBox()
        _isSyncingPassword = True
        PasswordVisibleTextBox.Text = PasswordInput.Password
        _isSyncingPassword = False
    End Sub

    Private Sub SyncPasswordBoxFromVisibleText()
        _isSyncingPassword = True
        PasswordInput.Password = PasswordVisibleTextBox.Text
        _isSyncingPassword = False
    End Sub

    Private Sub UpdatePasswordPlaceholderVisibility()
        Dim hasText As Boolean = If(_isPasswordVisible,
                                    Not String.IsNullOrWhiteSpace(PasswordVisibleTextBox.Text),
                                    Not String.IsNullOrWhiteSpace(PasswordInput.Password))

        Dim hasFocus As Boolean = PasswordInput.IsKeyboardFocused OrElse PasswordVisibleTextBox.IsKeyboardFocused
        PasswordPlaceholder.Visibility = If(hasText OrElse hasFocus, Visibility.Collapsed, Visibility.Visible)
    End Sub

    Private Sub UpdatePasswordEyeIcon()
        PasswordEyeOpenIcon.Visibility = If(_isPasswordVisible, Visibility.Collapsed, Visibility.Visible)
        PasswordEyeClosedIcon.Visibility = If(_isPasswordVisible, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub UpdateIdentifierLabel()
        If IdentifierLabelTextBlock Is Nothing OrElse
           StudentRoleRadioButton Is Nothing OrElse
           TeacherRoleRadioButton Is Nothing OrElse
           AdminRoleRadioButton Is Nothing Then
            Return
        End If

        If StudentRoleRadioButton.IsChecked = True Then
            IdentifierLabelTextBlock.Text = "Student ID"
            UpdateIdentifierPlaceholder("Enter your student ID")
            Return
        End If

        If TeacherRoleRadioButton.IsChecked = True Then
            IdentifierLabelTextBlock.Text = "Employee ID"
            UpdateIdentifierPlaceholder("Enter your employee ID")
            Return
        End If

        IdentifierLabelTextBlock.Text = "Email"
        UpdateIdentifierPlaceholder("Enter your email")
    End Sub

    Private Async Sub LoginButton_Click(sender As Object, e As RoutedEventArgs)
        If _isAuthenticating Then
            Return
        End If

        Dim identifier As String = If(UsernameTextBox.Text, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(identifier) Then
            ShowLoginStatus("Please enter your login identifier.")
            UsernameTextBox.Focus()
            Return
        End If

        Dim password As String = GetEnteredPassword()
        If String.IsNullOrWhiteSpace(password) Then
            ShowLoginStatus("Please enter your password.")
            If _isPasswordVisible Then
                PasswordVisibleTextBox.Focus()
            Else
                PasswordInput.Focus()
            End If
            Return
        End If

        Dim request As New LoginRequest() With {
            .Role = GetSelectedRole(),
            .Identifier = identifier,
            .Password = password
        }

        SetAuthenticationState(True)
        ShowLoginStatus("Signing in...", False)

        Try
            Dim result = Await Task.Run(Function() _authenticationService.Authenticate(request))
            If result Is Nothing OrElse Not result.IsSuccess OrElse result.Data Is Nothing Then
                ShowLoginStatus(AuthenticationService.InvalidCredentialsMessage)
                Return
            End If

            OpenDashboard(result.Data, request.Identifier)
        Catch ex As MySqlException
            ShowLoginStatus("Unable to connect to MySQL on 127.0.0.1. Check your database settings.")
        Catch ex As InvalidOperationException
            ShowLoginStatus("Unable to sign in right now. Please verify the database connection.")
        Catch ex As Exception
            ShowLoginStatus("An unexpected error occurred while signing in.")
        Finally
            SetAuthenticationState(False)
        End Try
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

    Private Function GetEnteredPassword() As String
        If _isPasswordVisible Then
            Return If(PasswordVisibleTextBox.Text, String.Empty)
        End If

        Return If(PasswordInput.Password, String.Empty)
    End Function

    Private Function GetSelectedRole() As UserRole
        If StudentRoleRadioButton IsNot Nothing AndAlso StudentRoleRadioButton.IsChecked = True Then
            Return UserRole.Student
        End If

        If TeacherRoleRadioButton IsNot Nothing AndAlso TeacherRoleRadioButton.IsChecked = True Then
            Return UserRole.Teacher
        End If

        Return UserRole.Admin
    End Function

    Private Sub OpenDashboard(user As UserAccount, fallbackIdentifier As String)
        Dim resolvedIdentifier As String = If(user.ReferenceCode, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(resolvedIdentifier) Then
            resolvedIdentifier = If(fallbackIdentifier, String.Empty).Trim()
        End If

        Dim resolvedDisplayName As String = If(user.Username, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(resolvedDisplayName) Then
            resolvedDisplayName = resolvedIdentifier
        End If

        If user.Role = UserRole.Student Then
            Dim studentDashboard As New StudentDashboardWindow()
            studentDashboard.SetLoggedInStudent(resolvedIdentifier, resolvedDisplayName)
            studentDashboard.Show()
            Close()
            Return
        End If

        If user.Role = UserRole.Teacher Then
            Dim teacherDashboard As New TeacherDashboardWindow()
            teacherDashboard.SetLoggedInTeacher(resolvedIdentifier, resolvedDisplayName)
            teacherDashboard.Show()
            Close()
            Return
        End If

        Dim adminDashboard As New AdminDashboardWindow()
        adminDashboard.Show()
        Close()
    End Sub

    Private Sub SetAuthenticationState(isAuthenticating As Boolean)
        _isAuthenticating = isAuthenticating

        If LoginButton IsNot Nothing Then
            LoginButton.IsEnabled = Not isAuthenticating
        End If

        If UsernameTextBox IsNot Nothing Then
            UsernameTextBox.IsEnabled = Not isAuthenticating
        End If

        If PasswordInput IsNot Nothing Then
            PasswordInput.IsEnabled = Not isAuthenticating
        End If

        If PasswordVisibleTextBox IsNot Nothing Then
            PasswordVisibleTextBox.IsEnabled = Not isAuthenticating
        End If

        If StudentRoleRadioButton IsNot Nothing Then
            StudentRoleRadioButton.IsEnabled = Not isAuthenticating
        End If

        If TeacherRoleRadioButton IsNot Nothing Then
            TeacherRoleRadioButton.IsEnabled = Not isAuthenticating
        End If

        If AdminRoleRadioButton IsNot Nothing Then
            AdminRoleRadioButton.IsEnabled = Not isAuthenticating
        End If
    End Sub

    Private Sub ShowLoginStatus(message As String, Optional isError As Boolean = True)
        If LoginStatusTextBlock Is Nothing Then
            Return
        End If

        LoginStatusTextBlock.Text = If(message, String.Empty).Trim()
        LoginStatusTextBlock.Visibility = If(String.IsNullOrWhiteSpace(LoginStatusTextBlock.Text),
                                             Visibility.Collapsed,
                                             Visibility.Visible)
        LoginStatusTextBlock.Foreground = If(isError, Brushes.IndianRed, Brushes.SteelBlue)
    End Sub

    Private Sub ClearLoginStatus()
        ShowLoginStatus(String.Empty)
    End Sub

    Private Sub UpdateIdentifierPlaceholder(placeholderText As String)
        If IdentifierPlaceholderTextBlock Is Nothing Then
            Return
        End If

        IdentifierPlaceholderTextBlock.Text = placeholderText
    End Sub
End Class
