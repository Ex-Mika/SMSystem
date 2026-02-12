Class LoginWindow
    Private _isPasswordVisible As Boolean
    Private _isSyncingPassword As Boolean

    Public Sub New()
        InitializeComponent()
        UpdateMaximizeRestoreIcon()
        SetPasswordVisibility(False, False)
        UpdateIdentifierLabel()
        UpdatePasswordPlaceholderVisibility()
    End Sub

    Private Sub RoleSelector_Checked(sender As Object, e As RoutedEventArgs)
        UpdateIdentifierLabel()
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

        If Not _isPasswordVisible Then
            SyncVisibleTextFromPasswordBox()
        End If

        UpdatePasswordPlaceholderVisibility()
    End Sub

    Private Sub PasswordVisibleTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        If _isSyncingPassword Then
            Return
        End If

        If _isPasswordVisible Then
            SyncPasswordBoxFromVisibleText()
        End If

        UpdatePasswordPlaceholderVisibility()
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
            Return
        End If

        If TeacherRoleRadioButton.IsChecked = True Then
            IdentifierLabelTextBlock.Text = "Employee ID"
            Return
        End If

        IdentifierLabelTextBlock.Text = "Email"
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
