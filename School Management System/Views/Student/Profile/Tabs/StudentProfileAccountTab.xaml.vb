Imports School_Management_System.Backend.Models

Class StudentProfileAccountTab
    Public Event EditRequested As EventHandler

    Private _isEditMode As Boolean

    Public Sub New()
        InitializeComponent()
        UpdateEditModeState()
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        SetStudentRecord(Nothing)
    End Sub

    Public Sub SetStudentRecord(student As StudentRecord)
        Dim emailAddress As String = String.Empty

        If student IsNot Nothing Then
            emailAddress = If(student.Email, String.Empty).Trim()
        End If

        EmailValueTextBlock.Text = If(String.IsNullOrWhiteSpace(emailAddress),
                                      "Not set",
                                      emailAddress)
        AccountStatusValueTextBlock.Text = "Active"
        PortalAccessValueTextBlock.Text = "Managed by your school account"
        CurrentPasswordValueTextBlock.Text = "Hidden for security"
        ClearPasswordInputs()
    End Sub

    Public Sub SetEditMode(isEditMode As Boolean)
        _isEditMode = isEditMode

        If Not _isEditMode Then
            ClearPasswordInputs()
        End If

        UpdateEditModeState()
    End Sub

    Public Function TryReadPassword(ByRef password As String,
                                    ByRef validationMessage As String) As Boolean
        Dim newPassword As String = If(NewPasswordInput.Password, String.Empty).Trim()
        Dim confirmPassword As String = If(ConfirmPasswordInput.Password, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(newPassword) AndAlso
           String.IsNullOrWhiteSpace(confirmPassword) Then
            password = String.Empty
            validationMessage = String.Empty
            Return True
        End If

        If String.IsNullOrWhiteSpace(newPassword) Then
            password = String.Empty
            validationMessage = "Enter a new password."
            Return False
        End If

        If String.IsNullOrWhiteSpace(confirmPassword) Then
            password = String.Empty
            validationMessage = "Confirm your new password."
            Return False
        End If

        If Not String.Equals(newPassword, confirmPassword, StringComparison.Ordinal) Then
            password = String.Empty
            validationMessage = "The password confirmation does not match."
            Return False
        End If

        If newPassword.Length < 8 Then
            password = String.Empty
            validationMessage = "Passwords must be at least 8 characters."
            Return False
        End If

        password = newPassword
        validationMessage = String.Empty
        Return True
    End Function

    Private Sub ChangePasswordButton_Click(sender As Object, e As RoutedEventArgs)
        RaiseEvent EditRequested(Me, EventArgs.Empty)
    End Sub

    Private Sub UpdateEditModeState()
        PasswordEditPanel.Visibility = If(_isEditMode,
                                          Visibility.Visible,
                                          Visibility.Collapsed)
        ChangePasswordButton.Visibility = If(_isEditMode,
                                             Visibility.Collapsed,
                                             Visibility.Visible)
    End Sub

    Private Sub ClearPasswordInputs()
        NewPasswordInput.Password = String.Empty
        ConfirmPasswordInput.Password = String.Empty
    End Sub
End Class
