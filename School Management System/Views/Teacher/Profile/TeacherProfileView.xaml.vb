Imports System.Collections.Generic
Imports Microsoft.Win32
Imports System.Windows.Media
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class TeacherProfileView
    Public Class TeacherProfileUpdatedEventArgs
        Inherits EventArgs

        Public Sub New(teacher As TeacherRecord)
            Me.Teacher = teacher
        End Sub

        Public ReadOnly Property Teacher As TeacherRecord
    End Class

    Private Class TeacherProfileValues
        Public Property FirstName As String = String.Empty
        Public Property MiddleName As String = String.Empty
        Public Property LastName As String = String.Empty
        Public Property Department As String = String.Empty
        Public Property Position As String = String.Empty
        Public Property Advisory As String = String.Empty
        Public Property PhotoPath As String = String.Empty
    End Class

    Public Event ProfileUpdated As EventHandler(Of TeacherProfileUpdatedEventArgs)

    Private ReadOnly _teacherManagementService As New TeacherManagementService()
    Private ReadOnly _profileHeaderAvatarFallbackBrush As Brush

    Private _currentTeacherRecord As TeacherRecord
    Private _currentTeacherId As String = String.Empty
    Private _currentTeacherName As String = String.Empty
    Private _currentTeacherPhotoPath As String = String.Empty
    Private _selectedPhotoPath As String = String.Empty
    Private _isEditMode As Boolean

    Public Sub New()
        InitializeComponent()
        _profileHeaderAvatarFallbackBrush = ProfileTeacherAvatarBorder.Background
        UpdateEditModeState()
        UpdateNotesText()
        UpdatePhotoStatus()
    End Sub

    Public Sub SetTeacherContext(teacherId As String,
                                 teacherName As String,
                                 Optional teacherPhotoPath As String = "")
        _currentTeacherId = If(teacherId, String.Empty).Trim()
        _currentTeacherName = If(teacherName, String.Empty).Trim()
        _currentTeacherPhotoPath = If(teacherPhotoPath, String.Empty).Trim()
        _currentTeacherRecord = LoadTeacherRecord(_currentTeacherId)

        ApplyCurrentTeacherRecord()
        SetEditMode(False)
        ShowProfileStatus(String.Empty, False)
    End Sub

    Private Sub EditProfileButton_Click(sender As Object, e As RoutedEventArgs)
        If _currentTeacherRecord Is Nothing Then
            MessageBox.Show("Your profile could not be loaded right now.",
                            "Teacher Profile",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        SetEditMode(True)
        ShowProfileStatus("Update your profile details, photo, or password. " &
                          "Employee ID stays fixed.",
                          False)
    End Sub

    Private Sub SaveProfileButton_Click(sender As Object, e As RoutedEventArgs)
        If _currentTeacherRecord Is Nothing Then
            MessageBox.Show("Your profile could not be loaded right now.",
                            "Teacher Profile",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        Dim validationMessage As String = String.Empty
        Dim profileValues As TeacherProfileValues = Nothing
        If Not TryReadProfileValues(profileValues, validationMessage) Then
            ShowSaveValidationMessage(validationMessage)
            Return
        End If

        Dim updatedPassword As String = String.Empty
        If Not TryReadPassword(updatedPassword, validationMessage) Then
            ShowSaveValidationMessage(validationMessage)
            Return
        End If

        Dim request As TeacherSaveRequest =
            BuildSaveRequest(_currentTeacherRecord,
                             profileValues,
                             updatedPassword)
        Dim result = _teacherManagementService.UpdateTeacher(request)

        If result Is Nothing OrElse
           Not result.IsSuccess OrElse
           result.Data Is Nothing Then
            Dim failureMessage As String = ResolveFailureMessage(result)
            ShowProfileStatus(failureMessage, True)
            MessageBox.Show(failureMessage,
                            "Teacher Profile",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        _currentTeacherRecord = result.Data
        _currentTeacherId = ResolveEmployeeId(result.Data, _currentTeacherId)
        _currentTeacherName = ResolveDisplayName(result.Data, _currentTeacherName)
        _currentTeacherPhotoPath = ResolveTeacherPhotoPath(result.Data, _currentTeacherPhotoPath)

        ApplyCurrentTeacherRecord()
        SetEditMode(False)
        ShowProfileStatus("Profile updated.", False)

        RaiseEvent ProfileUpdated(Me,
                                  New TeacherProfileUpdatedEventArgs(_currentTeacherRecord))
    End Sub

    Private Sub CancelEditProfileButton_Click(sender As Object, e As RoutedEventArgs)
        SetEditMode(False)
        ShowProfileStatus(String.Empty, False)
    End Sub

    Private Sub EditableProfileTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        If Not _isEditMode Then
            Return
        End If

        UpdateHeaderPreview()
    End Sub

    Private Sub BrowseTeacherPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dialog As New OpenFileDialog() With {
            .Title = "Select Profile Photo",
            .Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp|All Files|*.*",
            .CheckFileExists = True,
            .CheckPathExists = True
        }

        If dialog.ShowDialog() <> True Then
            Return
        End If

        _selectedPhotoPath = If(dialog.FileName, String.Empty).Trim()
        UpdateHeaderPreview()
        UpdatePhotoStatus()
    End Sub

    Private Sub ClearTeacherPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        _selectedPhotoPath = String.Empty
        UpdateHeaderPreview()
        UpdatePhotoStatus()
    End Sub

    Private Sub ApplyCurrentTeacherRecord()
        Dim displayName As String = ResolveDisplayName(_currentTeacherRecord, _currentTeacherName)
        Dim headerDetails As String = ResolveHeaderDetails(_currentTeacherRecord, _currentTeacherId)
        Dim photoPath As String = ResolveTeacherPhotoPath(_currentTeacherRecord, _currentTeacherPhotoPath)

        _selectedPhotoPath = photoPath

        ProfileTeacherNameTextBlock.Text = displayName
        ProfileTeacherIdTextBlock.Text = headerDetails
        ProfileTeacherEmployeeIdValueTextBlock.Text =
            ResolveDisplayValue(ResolveEmployeeId(_currentTeacherRecord, _currentTeacherId),
                                "Not set")
        ProfileTeacherInitialTextBlock.Text = BuildTeacherInitial(displayName, _currentTeacherId)
        DashboardProfileImageHelper.ApplyProfilePhoto(ProfileTeacherAvatarBorder,
                                                      ProfileTeacherInitialTextBlock,
                                                      photoPath,
                                                      _profileHeaderAvatarFallbackBrush)

        ApplyFieldValues()
        CurrentPasswordValueTextBlock.Text = "Hidden for security"
        ClearPasswordInputs()
        UpdatePhotoStatus()
        UpdateNotesText()
        UpdateEditModeState()
    End Sub

    Private Sub ApplyFieldValues()
        Dim firstName As String = ResolveFirstName(_currentTeacherRecord)
        Dim middleName As String = ResolveMiddleName(_currentTeacherRecord)
        Dim lastName As String = ResolveLastName(_currentTeacherRecord, _currentTeacherName)
        Dim department As String = ResolveDepartment(_currentTeacherRecord)
        Dim position As String = ResolvePosition(_currentTeacherRecord)
        Dim advisory As String = ResolveAdvisory(_currentTeacherRecord)

        FirstNameValueTextBlock.Text = ResolveDisplayValue(firstName, "Not set")
        MiddleNameValueTextBlock.Text = ResolveDisplayValue(middleName, "Not set")
        LastNameValueTextBlock.Text = ResolveDisplayValue(lastName, "Not set")
        DepartmentValueTextBlock.Text = ResolveDisplayValue(department, "Not set")
        PositionValueTextBlock.Text = ResolveDisplayValue(position, "Not set")
        AdvisoryValueTextBlock.Text = ResolveDisplayValue(advisory, "None")

        DepartmentStatTextBlock.Text = ResolveDisplayValue(department, "Not set")
        AdvisoryStatTextBlock.Text = ResolveDisplayValue(advisory, "None")
        PositionStatTextBlock.Text = ResolveDisplayValue(position, "Not set")

        FirstNameEditTextBox.Text = firstName
        MiddleNameEditTextBox.Text = middleName
        LastNameEditTextBox.Text = lastName
        DepartmentEditTextBox.Text = department
        PositionEditTextBox.Text = position
        AdvisoryEditTextBox.Text = advisory
    End Sub

    Private Sub SetEditMode(isEditMode As Boolean)
        _isEditMode = isEditMode AndAlso _currentTeacherRecord IsNot Nothing

        If Not _isEditMode Then
            ApplyFieldValues()
            _selectedPhotoPath = ResolveTeacherPhotoPath(_currentTeacherRecord, _currentTeacherPhotoPath)
            ClearPasswordInputs()
        End If

        UpdateEditModeState()
        UpdateHeaderPreview()
        UpdatePhotoStatus()
        UpdateNotesText()
    End Sub

    Private Sub UpdateEditModeState()
        Dim canEdit As Boolean = _currentTeacherRecord IsNot Nothing
        Dim editVisibility As Visibility = If(_isEditMode,
                                              Visibility.Visible,
                                              Visibility.Collapsed)
        Dim viewVisibility As Visibility = If(_isEditMode,
                                              Visibility.Collapsed,
                                              Visibility.Visible)

        EditProfileButton.IsEnabled = canEdit
        EditProfileButton.Visibility = If(_isEditMode, Visibility.Collapsed, Visibility.Visible)
        EditProfileActionsPanel.Visibility = If(_isEditMode, Visibility.Visible, Visibility.Collapsed)
        PhotoActionPanel.Visibility = editVisibility
        PasswordEditPanel.Visibility = editVisibility

        FirstNameValueBorder.Visibility = viewVisibility
        MiddleNameValueBorder.Visibility = viewVisibility
        LastNameValueBorder.Visibility = viewVisibility
        DepartmentValueBorder.Visibility = viewVisibility
        PositionValueBorder.Visibility = viewVisibility
        AdvisoryValueBorder.Visibility = viewVisibility

        FirstNameEditTextBox.Visibility = editVisibility
        MiddleNameEditTextBox.Visibility = editVisibility
        LastNameEditTextBox.Visibility = editVisibility
        DepartmentEditTextBox.Visibility = editVisibility
        PositionEditTextBox.Visibility = editVisibility
        AdvisoryEditTextBox.Visibility = editVisibility
    End Sub

    Private Sub UpdateHeaderPreview()
        Dim displayName As String =
            If(_isEditMode,
               BuildDisplayNameFromInputs(),
               ResolveDisplayName(_currentTeacherRecord, _currentTeacherName))
        Dim photoPath As String =
            If(_isEditMode,
               If(_selectedPhotoPath, String.Empty).Trim(),
               ResolveTeacherPhotoPath(_currentTeacherRecord, _currentTeacherPhotoPath))

        ProfileTeacherNameTextBlock.Text = displayName
        ProfileTeacherIdTextBlock.Text = ResolveHeaderDetails(_currentTeacherRecord, _currentTeacherId)
        ProfileTeacherInitialTextBlock.Text = BuildTeacherInitial(displayName, _currentTeacherId)
        DashboardProfileImageHelper.ApplyProfilePhoto(ProfileTeacherAvatarBorder,
                                                      ProfileTeacherInitialTextBlock,
                                                      photoPath,
                                                      _profileHeaderAvatarFallbackBrush)
    End Sub

    Private Sub UpdatePhotoStatus()
        If String.IsNullOrWhiteSpace(_selectedPhotoPath) Then
            ProfileTeacherPhotoStatusTextBlock.Text =
                If(_isEditMode,
                   "No photo selected. Browse to add one or leave it blank to remove it.",
                   "No profile photo on file.")
            Return
        End If

        If _isEditMode AndAlso
           Not String.Equals(_selectedPhotoPath,
                             ResolveTeacherPhotoPath(_currentTeacherRecord, _currentTeacherPhotoPath),
                             StringComparison.OrdinalIgnoreCase) Then
            ProfileTeacherPhotoStatusTextBlock.Text = "Selected photo is ready to save."
            Return
        End If

        ProfileTeacherPhotoStatusTextBlock.Text = "Current profile photo from your faculty record."
    End Sub

    Private Sub UpdateNotesText()
        ProfileEditingNotesTextBlock.Text =
            If(_isEditMode,
               "You can update your name, department, position, advisory section, photo, and password here. " &
               "Employee ID stays fixed to keep your schedule and login aligned.",
               "Use Edit Profile to update your faculty details, photo, and password without leaving the dashboard.")
    End Sub

    Private Function TryReadProfileValues(ByRef values As TeacherProfileValues,
                                          ByRef validationMessage As String) As Boolean
        Dim resolvedValues As New TeacherProfileValues() With {
            .FirstName = If(FirstNameEditTextBox.Text, String.Empty).Trim(),
            .MiddleName = If(MiddleNameEditTextBox.Text, String.Empty).Trim(),
            .LastName = If(LastNameEditTextBox.Text, String.Empty).Trim(),
            .Department = If(DepartmentEditTextBox.Text, String.Empty).Trim(),
            .Position = If(PositionEditTextBox.Text, String.Empty).Trim(),
            .Advisory = If(AdvisoryEditTextBox.Text, String.Empty).Trim(),
            .PhotoPath = If(_selectedPhotoPath, String.Empty).Trim()
        }

        If String.IsNullOrWhiteSpace(resolvedValues.FirstName) Then
            validationMessage = "First Name is required."
            values = Nothing
            Return False
        End If

        If String.IsNullOrWhiteSpace(resolvedValues.LastName) Then
            validationMessage = "Last Name is required."
            values = Nothing
            Return False
        End If

        validationMessage = String.Empty
        values = resolvedValues
        Return True
    End Function

    Private Function TryReadPassword(ByRef password As String,
                                     ByRef validationMessage As String) As Boolean
        Dim newPassword As String = If(NewPasswordInput.Password, String.Empty).Trim()
        Dim confirmPassword As String = If(ConfirmPasswordInput.Password, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(newPassword) AndAlso
           String.IsNullOrWhiteSpace(confirmPassword) Then
            validationMessage = String.Empty
            password = String.Empty
            Return True
        End If

        If String.IsNullOrWhiteSpace(newPassword) Then
            validationMessage = "New Password is required."
            password = String.Empty
            Return False
        End If

        If String.IsNullOrWhiteSpace(confirmPassword) Then
            validationMessage = "Confirm Password is required."
            password = String.Empty
            Return False
        End If

        If Not String.Equals(newPassword, confirmPassword, StringComparison.Ordinal) Then
            validationMessage = "Passwords do not match."
            password = String.Empty
            Return False
        End If

        If newPassword.Length < 8 Then
            validationMessage = "Passwords must be at least 8 characters."
            password = String.Empty
            Return False
        End If

        validationMessage = String.Empty
        password = newPassword
        Return True
    End Function

    Private Function BuildSaveRequest(teacher As TeacherRecord,
                                      profileValues As TeacherProfileValues,
                                      password As String) As TeacherSaveRequest
        Dim resolvedValues As TeacherProfileValues =
            If(profileValues, New TeacherProfileValues())
        Dim employeeNumber As String = ResolveEmployeeId(teacher, _currentTeacherId)

        Return New TeacherSaveRequest() With {
            .OriginalEmployeeNumber = employeeNumber,
            .EmployeeNumber = employeeNumber,
            .FirstName = resolvedValues.FirstName,
            .MiddleName = resolvedValues.MiddleName,
            .LastName = resolvedValues.LastName,
            .DepartmentText = resolvedValues.Department,
            .PositionTitle = resolvedValues.Position,
            .AdvisorySection = resolvedValues.Advisory,
            .PhotoPath = resolvedValues.PhotoPath,
            .Password = If(password, String.Empty).Trim()
        }
    End Function

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

    Private Function ResolveFailureMessage(result As ServiceResult(Of TeacherRecord)) As String
        If result IsNot Nothing AndAlso
           Not String.IsNullOrWhiteSpace(result.Message) Then
            Return result.Message
        End If

        Return "Unable to save your profile right now."
    End Function

    Private Sub ShowSaveValidationMessage(message As String)
        Dim resolvedMessage As String = If(message, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(resolvedMessage) Then
            resolvedMessage = "Please review the required fields."
        End If

        ShowProfileStatus(resolvedMessage, True)
        MessageBox.Show(resolvedMessage,
                        "Teacher Profile",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information)
    End Sub

    Private Sub ShowProfileStatus(message As String, isError As Boolean)
        Dim normalizedMessage As String = If(message, String.Empty).Trim()

        ProfileTeacherStatusTextBlock.Text = normalizedMessage
        ProfileTeacherStatusTextBlock.Visibility =
            If(String.IsNullOrWhiteSpace(normalizedMessage),
               Visibility.Collapsed,
               Visibility.Visible)
        ProfileTeacherStatusTextBlock.Foreground =
            CType(FindResource(If(isError,
                                  "DashboardDangerBrush",
                                  "DashboardTextMutedBrush")), Brush)
    End Sub

    Private Sub ClearPasswordInputs()
        NewPasswordInput.Password = String.Empty
        ConfirmPasswordInput.Password = String.Empty
    End Sub

    Private Function BuildDisplayNameFromInputs() As String
        Dim parts As New List(Of String)()
        Dim firstName As String = If(FirstNameEditTextBox.Text, String.Empty).Trim()
        Dim middleName As String = If(MiddleNameEditTextBox.Text, String.Empty).Trim()
        Dim lastName As String = If(LastNameEditTextBox.Text, String.Empty).Trim()

        If Not String.IsNullOrWhiteSpace(firstName) Then
            parts.Add(firstName)
        End If

        If Not String.IsNullOrWhiteSpace(middleName) Then
            parts.Add(middleName)
        End If

        If Not String.IsNullOrWhiteSpace(lastName) Then
            parts.Add(lastName)
        End If

        If parts.Count = 0 Then
            Return ResolveDisplayName(_currentTeacherRecord, _currentTeacherName)
        End If

        Return String.Join(" ", parts)
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

    Private Function ResolveHeaderDetails(teacher As TeacherRecord,
                                          fallbackTeacherId As String) As String
        Dim detailParts As New List(Of String)()
        Dim employeeNumber As String = ResolveEmployeeId(teacher, fallbackTeacherId)
        Dim department As String = ResolveDepartment(teacher)
        Dim position As String = ResolvePosition(teacher)

        If Not String.IsNullOrWhiteSpace(employeeNumber) Then
            detailParts.Add(employeeNumber)
        End If

        If Not String.IsNullOrWhiteSpace(department) Then
            detailParts.Add(department)
        End If

        If Not String.IsNullOrWhiteSpace(position) Then
            detailParts.Add(position)
        End If

        If detailParts.Count = 0 Then
            Return "No Teacher ID"
        End If

        Return String.Join(" | ", detailParts)
    End Function

    Private Function ResolveEmployeeId(teacher As TeacherRecord,
                                       fallbackTeacherId As String) As String
        If teacher IsNot Nothing Then
            Dim employeeNumber As String = If(teacher.EmployeeNumber, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(employeeNumber) Then
                Return employeeNumber
            End If
        End If

        Return If(fallbackTeacherId, String.Empty).Trim()
    End Function

    Private Function ResolveFirstName(teacher As TeacherRecord) As String
        If teacher Is Nothing Then
            Return String.Empty
        End If

        Return If(teacher.FirstName, String.Empty).Trim()
    End Function

    Private Function ResolveMiddleName(teacher As TeacherRecord) As String
        If teacher Is Nothing Then
            Return String.Empty
        End If

        Return If(teacher.MiddleName, String.Empty).Trim()
    End Function

    Private Function ResolveLastName(teacher As TeacherRecord,
                                     fallbackName As String) As String
        If teacher IsNot Nothing Then
            Dim lastName As String = If(teacher.LastName, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(lastName) Then
                Return lastName
            End If
        End If

        Dim normalizedFallbackName As String = If(fallbackName, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedFallbackName) Then
            Return String.Empty
        End If

        Dim nameParts As String() = normalizedFallbackName.Split({" "c},
                                                                 StringSplitOptions.RemoveEmptyEntries)
        If nameParts.Length = 0 Then
            Return String.Empty
        End If

        Return nameParts(nameParts.Length - 1)
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

    Private Function ResolveAdvisory(teacher As TeacherRecord) As String
        If teacher Is Nothing Then
            Return String.Empty
        End If

        Return If(teacher.AdvisorySection, String.Empty).Trim()
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

    Private Function ResolveDisplayValue(value As String,
                                         placeholder As String) As String
        Dim normalizedValue As String = If(value, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedValue) Then
            Return placeholder
        End If

        Return normalizedValue
    End Function
End Class
