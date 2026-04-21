Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Media
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class StudentProfileView
    Public Class StudentProfileUpdatedEventArgs
        Inherits EventArgs

        Public Sub New(student As StudentRecord)
            Me.Student = student
        End Sub

        Public ReadOnly Property Student As StudentRecord
    End Class

    Private Enum StudentProfileSection
        Personal
        Academic
        Account
    End Enum

    Public Event ProfileUpdated As EventHandler(Of StudentProfileUpdatedEventArgs)

    Private ReadOnly _studentManagementService As New StudentManagementService()
    Private ReadOnly _profileHeaderAvatarFallbackBrush As Brush
    Private _currentStudentRecord As StudentRecord
    Private _currentStudentId As String = String.Empty
    Private _currentStudentName As String = String.Empty
    Private _currentStudentPhotoPath As String = String.Empty
    Private _isEditMode As Boolean

    Public Sub New()
        InitializeComponent()
        _profileHeaderAvatarFallbackBrush = ProfileHeaderAvatarBorder.Background
        AddHandler AccountTabView.EditRequested, AddressOf AccountTabView_EditRequested
        SetActiveSection(StudentProfileSection.Personal)
        UpdateEditModeState()
    End Sub

    Public Sub SetStudentContext(studentId As String,
                                 studentName As String,
                                 Optional studentPhotoPath As String = "")
        _currentStudentId = If(studentId, String.Empty).Trim()
        _currentStudentName = If(studentName, String.Empty).Trim()
        _currentStudentPhotoPath = If(studentPhotoPath, String.Empty).Trim()
        _currentStudentRecord = LoadStudentRecord(_currentStudentId)

        ApplyCurrentStudentRecord()
        SetEditMode(False)
        ShowProfileStatus(String.Empty, False)
    End Sub

    Private Sub PersonalSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentProfileSection.Personal)
    End Sub

    Private Sub AcademicSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentProfileSection.Academic)
    End Sub

    Private Sub AccountSectionButton_Click(sender As Object, e As RoutedEventArgs)
        SetActiveSection(StudentProfileSection.Account)
    End Sub

    Private Sub EditProfileButton_Click(sender As Object, e As RoutedEventArgs)
        If _currentStudentRecord Is Nothing Then
            MessageBox.Show("Your profile could not be loaded right now.",
                            "Student Profile",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        SetEditMode(True)
        ShowProfileStatus("Edit your personal details, photo, or password. " &
                          "Academic placement stays admin-managed.",
                          False)
    End Sub

    Private Sub SaveProfileButton_Click(sender As Object, e As RoutedEventArgs)
        If _currentStudentRecord Is Nothing Then
            MessageBox.Show("Your profile could not be loaded right now.",
                            "Student Profile",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        Dim validationMessage As String = String.Empty
        Dim personalValues As StudentProfilePersonalTab.StudentPersonalProfileValues = Nothing
        If Not PersonalTabView.TryReadProfileValues(personalValues, validationMessage) Then
            SetActiveSection(StudentProfileSection.Personal)
            ShowSaveValidationMessage(validationMessage)
            Return
        End If

        Dim updatedPassword As String = String.Empty
        If Not AccountTabView.TryReadPassword(updatedPassword, validationMessage) Then
            SetActiveSection(StudentProfileSection.Account)
            ShowSaveValidationMessage(validationMessage)
            Return
        End If

        Dim request As StudentSaveRequest =
            BuildSaveRequest(_currentStudentRecord,
                             personalValues,
                             updatedPassword)
        Dim result = _studentManagementService.UpdateStudent(request)
        If result Is Nothing OrElse
           Not result.IsSuccess OrElse
           result.Data Is Nothing Then
            Dim failureMessage As String = ResolveFailureMessage(result)
            ShowProfileStatus(failureMessage, True)
            MessageBox.Show(failureMessage,
                            "Student Profile",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information)
            Return
        End If

        _currentStudentRecord = result.Data
        _currentStudentId = If(result.Data.StudentNumber, String.Empty).Trim()
        _currentStudentName = ResolveDisplayName(result.Data, _currentStudentName)
        _currentStudentPhotoPath = ResolveStudentPhotoPath(result.Data,
                                                          _currentStudentPhotoPath)

        ApplyCurrentStudentRecord()
        SetEditMode(False)
        ShowProfileStatus("Profile updated.", False)

        RaiseEvent ProfileUpdated(Me,
                                  New StudentProfileUpdatedEventArgs(_currentStudentRecord))
    End Sub

    Private Sub CancelEditProfileButton_Click(sender As Object, e As RoutedEventArgs)
        SetEditMode(False)
        ShowProfileStatus(String.Empty, False)
    End Sub

    Private Sub AccountTabView_EditRequested(sender As Object, e As EventArgs)
        SetActiveSection(StudentProfileSection.Account)
        SetEditMode(True)
        ShowProfileStatus("Enter a new password in the Account tab, then save your changes.",
                          False)
    End Sub

    Private Sub SetActiveSection(section As StudentProfileSection)
        PersonalTabView.Visibility = If(section = StudentProfileSection.Personal,
                                        Visibility.Visible,
                                        Visibility.Collapsed)
        AcademicTabView.Visibility = If(section = StudentProfileSection.Academic,
                                        Visibility.Visible,
                                        Visibility.Collapsed)
        AccountTabView.Visibility = If(section = StudentProfileSection.Account,
                                       Visibility.Visible,
                                       Visibility.Collapsed)

        ApplySectionButtonState(PersonalSectionButton, section = StudentProfileSection.Personal)
        ApplySectionButtonState(AcademicSectionButton, section = StudentProfileSection.Academic)
        ApplySectionButtonState(AccountSectionButton, section = StudentProfileSection.Account)
    End Sub

    Private Sub ApplySectionButtonState(sectionButton As Button, isSelected As Boolean)
        sectionButton.Style = CType(FindResource(If(isSelected,
                                                    "DashboardProfileSegmentSelectedButtonStyle",
                                                    "DashboardProfileSegmentButtonStyle")), Style)
    End Sub

    Private Sub SetEditMode(isEditMode As Boolean)
        _isEditMode = isEditMode AndAlso _currentStudentRecord IsNot Nothing
        PersonalTabView.SetEditMode(_isEditMode)
        AccountTabView.SetEditMode(_isEditMode)
        UpdateEditModeState()
    End Sub

    Private Sub UpdateEditModeState()
        Dim canEdit As Boolean = _currentStudentRecord IsNot Nothing

        EditProfileButton.IsEnabled = canEdit
        EditProfileButton.Visibility =
            If(_isEditMode, Visibility.Collapsed, Visibility.Visible)
        EditProfileActionsPanel.Visibility =
            If(_isEditMode, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub ApplyCurrentStudentRecord()
        Dim displayName As String = ResolveDisplayName(_currentStudentRecord, _currentStudentName)
        Dim headerDetails As String = ResolveHeaderDetails(_currentStudentRecord,
                                                           _currentStudentId)
        Dim photoPath As String = ResolveStudentPhotoPath(_currentStudentRecord,
                                                          _currentStudentPhotoPath)

        ProfileHeaderNameTextBlock.Text = displayName
        ProfileHeaderIdTextBlock.Text = headerDetails
        ProfileHeaderInitialTextBlock.Text = GetProfileInitial(displayName)
        DashboardProfileImageHelper.ApplyProfilePhoto(ProfileHeaderAvatarBorder,
                                                      ProfileHeaderInitialTextBlock,
                                                      photoPath,
                                                      _profileHeaderAvatarFallbackBrush)
        PersonalTabView.SetStudentRecord(_currentStudentRecord, _currentStudentId, displayName)
        PersonalTabView.SetEditMode(_isEditMode)
        AcademicTabView.SetStudentRecord(_currentStudentRecord)
        AccountTabView.SetStudentRecord(_currentStudentRecord)
        AccountTabView.SetEditMode(_isEditMode)
        UpdateEditModeState()
    End Sub

    Private Function BuildSaveRequest(student As StudentRecord,
                                      personalValues As StudentProfilePersonalTab.StudentPersonalProfileValues,
                                      password As String) As StudentSaveRequest
        Dim resolvedPersonalValues As StudentProfilePersonalTab.StudentPersonalProfileValues =
            If(personalValues, New StudentProfilePersonalTab.StudentPersonalProfileValues())

        Return New StudentSaveRequest() With {
            .OriginalStudentNumber = If(student.StudentNumber, String.Empty).Trim(),
            .StudentNumber = If(student.StudentNumber, String.Empty).Trim(),
            .FirstName = If(resolvedPersonalValues.FirstName, String.Empty).Trim(),
            .MiddleName = If(resolvedPersonalValues.MiddleName, String.Empty).Trim(),
            .LastName = If(resolvedPersonalValues.LastName, String.Empty).Trim(),
            .YearLevel = student.YearLevel,
            .CourseText = ResolveCourseText(student),
            .SectionName = If(student.SectionName, String.Empty).Trim(),
            .PhotoPath = If(resolvedPersonalValues.PhotoPath, String.Empty).Trim(),
            .Password = If(password, String.Empty).Trim()
        }
    End Function

    Private Function ResolveCourseText(student As StudentRecord) As String
        If student Is Nothing Then
            Return String.Empty
        End If

        Dim courseCode As String = If(student.CourseCode, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(courseCode) Then
            Return courseCode
        End If

        Return If(student.CourseName, String.Empty).Trim()
    End Function

    Private Function ResolveFailureMessage(result As ServiceResult(Of StudentRecord)) As String
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
                        "Student Profile",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information)
    End Sub

    Private Sub ShowProfileStatus(message As String, isError As Boolean)
        Dim normalizedMessage As String = If(message, String.Empty).Trim()

        ProfileHeaderStatusTextBlock.Text = normalizedMessage
        ProfileHeaderStatusTextBlock.Visibility =
            If(String.IsNullOrWhiteSpace(normalizedMessage),
               Visibility.Collapsed,
               Visibility.Visible)
        ProfileHeaderStatusTextBlock.Foreground =
            CType(FindResource(If(isError,
                                  "DashboardDangerBrush",
                                  "DashboardTextMutedBrush")), Brush)
    End Sub

    Private Function GetProfileInitial(fullName As String) As String
        Dim normalizedName As String = If(fullName, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedName) Then
            Return "S"
        End If

        Return normalizedName.Substring(0, 1).ToUpperInvariant()
    End Function

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

    Private Function ResolveHeaderDetails(student As StudentRecord,
                                          fallbackStudentId As String) As String
        Dim detailParts As New List(Of String)()
        Dim resolvedStudentId As String = If(fallbackStudentId, String.Empty).Trim()

        If student IsNot Nothing Then
            Dim recordStudentId As String = If(student.StudentNumber, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(recordStudentId) Then
                resolvedStudentId = recordStudentId
            End If
        End If

        If Not String.IsNullOrWhiteSpace(resolvedStudentId) Then
            detailParts.Add(resolvedStudentId)
        End If

        If student IsNot Nothing Then
            Dim courseName As String = If(student.CourseDisplayName, String.Empty).Trim()
            Dim yearLabel As String = If(student.YearLevelLabel, String.Empty).Trim()
            Dim sectionLabel As String = BuildSectionDisplayLabel(student)

            If Not String.IsNullOrWhiteSpace(courseName) Then
                detailParts.Add(courseName)
            End If

            If Not String.IsNullOrWhiteSpace(yearLabel) Then
                detailParts.Add(yearLabel)
            End If

            If Not String.IsNullOrWhiteSpace(sectionLabel) Then
                detailParts.Add(sectionLabel)
            End If
        End If

        If detailParts.Count = 0 Then
            Return "No Student ID"
        End If

        Return String.Join(" | ", detailParts)
    End Function

    Private Function BuildSectionDisplayLabel(student As StudentRecord) As String
        Return StudentScheduleHelper.BuildStudentSectionValue(student, String.Empty)
    End Function
End Class
