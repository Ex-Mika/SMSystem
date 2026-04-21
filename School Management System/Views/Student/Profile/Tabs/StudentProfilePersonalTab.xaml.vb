Imports System.Collections.Generic
Imports Microsoft.Win32
Imports System.Windows.Media
Imports School_Management_System.Backend.Models

Class StudentProfilePersonalTab
    Public Class StudentPersonalProfileValues
        Public Property FirstName As String = String.Empty
        Public Property MiddleName As String = String.Empty
        Public Property LastName As String = String.Empty
        Public Property PhotoPath As String = String.Empty
    End Class

    Private ReadOnly _photoAvatarFallbackBrush As Brush
    Private _currentStudent As StudentRecord
    Private _fallbackStudentId As String = String.Empty
    Private _fallbackStudentName As String = String.Empty
    Private _selectedPhotoPath As String = String.Empty
    Private _isEditMode As Boolean

    Public Sub New()
        InitializeComponent()
        _photoAvatarFallbackBrush = StudentPhotoAvatarBorder.Background
        UpdateEditModeState()
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        SetStudentRecord(Nothing, studentId, studentName)
    End Sub

    Public Sub SetStudentRecord(student As StudentRecord,
                                fallbackStudentId As String,
                                fallbackStudentName As String)
        _currentStudent = student
        _fallbackStudentId = If(fallbackStudentId, String.Empty).Trim()
        _fallbackStudentName = If(fallbackStudentName, String.Empty).Trim()
        _selectedPhotoPath = ResolvePhotoPath(student)

        Dim resolvedStudentId As String = ResolveStudentId(student, _fallbackStudentId)
        Dim resolvedFirstName As String = ResolveFirstName(student)
        Dim resolvedMiddleName As String = ResolveMiddleName(student)
        Dim resolvedLastName As String = ResolveLastName(student, _fallbackStudentName)

        StudentIdValueTextBlock.Text = If(String.IsNullOrWhiteSpace(resolvedStudentId),
                                          "No Student ID",
                                          resolvedStudentId)
        FirstNameValueTextBlock.Text = ResolveDisplayValue(resolvedFirstName, "Not set")
        MiddleNameValueTextBlock.Text = ResolveDisplayValue(resolvedMiddleName, "Not set")
        LastNameValueTextBlock.Text = ResolveDisplayValue(resolvedLastName, "Not set")

        FirstNameEditTextBox.Text = resolvedFirstName
        MiddleNameEditTextBox.Text = resolvedMiddleName
        LastNameEditTextBox.Text = resolvedLastName

        UpdatePhotoPreview()
        UpdatePhotoStatus()
        UpdateNotesText()
    End Sub

    Public Sub SetEditMode(isEditMode As Boolean)
        _isEditMode = isEditMode

        If Not _isEditMode Then
            FirstNameEditTextBox.Text = ResolveFirstName(_currentStudent)
            MiddleNameEditTextBox.Text = ResolveMiddleName(_currentStudent)
            LastNameEditTextBox.Text = ResolveLastName(_currentStudent, _fallbackStudentName)
            _selectedPhotoPath = ResolvePhotoPath(_currentStudent)
        End If

        UpdateEditModeState()
        UpdatePhotoPreview()
        UpdatePhotoStatus()
        UpdateNotesText()
    End Sub

    Public Function TryReadProfileValues(ByRef values As StudentPersonalProfileValues,
                                         ByRef validationMessage As String) As Boolean
        Dim resolvedValues As New StudentPersonalProfileValues() With {
            .FirstName = If(FirstNameEditTextBox.Text, String.Empty).Trim(),
            .MiddleName = If(MiddleNameEditTextBox.Text, String.Empty).Trim(),
            .LastName = If(LastNameEditTextBox.Text, String.Empty).Trim(),
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

    Private Sub EditableNameTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        If Not _isEditMode Then
            Return
        End If

        UpdatePhotoPreview()
    End Sub

    Private Sub BrowseStudentPhotoButton_Click(sender As Object, e As RoutedEventArgs)
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
        UpdatePhotoPreview()
        UpdatePhotoStatus()
    End Sub

    Private Sub ClearStudentPhotoButton_Click(sender As Object, e As RoutedEventArgs)
        _selectedPhotoPath = String.Empty
        UpdatePhotoPreview()
        UpdatePhotoStatus()
    End Sub

    Private Sub UpdateEditModeState()
        Dim editVisibility As Visibility = If(_isEditMode,
                                              Visibility.Visible,
                                              Visibility.Collapsed)
        Dim viewVisibility As Visibility = If(_isEditMode,
                                              Visibility.Collapsed,
                                              Visibility.Visible)

        FirstNameValueBorder.Visibility = viewVisibility
        MiddleNameValueBorder.Visibility = viewVisibility
        LastNameValueBorder.Visibility = viewVisibility

        FirstNameEditTextBox.Visibility = editVisibility
        MiddleNameEditTextBox.Visibility = editVisibility
        LastNameEditTextBox.Visibility = editVisibility

        BrowseStudentPhotoButton.Visibility = editVisibility
        ClearStudentPhotoButton.Visibility = editVisibility
    End Sub

    Private Sub UpdatePhotoPreview()
        Dim displayName As String = BuildDisplayName()

        StudentPhotoInitialTextBlock.Text = BuildInitial(displayName)
        DashboardProfileImageHelper.ApplyProfilePhoto(StudentPhotoAvatarBorder,
                                                      StudentPhotoInitialTextBlock,
                                                      _selectedPhotoPath,
                                                      _photoAvatarFallbackBrush)
    End Sub

    Private Sub UpdatePhotoStatus()
        If String.IsNullOrWhiteSpace(_selectedPhotoPath) Then
            StudentPhotoStatusTextBlock.Text =
                If(_isEditMode,
                   "No photo selected. Browse to add one or leave it blank to remove it.",
                   "No profile photo on file.")
            Return
        End If

        If _isEditMode AndAlso
           Not String.Equals(_selectedPhotoPath,
                             ResolvePhotoPath(_currentStudent),
                             StringComparison.OrdinalIgnoreCase) Then
            StudentPhotoStatusTextBlock.Text = "Selected photo is ready to save."
            Return
        End If

        StudentPhotoStatusTextBlock.Text = "Current profile photo from your student record."
    End Sub

    Private Sub UpdateNotesText()
        ProfilePersonalNotesTextBlock.Text =
            If(_isEditMode,
               "Only your name, photo, and password can be updated here. " &
               "Course, year level, and section stay managed by administration.",
               "Use Edit Profile to update your name and photo. " &
               "Academic placement is managed by the administration team.")
    End Sub

    Private Function ResolveStudentId(student As StudentRecord,
                                      fallbackStudentId As String) As String
        If student IsNot Nothing Then
            Dim studentId As String = If(student.StudentNumber, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(studentId) Then
                Return studentId
            End If
        End If

        Return If(fallbackStudentId, String.Empty).Trim()
    End Function

    Private Function ResolveFirstName(student As StudentRecord) As String
        If student Is Nothing Then
            Return String.Empty
        End If

        Return If(student.FirstName, String.Empty).Trim()
    End Function

    Private Function ResolveMiddleName(student As StudentRecord) As String
        If student Is Nothing Then
            Return String.Empty
        End If

        Return If(student.MiddleName, String.Empty).Trim()
    End Function

    Private Function ResolveLastName(student As StudentRecord,
                                     fallbackStudentName As String) As String
        If student IsNot Nothing Then
            Dim lastName As String = If(student.LastName, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(lastName) Then
                Return lastName
            End If
        End If

        Dim fallbackName As String = If(fallbackStudentName, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(fallbackName) Then
            Return String.Empty
        End If

        Dim nameParts As String() = fallbackName.Split({" "c},
                                                       StringSplitOptions.RemoveEmptyEntries)
        If nameParts.Length = 0 Then
            Return String.Empty
        End If

        Return nameParts(nameParts.Length - 1)
    End Function

    Private Function ResolvePhotoPath(student As StudentRecord) As String
        If student Is Nothing Then
            Return String.Empty
        End If

        Return If(student.PhotoPath, String.Empty).Trim()
    End Function

    Private Function BuildDisplayName() As String
        Dim firstName As String = If(FirstNameEditTextBox.Text, String.Empty).Trim()
        Dim middleName As String = If(MiddleNameEditTextBox.Text, String.Empty).Trim()
        Dim lastName As String = If(LastNameEditTextBox.Text, String.Empty).Trim()
        Dim parts As New List(Of String)()

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
            Return If(String.IsNullOrWhiteSpace(_fallbackStudentName),
                      "Student",
                      _fallbackStudentName)
        End If

        Return String.Join(" ", parts)
    End Function

    Private Function BuildInitial(displayName As String) As String
        Dim normalizedName As String = If(displayName, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedName) Then
            Return "S"
        End If

        Return normalizedName.Substring(0, 1).ToUpperInvariant()
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
