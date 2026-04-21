Imports System.IO
Imports System.Text
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class TeacherManagementService
        Public Const DefaultTeacherPassword As String = "Admin@123"

        Private ReadOnly _teacherRepository As TeacherRepository
        Private ReadOnly _profileImageStorageService As ProfileImageStorageService

        Public Sub New()
            Me.New(New TeacherRepository(), New ProfileImageStorageService())
        End Sub

        Public Sub New(teacherRepository As TeacherRepository)
            Me.New(teacherRepository, New ProfileImageStorageService())
        End Sub

        Public Sub New(teacherRepository As TeacherRepository,
                       profileImageStorageService As ProfileImageStorageService)
            _teacherRepository = teacherRepository
            _profileImageStorageService =
                If(profileImageStorageService, New ProfileImageStorageService())
        End Sub

        Public Function GetTeachers() As ServiceResult(Of List(Of TeacherRecord))
            Try
                Dim teachers As List(Of TeacherRecord) = _teacherRepository.GetAll()
                Return ServiceResult(Of List(Of TeacherRecord)).Success(teachers)
            Catch ex As MySqlException
                Return ServiceResult(Of List(Of TeacherRecord)).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            End Try
        End Function

        Public Function GetTeacherByEmployeeNumber(employeeNumber As String) As ServiceResult(Of TeacherRecord)
            If String.IsNullOrWhiteSpace(employeeNumber) Then
                Return ServiceResult(Of TeacherRecord).Failure("Teacher ID is required.")
            End If

            Try
                Dim teacher As TeacherRecord =
                    _teacherRepository.GetByEmployeeNumber(employeeNumber.Trim())
                If teacher Is Nothing Then
                    Return ServiceResult(Of TeacherRecord).Failure(
                        "The selected teacher no longer exists.")
                End If

                Return ServiceResult(Of TeacherRecord).Success(teacher)
            Catch ex As MySqlException
                Return ServiceResult(Of TeacherRecord).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            Catch ex As Exception
                Return ServiceResult(Of TeacherRecord).Failure(
                    "Unable to load teacher records." &
                    Environment.NewLine &
                    ex.Message)
            End Try
        End Function

        Public Function CreateTeacher(request As TeacherSaveRequest) As ServiceResult(Of TeacherRecord)
            Dim validationMessage As String = ValidateRequest(request, False)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of TeacherRecord).Failure(validationMessage)
            End If

            Dim preparedPhotoPath As String = String.Empty

            Try
                If _teacherRepository.GetByEmployeeNumber(request.EmployeeNumber) IsNot Nothing Then
                    Return ServiceResult(Of TeacherRecord).Failure("Teacher ID already exists.")
                End If

                preparedPhotoPath =
                    _profileImageStorageService.StoreProfileImage(
                        ProfileImageStorageService.ProfileImageOwnerType.Teacher,
                        request.EmployeeNumber,
                        request.PhotoPath)
                request.PhotoPath = preparedPhotoPath

                Dim passwordToUse As String = GetPasswordForCreate(request.Password)
                Dim createdRecord As TeacherRecord =
                    _teacherRepository.Create(request,
                                              Database.DatabaseModule.HashPassword(passwordToUse),
                                              BuildGeneratedEmail(request.EmployeeNumber))

                Dim message As String = "Teacher created."
                If String.IsNullOrWhiteSpace(request.Password) Then
                    message = "Teacher created. Temporary password: " & DefaultTeacherPassword
                End If

                Return ServiceResult(Of TeacherRecord).Success(createdRecord, message)
            Catch ex As MySqlException
                CleanupPreparedProfileImage(preparedPhotoPath, String.Empty)
                Return ServiceResult(Of TeacherRecord).Failure(
                    BuildDatabaseErrorMessage("create", ex))
            Catch ex As IOException
                CleanupPreparedProfileImage(preparedPhotoPath, String.Empty)
                Return ServiceResult(Of TeacherRecord).Failure(
                    BuildProfileImageErrorMessage("save", ex))
            End Try
        End Function

        Public Function UpdateTeacher(request As TeacherSaveRequest) As ServiceResult(Of TeacherRecord)
            Dim validationMessage As String = ValidateRequest(request, True)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of TeacherRecord).Failure(validationMessage)
            End If

            Dim existingRecord As TeacherRecord = Nothing
            Dim preparedPhotoPath As String = String.Empty

            Try
                existingRecord =
                    _teacherRepository.GetByEmployeeNumber(request.OriginalEmployeeNumber)
                If existingRecord Is Nothing Then
                    Return ServiceResult(Of TeacherRecord).Failure(
                        "The selected teacher no longer exists.")
                End If

                Dim duplicateTeacher As TeacherRecord =
                    _teacherRepository.GetByEmployeeNumber(request.EmployeeNumber)
                If duplicateTeacher IsNot Nothing AndAlso
                   duplicateTeacher.TeacherRecordId <> existingRecord.TeacherRecordId Then
                    Return ServiceResult(Of TeacherRecord).Failure("Teacher ID already exists.")
                End If

                Dim shouldUpdatePassword As Boolean = Not String.IsNullOrWhiteSpace(request.Password)
                Dim passwordHash As String = String.Empty

                If shouldUpdatePassword Then
                    passwordHash = Database.DatabaseModule.HashPassword(request.Password.Trim())
                End If

                preparedPhotoPath =
                    _profileImageStorageService.StoreProfileImage(
                        ProfileImageStorageService.ProfileImageOwnerType.Teacher,
                        request.EmployeeNumber,
                        request.PhotoPath,
                        existingRecord.PhotoPath)
                request.PhotoPath = preparedPhotoPath

                Dim updatedRecord As TeacherRecord =
                    _teacherRepository.Update(existingRecord,
                                              request,
                                              ResolveEmailAddress(existingRecord, request.EmployeeNumber),
                                              shouldUpdatePassword,
                                              passwordHash)

                CleanupObsoleteManagedImage(existingRecord.PhotoPath, updatedRecord.PhotoPath)
                Return ServiceResult(Of TeacherRecord).Success(updatedRecord, "Teacher updated.")
            Catch ex As MySqlException
                CleanupPreparedProfileImage(preparedPhotoPath,
                                            If(existingRecord Is Nothing,
                                               String.Empty,
                                               existingRecord.PhotoPath))
                Return ServiceResult(Of TeacherRecord).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            Catch ex As IOException
                CleanupPreparedProfileImage(preparedPhotoPath,
                                            If(existingRecord Is Nothing,
                                               String.Empty,
                                               existingRecord.PhotoPath))
                Return ServiceResult(Of TeacherRecord).Failure(
                    BuildProfileImageErrorMessage("save", ex))
            End Try
        End Function

        Public Function DeleteTeacher(employeeNumber As String) As ServiceResult(Of Boolean)
            If String.IsNullOrWhiteSpace(employeeNumber) Then
                Return ServiceResult(Of Boolean).Failure("Teacher ID is required.")
            End If

            Try
                Dim existingRecord As TeacherRecord =
                    _teacherRepository.GetByEmployeeNumber(employeeNumber.Trim())
                If existingRecord Is Nothing Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected teacher no longer exists.")
                End If

                Dim deleted As Boolean =
                    _teacherRepository.DeleteByEmployeeNumber(employeeNumber.Trim())
                If Not deleted Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected teacher no longer exists.")
                End If

                CleanupObsoleteManagedImage(existingRecord.PhotoPath, String.Empty)
                Return ServiceResult(Of Boolean).Success(True, "Teacher deleted.")
            Catch ex As MySqlException
                Return ServiceResult(Of Boolean).Failure(
                    BuildDatabaseErrorMessage("delete", ex))
            End Try
        End Function

        Private Function ValidateRequest(request As TeacherSaveRequest,
                                         requireOriginalEmployeeNumber As Boolean) As String
            If request Is Nothing Then
                Return "Teacher data is required."
            End If

            If requireOriginalEmployeeNumber AndAlso
               String.IsNullOrWhiteSpace(request.OriginalEmployeeNumber) Then
                Return "The original teacher record is required."
            End If

            If String.IsNullOrWhiteSpace(request.EmployeeNumber) Then
                Return "Teacher ID is required."
            End If

            If String.IsNullOrWhiteSpace(request.FirstName) Then
                Return "First Name is required."
            End If

            If String.IsNullOrWhiteSpace(request.LastName) Then
                Return "Last Name is required."
            End If

            If Not String.IsNullOrWhiteSpace(request.Password) AndAlso request.Password.Trim().Length < 8 Then
                Return "Passwords must be at least 8 characters."
            End If

            Return String.Empty
        End Function

        Private Function GetPasswordForCreate(password As String) As String
            If String.IsNullOrWhiteSpace(password) Then
                Return DefaultTeacherPassword
            End If

            Return password.Trim()
        End Function

        Private Function ResolveEmailAddress(existingRecord As TeacherRecord,
                                             employeeNumber As String) As String
            If existingRecord Is Nothing Then
                Return BuildGeneratedEmail(employeeNumber)
            End If

            Dim originalGeneratedEmail As String = BuildGeneratedEmail(existingRecord.EmployeeNumber)
            If String.Equals(existingRecord.Email,
                             originalGeneratedEmail,
                             StringComparison.OrdinalIgnoreCase) Then
                Return BuildGeneratedEmail(employeeNumber)
            End If

            Return existingRecord.Email
        End Function

        Private Function BuildGeneratedEmail(employeeNumber As String) As String
            Dim builder As New StringBuilder()

            For Each currentCharacter As Char In If(employeeNumber, String.Empty).Trim().ToLowerInvariant()
                If Char.IsLetterOrDigit(currentCharacter) Then
                    builder.Append(currentCharacter)
                ElseIf currentCharacter = "."c OrElse
                       currentCharacter = "_"c OrElse
                       currentCharacter = "-"c Then
                    builder.Append(currentCharacter)
                Else
                    builder.Append("-"c)
                End If
            Next

            Dim localPart As String = builder.ToString().Trim("-"c)
            If String.IsNullOrWhiteSpace(localPart) Then
                localPart = "teacher"
            End If

            Return "teacher." & localPart & "@prmsu.local"
        End Function

        Private Function BuildDatabaseErrorMessage(operationName As String,
                                                   ex As MySqlException) As String
            If ex Is Nothing Then
                Return "Unable to " & operationName & " teacher records."
            End If

            If ex.Number = 1062 Then
                Return "A duplicate teacher ID or generated login already exists."
            End If

            Return "Unable to " & operationName & " teacher records." &
                Environment.NewLine &
                ex.Message
        End Function

        Private Function BuildProfileImageErrorMessage(operationName As String,
                                                       ex As IOException) As String
            If TypeOf ex Is FileNotFoundException Then
                Return "The selected teacher photo file could not be found."
            End If

            Return "Unable to " & operationName & " the selected teacher photo locally." &
                Environment.NewLine &
                ex.Message
        End Function

        Private Sub CleanupPreparedProfileImage(preparedPhotoPath As String,
                                                preservedPhotoPath As String)
            Dim normalizedPreparedPhotoPath As String =
                If(preparedPhotoPath, String.Empty).Trim()
            Dim normalizedPreservedPhotoPath As String =
                If(preservedPhotoPath, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedPreparedPhotoPath) Then
                Return
            End If

            If String.Equals(normalizedPreparedPhotoPath,
                             normalizedPreservedPhotoPath,
                             StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            _profileImageStorageService.DeleteManagedImage(normalizedPreparedPhotoPath)
        End Sub

        Private Sub CleanupObsoleteManagedImage(previousPhotoPath As String,
                                                currentPhotoPath As String)
            Dim normalizedPreviousPhotoPath As String =
                If(previousPhotoPath, String.Empty).Trim()
            Dim normalizedCurrentPhotoPath As String =
                If(currentPhotoPath, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedPreviousPhotoPath) Then
                Return
            End If

            If String.Equals(normalizedPreviousPhotoPath,
                             normalizedCurrentPhotoPath,
                             StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            _profileImageStorageService.DeleteManagedImage(normalizedPreviousPhotoPath)
        End Sub
    End Class
End Namespace
