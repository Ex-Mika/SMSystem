Imports System.IO
Imports System.Text
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class StudentManagementService
        Public Const DefaultStudentPassword As String = "Admin@123"

        Private ReadOnly _studentRepository As StudentRepository
        Private ReadOnly _profileImageStorageService As ProfileImageStorageService

        Public Sub New()
            Me.New(New StudentRepository(), New ProfileImageStorageService())
        End Sub

        Public Sub New(studentRepository As StudentRepository)
            Me.New(studentRepository, New ProfileImageStorageService())
        End Sub

        Public Sub New(studentRepository As StudentRepository,
                       profileImageStorageService As ProfileImageStorageService)
            _studentRepository = studentRepository
            _profileImageStorageService =
                If(profileImageStorageService, New ProfileImageStorageService())
        End Sub

        Public Function GetStudents() As ServiceResult(Of List(Of StudentRecord))
            Try
                Dim students As List(Of StudentRecord) = _studentRepository.GetAll()
                Return ServiceResult(Of List(Of StudentRecord)).Success(students)
            Catch ex As MySqlException
                Return ServiceResult(Of List(Of StudentRecord)).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            End Try
        End Function

        Public Function GetStudentByStudentNumber(studentNumber As String) As ServiceResult(Of StudentRecord)
            If String.IsNullOrWhiteSpace(studentNumber) Then
                Return ServiceResult(Of StudentRecord).Failure("Student ID is required.")
            End If

            Try
                Dim student As StudentRecord = _studentRepository.GetByStudentNumber(studentNumber.Trim())
                If student Is Nothing Then
                    Return ServiceResult(Of StudentRecord).Failure(
                        "The selected student no longer exists.")
                End If

                Return ServiceResult(Of StudentRecord).Success(student)
            Catch ex As MySqlException
                Return ServiceResult(Of StudentRecord).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            Catch ex As Exception
                Return ServiceResult(Of StudentRecord).Failure(
                    "Unable to load student records." &
                    Environment.NewLine &
                    ex.Message)
            End Try
        End Function

        Public Function CreateStudent(request As StudentSaveRequest) As ServiceResult(Of StudentRecord)
            Dim validationMessage As String = ValidateRequest(request, False)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of StudentRecord).Failure(validationMessage)
            End If

            Dim preparedPhotoPath As String = String.Empty

            Try
                If _studentRepository.GetByStudentNumber(request.StudentNumber) IsNot Nothing Then
                    Return ServiceResult(Of StudentRecord).Failure("Student ID already exists.")
                End If

                preparedPhotoPath =
                    _profileImageStorageService.StoreProfileImage(
                        ProfileImageStorageService.ProfileImageOwnerType.Student,
                        request.StudentNumber,
                        request.PhotoPath)
                request.PhotoPath = preparedPhotoPath

                Dim passwordToUse As String = GetPasswordForCreate(request.Password)
                Dim createdRecord As StudentRecord =
                    _studentRepository.Create(request,
                                              Database.DatabaseModule.HashPassword(passwordToUse),
                                              BuildGeneratedEmail(request.StudentNumber))

                Dim message As String = "Student created."
                If String.IsNullOrWhiteSpace(request.Password) Then
                    message = "Student created. Temporary password: " & DefaultStudentPassword
                End If

                Return ServiceResult(Of StudentRecord).Success(createdRecord, message)
            Catch ex As InvalidOperationException
                CleanupPreparedProfileImage(preparedPhotoPath, String.Empty)
                Return ServiceResult(Of StudentRecord).Failure(ex.Message)
            Catch ex As MySqlException
                CleanupPreparedProfileImage(preparedPhotoPath, String.Empty)
                Return ServiceResult(Of StudentRecord).Failure(
                    BuildDatabaseErrorMessage("create", ex))
            Catch ex As IOException
                CleanupPreparedProfileImage(preparedPhotoPath, String.Empty)
                Return ServiceResult(Of StudentRecord).Failure(
                    BuildProfileImageErrorMessage("save", ex))
            End Try
        End Function

        Public Function UpdateStudent(request As StudentSaveRequest) As ServiceResult(Of StudentRecord)
            Dim validationMessage As String = ValidateRequest(request, True)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of StudentRecord).Failure(validationMessage)
            End If

            Dim existingRecord As StudentRecord = Nothing
            Dim preparedPhotoPath As String = String.Empty

            Try
                existingRecord =
                    _studentRepository.GetByStudentNumber(request.OriginalStudentNumber)
                If existingRecord Is Nothing Then
                    Return ServiceResult(Of StudentRecord).Failure(
                        "The selected student no longer exists.")
                End If

                Dim duplicateStudent As StudentRecord =
                    _studentRepository.GetByStudentNumber(request.StudentNumber)
                If duplicateStudent IsNot Nothing AndAlso
                   duplicateStudent.StudentRecordId <> existingRecord.StudentRecordId Then
                    Return ServiceResult(Of StudentRecord).Failure("Student ID already exists.")
                End If

                Dim shouldUpdatePassword As Boolean = Not String.IsNullOrWhiteSpace(request.Password)
                Dim passwordHash As String = String.Empty

                If shouldUpdatePassword Then
                    passwordHash = Database.DatabaseModule.HashPassword(request.Password.Trim())
                End If

                preparedPhotoPath =
                    _profileImageStorageService.StoreProfileImage(
                        ProfileImageStorageService.ProfileImageOwnerType.Student,
                        request.StudentNumber,
                        request.PhotoPath,
                        existingRecord.PhotoPath)
                request.PhotoPath = preparedPhotoPath

                Dim updatedRecord As StudentRecord =
                    _studentRepository.Update(existingRecord,
                                              request,
                                              ResolveEmailAddress(existingRecord, request.StudentNumber),
                                              shouldUpdatePassword,
                                              passwordHash)

                CleanupObsoleteManagedImage(existingRecord.PhotoPath, updatedRecord.PhotoPath)
                Return ServiceResult(Of StudentRecord).Success(updatedRecord, "Student updated.")
            Catch ex As InvalidOperationException
                CleanupPreparedProfileImage(preparedPhotoPath,
                                            If(existingRecord Is Nothing,
                                               String.Empty,
                                               existingRecord.PhotoPath))
                Return ServiceResult(Of StudentRecord).Failure(ex.Message)
            Catch ex As MySqlException
                CleanupPreparedProfileImage(preparedPhotoPath,
                                            If(existingRecord Is Nothing,
                                               String.Empty,
                                               existingRecord.PhotoPath))
                Return ServiceResult(Of StudentRecord).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            Catch ex As IOException
                CleanupPreparedProfileImage(preparedPhotoPath,
                                            If(existingRecord Is Nothing,
                                               String.Empty,
                                               existingRecord.PhotoPath))
                Return ServiceResult(Of StudentRecord).Failure(
                    BuildProfileImageErrorMessage("save", ex))
            End Try
        End Function

        Public Function DeleteStudent(studentNumber As String) As ServiceResult(Of Boolean)
            If String.IsNullOrWhiteSpace(studentNumber) Then
                Return ServiceResult(Of Boolean).Failure("Student ID is required.")
            End If

            Try
                Dim existingRecord As StudentRecord =
                    _studentRepository.GetByStudentNumber(studentNumber.Trim())
                If existingRecord Is Nothing Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected student no longer exists.")
                End If

                Dim deleted As Boolean =
                    _studentRepository.DeleteByStudentNumber(studentNumber.Trim())
                If Not deleted Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected student no longer exists.")
                End If

                CleanupObsoleteManagedImage(existingRecord.PhotoPath, String.Empty)
                Return ServiceResult(Of Boolean).Success(True, "Student deleted.")
            Catch ex As MySqlException
                Return ServiceResult(Of Boolean).Failure(
                    BuildDatabaseErrorMessage("delete", ex))
            End Try
        End Function

        Public Function UpdateStudentSection(studentNumber As String,
                                             sectionName As String) As ServiceResult(Of StudentRecord)
            Dim normalizedStudentNumber As String = If(studentNumber, String.Empty).Trim()
            Dim normalizedSectionName As String = If(sectionName, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedStudentNumber) Then
                Return ServiceResult(Of StudentRecord).Failure("Student ID is required.")
            End If

            If String.IsNullOrWhiteSpace(normalizedSectionName) Then
                Return ServiceResult(Of StudentRecord).Failure("Section is required.")
            End If

            Try
                Dim existingRecord As StudentRecord =
                    _studentRepository.GetByStudentNumber(normalizedStudentNumber)
                If existingRecord Is Nothing Then
                    Return ServiceResult(Of StudentRecord).Failure(
                        "The selected student no longer exists.")
                End If

                Dim currentSection As String = If(existingRecord.SectionName, String.Empty).Trim()
                If String.Equals(currentSection,
                                 normalizedSectionName,
                                 StringComparison.OrdinalIgnoreCase) Then
                    Return ServiceResult(Of StudentRecord).Success(existingRecord,
                                                                   "Section updated.")
                End If

                If Not _studentRepository.UpdateSection(existingRecord.StudentRecordId,
                                                        normalizedSectionName) Then
                    Return ServiceResult(Of StudentRecord).Failure(
                        "Unable to update the selected section right now.")
                End If

                Dim updatedRecord As StudentRecord =
                    _studentRepository.GetByStudentNumber(normalizedStudentNumber)
                If updatedRecord Is Nothing Then
                    Return ServiceResult(Of StudentRecord).Failure(
                        "The selected student no longer exists.")
                End If

                Return ServiceResult(Of StudentRecord).Success(updatedRecord, "Section updated.")
            Catch ex As MySqlException
                Return ServiceResult(Of StudentRecord).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            End Try
        End Function

        Private Function ValidateRequest(request As StudentSaveRequest,
                                         requireOriginalStudentNumber As Boolean) As String
            If request Is Nothing Then
                Return "Student data is required."
            End If

            If requireOriginalStudentNumber AndAlso
               String.IsNullOrWhiteSpace(request.OriginalStudentNumber) Then
                Return "The original student record is required."
            End If

            If String.IsNullOrWhiteSpace(request.StudentNumber) Then
                Return "Student ID is required."
            End If

            If String.IsNullOrWhiteSpace(request.FirstName) Then
                Return "First Name is required."
            End If

            If String.IsNullOrWhiteSpace(request.LastName) Then
                Return "Last Name is required."
            End If

            If request.YearLevel.HasValue AndAlso
               (request.YearLevel.Value < 1 OrElse request.YearLevel.Value > 6) Then
                Return "Year Level is invalid."
            End If

            If Not String.IsNullOrWhiteSpace(request.Password) AndAlso request.Password.Trim().Length < 8 Then
                Return "Passwords must be at least 8 characters."
            End If

            Return String.Empty
        End Function

        Private Function GetPasswordForCreate(password As String) As String
            If String.IsNullOrWhiteSpace(password) Then
                Return DefaultStudentPassword
            End If

            Return password.Trim()
        End Function

        Private Function ResolveEmailAddress(existingRecord As StudentRecord,
                                             studentNumber As String) As String
            If existingRecord Is Nothing Then
                Return BuildGeneratedEmail(studentNumber)
            End If

            Dim originalGeneratedEmail As String = BuildGeneratedEmail(existingRecord.StudentNumber)
            If String.Equals(existingRecord.Email,
                             originalGeneratedEmail,
                             StringComparison.OrdinalIgnoreCase) Then
                Return BuildGeneratedEmail(studentNumber)
            End If

            Return existingRecord.Email
        End Function

        Private Function BuildGeneratedEmail(studentNumber As String) As String
            Dim builder As New StringBuilder()

            For Each currentCharacter As Char In If(studentNumber, String.Empty).Trim().ToLowerInvariant()
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
                localPart = "student"
            End If

            Return "student." & localPart & "@prmsu.local"
        End Function

        Private Function BuildDatabaseErrorMessage(operationName As String,
                                                   ex As MySqlException) As String
            If ex Is Nothing Then
                Return "Unable to " & operationName & " student records."
            End If

            If ex.Number = 1062 Then
                Return "A duplicate student ID or generated login already exists."
            End If

            Return "Unable to " & operationName & " student records." &
                Environment.NewLine &
                ex.Message
        End Function

        Private Function BuildProfileImageErrorMessage(operationName As String,
                                                       ex As IOException) As String
            If TypeOf ex Is FileNotFoundException Then
                Return "The selected student photo file could not be found."
            End If

            Return "Unable to " & operationName & " the selected student photo locally." &
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
