Imports System.Net.Mail
Imports System.IO
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class AdministratorManagementService
        Public Const DefaultAdministratorPassword As String = "Admin@123"

        Private ReadOnly _administratorRepository As AdministratorRepository
        Private ReadOnly _profileImageStorageService As ProfileImageStorageService

        Public Sub New()
            Me.New(New AdministratorRepository(), New ProfileImageStorageService())
        End Sub

        Public Sub New(administratorRepository As AdministratorRepository)
            Me.New(administratorRepository, New ProfileImageStorageService())
        End Sub

        Public Sub New(administratorRepository As AdministratorRepository,
                       profileImageStorageService As ProfileImageStorageService)
            _administratorRepository = administratorRepository
            _profileImageStorageService =
                If(profileImageStorageService, New ProfileImageStorageService())
        End Sub

        Public Function GetAdministrators() As ServiceResult(Of List(Of AdministratorRecord))
            Try
                Dim administrators As List(Of AdministratorRecord) = _administratorRepository.GetAll()
                Return ServiceResult(Of List(Of AdministratorRecord)).Success(administrators)
            Catch ex As MySqlException
                Return ServiceResult(Of List(Of AdministratorRecord)).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            End Try
        End Function

        Public Function CreateAdministrator(request As AdministratorSaveRequest) As ServiceResult(Of AdministratorRecord)
            Dim validationMessage As String = ValidateRequest(request, False)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of AdministratorRecord).Failure(validationMessage)
            End If

            Dim preparedPhotoPath As String = String.Empty

            Try
                If _administratorRepository.GetByAdministratorCode(request.AdministratorCode) IsNot Nothing Then
                    Return ServiceResult(Of AdministratorRecord).Failure("Administrator ID already exists.")
                End If

                If _administratorRepository.GetByEmail(request.Email) IsNot Nothing Then
                    Return ServiceResult(Of AdministratorRecord).Failure("Email address already exists.")
                End If

                preparedPhotoPath =
                    _profileImageStorageService.StoreProfileImage(
                        ProfileImageStorageService.ProfileImageOwnerType.Administrator,
                        request.AdministratorCode,
                        request.PhotoPath)
                request.PhotoPath = preparedPhotoPath

                Dim passwordToUse As String = GetPasswordForCreate(request.Password)
                Dim createdRecord As AdministratorRecord =
                    _administratorRepository.Create(request,
                                                    Database.DatabaseModule.HashPassword(passwordToUse))

                Dim message As String = "Administrator created."
                If String.IsNullOrWhiteSpace(request.Password) Then
                    message = "Administrator created. Temporary password: " & DefaultAdministratorPassword
                End If

                Return ServiceResult(Of AdministratorRecord).Success(createdRecord, message)
            Catch ex As MySqlException
                CleanupPreparedProfileImage(preparedPhotoPath, String.Empty)
                Return ServiceResult(Of AdministratorRecord).Failure(
                    BuildDatabaseErrorMessage("create", ex))
            Catch ex As IOException
                CleanupPreparedProfileImage(preparedPhotoPath, String.Empty)
                Return ServiceResult(Of AdministratorRecord).Failure(
                    BuildProfileImageErrorMessage("save", ex))
            End Try
        End Function

        Public Function UpdateAdministrator(request As AdministratorSaveRequest) As ServiceResult(Of AdministratorRecord)
            Dim validationMessage As String = ValidateRequest(request, True)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of AdministratorRecord).Failure(validationMessage)
            End If

            Dim existingRecord As AdministratorRecord = Nothing
            Dim preparedPhotoPath As String = String.Empty

            Try
                existingRecord =
                    _administratorRepository.GetByAdministratorCode(request.OriginalAdministratorCode)
                If existingRecord Is Nothing Then
                    Return ServiceResult(Of AdministratorRecord).Failure(
                        "The selected administrator no longer exists.")
                End If

                Dim duplicateCodeRecord As AdministratorRecord =
                    _administratorRepository.GetByAdministratorCode(request.AdministratorCode)
                If duplicateCodeRecord IsNot Nothing AndAlso
                   duplicateCodeRecord.AdministratorRecordId <> existingRecord.AdministratorRecordId Then
                    Return ServiceResult(Of AdministratorRecord).Failure("Administrator ID already exists.")
                End If

                Dim duplicateEmailRecord As AdministratorRecord =
                    _administratorRepository.GetByEmail(request.Email)
                If duplicateEmailRecord IsNot Nothing AndAlso
                   duplicateEmailRecord.UserId <> existingRecord.UserId Then
                    Return ServiceResult(Of AdministratorRecord).Failure("Email address already exists.")
                End If

                Dim shouldUpdatePassword As Boolean = Not String.IsNullOrWhiteSpace(request.Password)
                Dim passwordHash As String = String.Empty

                If shouldUpdatePassword Then
                    passwordHash = Database.DatabaseModule.HashPassword(request.Password.Trim())
                End If

                preparedPhotoPath =
                    _profileImageStorageService.StoreProfileImage(
                        ProfileImageStorageService.ProfileImageOwnerType.Administrator,
                        request.AdministratorCode,
                        request.PhotoPath,
                        existingRecord.PhotoPath)
                request.PhotoPath = preparedPhotoPath

                Dim updatedRecord As AdministratorRecord =
                    _administratorRepository.Update(existingRecord,
                                                    request,
                                                    shouldUpdatePassword,
                                                    passwordHash)

                CleanupObsoleteManagedImage(existingRecord.PhotoPath, updatedRecord.PhotoPath)
                Return ServiceResult(Of AdministratorRecord).Success(updatedRecord,
                                                                     "Administrator updated.")
            Catch ex As MySqlException
                CleanupPreparedProfileImage(preparedPhotoPath,
                                            If(existingRecord Is Nothing,
                                               String.Empty,
                                               existingRecord.PhotoPath))
                Return ServiceResult(Of AdministratorRecord).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            Catch ex As IOException
                CleanupPreparedProfileImage(preparedPhotoPath,
                                            If(existingRecord Is Nothing,
                                               String.Empty,
                                               existingRecord.PhotoPath))
                Return ServiceResult(Of AdministratorRecord).Failure(
                    BuildProfileImageErrorMessage("save", ex))
            End Try
        End Function

        Public Function DeleteAdministrator(administratorCode As String) As ServiceResult(Of Boolean)
            If String.IsNullOrWhiteSpace(administratorCode) Then
                Return ServiceResult(Of Boolean).Failure("Administrator ID is required.")
            End If

            Try
                Dim existingRecord As AdministratorRecord =
                    _administratorRepository.GetByAdministratorCode(administratorCode.Trim())
                If existingRecord Is Nothing Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected administrator no longer exists.")
                End If

                Dim deleted As Boolean =
                    _administratorRepository.DeleteByAdministratorCode(administratorCode.Trim())
                If Not deleted Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected administrator no longer exists.")
                End If

                CleanupObsoleteManagedImage(existingRecord.PhotoPath, String.Empty)
                Return ServiceResult(Of Boolean).Success(True, "Administrator deleted.")
            Catch ex As MySqlException
                Return ServiceResult(Of Boolean).Failure(
                    BuildDatabaseErrorMessage("delete", ex))
            End Try
        End Function

        Private Function ValidateRequest(request As AdministratorSaveRequest,
                                         requireOriginalAdministratorCode As Boolean) As String
            If request Is Nothing Then
                Return "Administrator data is required."
            End If

            If requireOriginalAdministratorCode AndAlso
               String.IsNullOrWhiteSpace(request.OriginalAdministratorCode) Then
                Return "The original administrator record is required."
            End If

            If String.IsNullOrWhiteSpace(request.AdministratorCode) Then
                Return "Administrator ID is required."
            End If

            If String.IsNullOrWhiteSpace(request.FirstName) Then
                Return "First Name is required."
            End If

            If String.IsNullOrWhiteSpace(request.LastName) Then
                Return "Last Name is required."
            End If

            If String.IsNullOrWhiteSpace(request.Email) Then
                Return "Email is required."
            End If

            If Not IsValidEmailAddress(request.Email) Then
                Return "Enter a valid email address."
            End If

            If Not String.IsNullOrWhiteSpace(request.Password) AndAlso request.Password.Trim().Length < 8 Then
                Return "Passwords must be at least 8 characters."
            End If

            Return String.Empty
        End Function

        Private Function IsValidEmailAddress(email As String) As Boolean
            Try
                Dim address As New MailAddress(email.Trim())
                Return String.Equals(address.Address,
                                     email.Trim(),
                                     StringComparison.OrdinalIgnoreCase)
            Catch
                Return False
            End Try
        End Function

        Private Function GetPasswordForCreate(password As String) As String
            If String.IsNullOrWhiteSpace(password) Then
                Return DefaultAdministratorPassword
            End If

            Return password.Trim()
        End Function

        Private Function BuildDatabaseErrorMessage(operationName As String,
                                                   ex As MySqlException) As String
            If ex Is Nothing Then
                Return "Unable to " & operationName & " administrator records."
            End If

            If ex.Number = 1062 Then
                Return "A duplicate administrator ID or email already exists."
            End If

            Return "Unable to " & operationName & " administrator records." &
                Environment.NewLine &
                ex.Message
        End Function

        Private Function BuildProfileImageErrorMessage(operationName As String,
                                                       ex As IOException) As String
            If TypeOf ex Is FileNotFoundException Then
                Return "The selected administrator photo file could not be found."
            End If

            Return "Unable to " & operationName &
                " the selected administrator photo locally." &
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
