Imports System.Net.Mail
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class AdministratorManagementService
        Public Const DefaultAdministratorPassword As String = "Admin@123"

        Private ReadOnly _administratorRepository As AdministratorRepository

        Public Sub New()
            Me.New(New AdministratorRepository())
        End Sub

        Public Sub New(administratorRepository As AdministratorRepository)
            _administratorRepository = administratorRepository
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

            Try
                If _administratorRepository.GetByAdministratorCode(request.AdministratorCode) IsNot Nothing Then
                    Return ServiceResult(Of AdministratorRecord).Failure("Administrator ID already exists.")
                End If

                If _administratorRepository.GetByEmail(request.Email) IsNot Nothing Then
                    Return ServiceResult(Of AdministratorRecord).Failure("Email address already exists.")
                End If

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
                Return ServiceResult(Of AdministratorRecord).Failure(
                    BuildDatabaseErrorMessage("create", ex))
            End Try
        End Function

        Public Function UpdateAdministrator(request As AdministratorSaveRequest) As ServiceResult(Of AdministratorRecord)
            Dim validationMessage As String = ValidateRequest(request, True)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of AdministratorRecord).Failure(validationMessage)
            End If

            Try
                Dim existingRecord As AdministratorRecord =
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

                Dim updatedRecord As AdministratorRecord =
                    _administratorRepository.Update(existingRecord,
                                                    request,
                                                    shouldUpdatePassword,
                                                    passwordHash)

                Return ServiceResult(Of AdministratorRecord).Success(updatedRecord,
                                                                     "Administrator updated.")
            Catch ex As MySqlException
                Return ServiceResult(Of AdministratorRecord).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            End Try
        End Function

        Public Function DeleteAdministrator(administratorCode As String) As ServiceResult(Of Boolean)
            If String.IsNullOrWhiteSpace(administratorCode) Then
                Return ServiceResult(Of Boolean).Failure("Administrator ID is required.")
            End If

            Try
                Dim deleted As Boolean =
                    _administratorRepository.DeleteByAdministratorCode(administratorCode.Trim())
                If Not deleted Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected administrator no longer exists.")
                End If

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
    End Class
End Namespace
