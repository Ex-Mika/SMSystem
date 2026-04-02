Imports System.Text
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class TeacherManagementService
        Public Const DefaultTeacherPassword As String = "Admin@123"

        Private ReadOnly _teacherRepository As TeacherRepository

        Public Sub New()
            Me.New(New TeacherRepository())
        End Sub

        Public Sub New(teacherRepository As TeacherRepository)
            _teacherRepository = teacherRepository
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

        Public Function CreateTeacher(request As TeacherSaveRequest) As ServiceResult(Of TeacherRecord)
            Dim validationMessage As String = ValidateRequest(request, False)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of TeacherRecord).Failure(validationMessage)
            End If

            Try
                If _teacherRepository.GetByEmployeeNumber(request.EmployeeNumber) IsNot Nothing Then
                    Return ServiceResult(Of TeacherRecord).Failure("Teacher ID already exists.")
                End If

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
                Return ServiceResult(Of TeacherRecord).Failure(
                    BuildDatabaseErrorMessage("create", ex))
            End Try
        End Function

        Public Function UpdateTeacher(request As TeacherSaveRequest) As ServiceResult(Of TeacherRecord)
            Dim validationMessage As String = ValidateRequest(request, True)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of TeacherRecord).Failure(validationMessage)
            End If

            Try
                Dim existingRecord As TeacherRecord =
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

                Dim updatedRecord As TeacherRecord =
                    _teacherRepository.Update(existingRecord,
                                              request,
                                              ResolveEmailAddress(existingRecord, request.EmployeeNumber),
                                              shouldUpdatePassword,
                                              passwordHash)

                Return ServiceResult(Of TeacherRecord).Success(updatedRecord, "Teacher updated.")
            Catch ex As MySqlException
                Return ServiceResult(Of TeacherRecord).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            End Try
        End Function

        Public Function DeleteTeacher(employeeNumber As String) As ServiceResult(Of Boolean)
            If String.IsNullOrWhiteSpace(employeeNumber) Then
                Return ServiceResult(Of Boolean).Failure("Teacher ID is required.")
            End If

            Try
                Dim deleted As Boolean =
                    _teacherRepository.DeleteByEmployeeNumber(employeeNumber.Trim())
                If Not deleted Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected teacher no longer exists.")
                End If

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
    End Class
End Namespace
