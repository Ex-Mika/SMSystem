Imports System.Text
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class StudentManagementService
        Public Const DefaultStudentPassword As String = "Admin@123"

        Private ReadOnly _studentRepository As StudentRepository

        Public Sub New()
            Me.New(New StudentRepository())
        End Sub

        Public Sub New(studentRepository As StudentRepository)
            _studentRepository = studentRepository
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

        Public Function CreateStudent(request As StudentSaveRequest) As ServiceResult(Of StudentRecord)
            Dim validationMessage As String = ValidateRequest(request, False)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of StudentRecord).Failure(validationMessage)
            End If

            Try
                If _studentRepository.GetByStudentNumber(request.StudentNumber) IsNot Nothing Then
                    Return ServiceResult(Of StudentRecord).Failure("Student ID already exists.")
                End If

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
                Return ServiceResult(Of StudentRecord).Failure(ex.Message)
            Catch ex As MySqlException
                Return ServiceResult(Of StudentRecord).Failure(
                    BuildDatabaseErrorMessage("create", ex))
            End Try
        End Function

        Public Function UpdateStudent(request As StudentSaveRequest) As ServiceResult(Of StudentRecord)
            Dim validationMessage As String = ValidateRequest(request, True)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of StudentRecord).Failure(validationMessage)
            End If

            Try
                Dim existingRecord As StudentRecord =
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

                Dim updatedRecord As StudentRecord =
                    _studentRepository.Update(existingRecord,
                                              request,
                                              ResolveEmailAddress(existingRecord, request.StudentNumber),
                                              shouldUpdatePassword,
                                              passwordHash)

                Return ServiceResult(Of StudentRecord).Success(updatedRecord, "Student updated.")
            Catch ex As InvalidOperationException
                Return ServiceResult(Of StudentRecord).Failure(ex.Message)
            Catch ex As MySqlException
                Return ServiceResult(Of StudentRecord).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            End Try
        End Function

        Public Function DeleteStudent(studentNumber As String) As ServiceResult(Of Boolean)
            If String.IsNullOrWhiteSpace(studentNumber) Then
                Return ServiceResult(Of Boolean).Failure("Student ID is required.")
            End If

            Try
                Dim deleted As Boolean =
                    _studentRepository.DeleteByStudentNumber(studentNumber.Trim())
                If Not deleted Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected student no longer exists.")
                End If

                Return ServiceResult(Of Boolean).Success(True, "Student deleted.")
            Catch ex As MySqlException
                Return ServiceResult(Of Boolean).Failure(
                    BuildDatabaseErrorMessage("delete", ex))
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
    End Class
End Namespace
