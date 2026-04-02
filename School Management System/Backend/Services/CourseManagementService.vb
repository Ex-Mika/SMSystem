Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class CourseManagementService
        Private Class LegacyCourseStorageRecord
            Public Property CourseCode As String
            Public Property CourseTitle As String
            Public Property Department As String
            Public Property Units As String
        End Class

        Private ReadOnly _courseRepository As CourseRepository
        Private ReadOnly _legacyCoursesStoragePath As String =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "SchoolManagementSystem",
                         "courses.json")
        Private ReadOnly _legacyJsonOptions As New JsonSerializerOptions() With {
            .PropertyNameCaseInsensitive = True
        }
        Private _legacyImportChecked As Boolean

        Public Sub New()
            Me.New(New CourseRepository())
        End Sub

        Public Sub New(courseRepository As CourseRepository)
            _courseRepository = courseRepository
        End Sub

        Public Function GetCourses() As ServiceResult(Of List(Of CourseRecord))
            Try
                EnsureLegacyCoursesImported()

                Dim courses As List(Of CourseRecord) = _courseRepository.GetAll()
                Return ServiceResult(Of List(Of CourseRecord)).Success(courses)
            Catch ex As MySqlException
                Return ServiceResult(Of List(Of CourseRecord)).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            End Try
        End Function

        Public Function CreateCourse(request As CourseSaveRequest) As ServiceResult(Of CourseRecord)
            Dim validationMessage As String = ValidateRequest(request, False)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of CourseRecord).Failure(validationMessage)
            End If

            Try
                EnsureLegacyCoursesImported()

                If _courseRepository.GetByCourseCode(request.CourseCode) IsNot Nothing Then
                    Return ServiceResult(Of CourseRecord).Failure("Course Code already exists.")
                End If

                Dim createdRecord As CourseRecord = _courseRepository.Create(request)
                Return ServiceResult(Of CourseRecord).Success(createdRecord, "Course created.")
            Catch ex As MySqlException
                Return ServiceResult(Of CourseRecord).Failure(
                    BuildDatabaseErrorMessage("create", ex))
            End Try
        End Function

        Public Function UpdateCourse(request As CourseSaveRequest) As ServiceResult(Of CourseRecord)
            Dim validationMessage As String = ValidateRequest(request, True)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of CourseRecord).Failure(validationMessage)
            End If

            Try
                EnsureLegacyCoursesImported()

                Dim existingRecord As CourseRecord =
                    _courseRepository.GetByCourseCode(request.OriginalCourseCode)
                If existingRecord Is Nothing Then
                    Return ServiceResult(Of CourseRecord).Failure(
                        "The selected course no longer exists.")
                End If

                Dim duplicateCourse As CourseRecord =
                    _courseRepository.GetByCourseCode(request.CourseCode)
                If duplicateCourse IsNot Nothing AndAlso
                   duplicateCourse.CourseId <> existingRecord.CourseId Then
                    Return ServiceResult(Of CourseRecord).Failure("Course Code already exists.")
                End If

                Dim updatedRecord As CourseRecord =
                    _courseRepository.Update(existingRecord, request)

                Return ServiceResult(Of CourseRecord).Success(updatedRecord, "Course updated.")
            Catch ex As MySqlException
                Return ServiceResult(Of CourseRecord).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            End Try
        End Function

        Public Function DeleteCourse(courseCode As String) As ServiceResult(Of Boolean)
            If String.IsNullOrWhiteSpace(courseCode) Then
                Return ServiceResult(Of Boolean).Failure("Course Code is required.")
            End If

            Try
                EnsureLegacyCoursesImported()

                Dim deleted As Boolean =
                    _courseRepository.DeleteByCourseCode(courseCode.Trim())
                If Not deleted Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected course no longer exists.")
                End If

                Return ServiceResult(Of Boolean).Success(True, "Course deleted.")
            Catch ex As MySqlException
                Return ServiceResult(Of Boolean).Failure(
                    BuildDatabaseErrorMessage("delete", ex))
            End Try
        End Function

        Private Function ValidateRequest(request As CourseSaveRequest,
                                         requireOriginalCourseCode As Boolean) As String
            If request Is Nothing Then
                Return "Course data is required."
            End If

            If requireOriginalCourseCode AndAlso
               String.IsNullOrWhiteSpace(request.OriginalCourseCode) Then
                Return "The original course record is required."
            End If

            If String.IsNullOrWhiteSpace(request.CourseCode) Then
                Return "Course Code is required."
            End If

            If String.IsNullOrWhiteSpace(request.CourseName) Then
                Return "Course Title is required."
            End If

            Return String.Empty
        End Function

        Private Sub EnsureLegacyCoursesImported()
            If _legacyImportChecked Then
                Return
            End If

            _legacyImportChecked = True

            If Not File.Exists(_legacyCoursesStoragePath) Then
                Return
            End If

            Dim requests As List(Of CourseSaveRequest) = ReadLegacyCourseRequests()
            If requests.Count = 0 Then
                Return
            End If

            Dim existingCodes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each record As CourseRecord In _courseRepository.GetAll()
                Dim courseCode As String = If(If(record Is Nothing, Nothing, record.CourseCode), String.Empty).Trim()
                If Not String.IsNullOrWhiteSpace(courseCode) Then
                    existingCodes.Add(courseCode)
                End If
            Next

            For Each request As CourseSaveRequest In requests
                Dim courseCode As String = If(If(request Is Nothing, Nothing, request.CourseCode), String.Empty).Trim()
                If String.IsNullOrWhiteSpace(courseCode) OrElse
                   String.IsNullOrWhiteSpace(If(If(request Is Nothing, Nothing, request.CourseName), String.Empty).Trim()) OrElse
                   existingCodes.Contains(courseCode) Then
                    Continue For
                End If

                _courseRepository.Create(request)
                existingCodes.Add(courseCode)
            Next

            ArchiveLegacyCoursesStorage()
        End Sub

        Private Function ReadLegacyCourseRequests() As List(Of CourseSaveRequest)
            Dim requests As New List(Of CourseSaveRequest)()

            Try
                Dim json As String = File.ReadAllText(_legacyCoursesStoragePath)
                If String.IsNullOrWhiteSpace(json) Then
                    Return requests
                End If

                Dim records As List(Of LegacyCourseStorageRecord) =
                    JsonSerializer.Deserialize(Of List(Of LegacyCourseStorageRecord))(json, _legacyJsonOptions)
                If records Is Nothing Then
                    Return requests
                End If

                For Each record As LegacyCourseStorageRecord In records
                    Dim courseCode As String = If(If(record Is Nothing, Nothing, record.CourseCode), String.Empty).Trim()
                    Dim courseName As String = If(If(record Is Nothing, Nothing, record.CourseTitle), String.Empty).Trim()

                    If String.IsNullOrWhiteSpace(courseCode) OrElse
                       String.IsNullOrWhiteSpace(courseName) Then
                        Continue For
                    End If

                    requests.Add(New CourseSaveRequest() With {
                        .CourseCode = courseCode,
                        .CourseName = courseName,
                        .DepartmentText = If(record.Department, String.Empty).Trim(),
                        .Units = If(record.Units, String.Empty).Trim()
                    })
                Next
            Catch
                Return New List(Of CourseSaveRequest)()
            End Try

            Return requests
        End Function

        Private Sub ArchiveLegacyCoursesStorage()
            If Not File.Exists(_legacyCoursesStoragePath) Then
                Return
            End If

            Try
                Dim backupPath As String = _legacyCoursesStoragePath & ".migrated"
                If File.Exists(backupPath) Then
                    backupPath = _legacyCoursesStoragePath & "." &
                        DateTime.UtcNow.ToString("yyyyMMddHHmmss") &
                        ".migrated"
                End If

                File.Move(_legacyCoursesStoragePath, backupPath)
            Catch
            End Try
        End Sub

        Private Function BuildDatabaseErrorMessage(operationName As String,
                                                   ex As MySqlException) As String
            If ex Is Nothing Then
                Return "Unable to " & operationName & " course records."
            End If

            If ex.Number = 1062 Then
                Return "A duplicate course code already exists."
            End If

            Return "Unable to " & operationName & " course records." &
                Environment.NewLine &
                ex.Message
        End Function
    End Class
End Namespace
