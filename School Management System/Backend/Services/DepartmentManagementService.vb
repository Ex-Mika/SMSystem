Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class DepartmentManagementService
        Private Class LegacyDepartmentStorageRecord
            Public Property DepartmentId As String
            Public Property DepartmentName As String
            Public Property Head As String
        End Class

        Private ReadOnly _departmentRepository As DepartmentRepository
        Private ReadOnly _legacyDepartmentsStoragePath As String =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "SchoolManagementSystem",
                         "departments.json")
        Private ReadOnly _legacyJsonOptions As New JsonSerializerOptions() With {
            .PropertyNameCaseInsensitive = True
        }
        Private _legacyImportChecked As Boolean

        Public Sub New()
            Me.New(New DepartmentRepository())
        End Sub

        Public Sub New(departmentRepository As DepartmentRepository)
            _departmentRepository = departmentRepository
        End Sub

        Public Function GetDepartments() As ServiceResult(Of List(Of DepartmentRecord))
            Try
                EnsureLegacyDepartmentsImported()

                Dim departments As List(Of DepartmentRecord) = _departmentRepository.GetAll()
                Return ServiceResult(Of List(Of DepartmentRecord)).Success(departments)
            Catch ex As MySqlException
                Return ServiceResult(Of List(Of DepartmentRecord)).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            End Try
        End Function

        Public Function CreateDepartment(request As DepartmentSaveRequest) As ServiceResult(Of DepartmentRecord)
            Dim validationMessage As String = ValidateRequest(request, False)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of DepartmentRecord).Failure(validationMessage)
            End If

            Try
                EnsureLegacyDepartmentsImported()

                If _departmentRepository.GetByDepartmentCode(request.DepartmentCode) IsNot Nothing Then
                    Return ServiceResult(Of DepartmentRecord).Failure("Department ID already exists.")
                End If

                Dim createdRecord As DepartmentRecord = _departmentRepository.Create(request)
                Return ServiceResult(Of DepartmentRecord).Success(createdRecord, "Department created.")
            Catch ex As MySqlException
                Return ServiceResult(Of DepartmentRecord).Failure(
                    BuildDatabaseErrorMessage("create", ex))
            End Try
        End Function

        Public Function UpdateDepartment(request As DepartmentSaveRequest) As ServiceResult(Of DepartmentRecord)
            Dim validationMessage As String = ValidateRequest(request, True)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of DepartmentRecord).Failure(validationMessage)
            End If

            Try
                EnsureLegacyDepartmentsImported()

                Dim existingRecord As DepartmentRecord =
                    _departmentRepository.GetByDepartmentCode(request.OriginalDepartmentCode)
                If existingRecord Is Nothing Then
                    Return ServiceResult(Of DepartmentRecord).Failure(
                        "The selected department no longer exists.")
                End If

                Dim duplicateDepartment As DepartmentRecord =
                    _departmentRepository.GetByDepartmentCode(request.DepartmentCode)
                If duplicateDepartment IsNot Nothing AndAlso
                   duplicateDepartment.DepartmentRecordId <> existingRecord.DepartmentRecordId Then
                    Return ServiceResult(Of DepartmentRecord).Failure("Department ID already exists.")
                End If

                Dim updatedRecord As DepartmentRecord =
                    _departmentRepository.Update(existingRecord, request)
                Return ServiceResult(Of DepartmentRecord).Success(updatedRecord, "Department updated.")
            Catch ex As MySqlException
                Return ServiceResult(Of DepartmentRecord).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            End Try
        End Function

        Public Function DeleteDepartment(departmentCode As String) As ServiceResult(Of Boolean)
            If String.IsNullOrWhiteSpace(departmentCode) Then
                Return ServiceResult(Of Boolean).Failure("Department ID is required.")
            End If

            Try
                EnsureLegacyDepartmentsImported()

                Dim deleted As Boolean =
                    _departmentRepository.DeleteByDepartmentCode(departmentCode.Trim())
                If Not deleted Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected department no longer exists.")
                End If

                Return ServiceResult(Of Boolean).Success(True, "Department deleted.")
            Catch ex As MySqlException
                Return ServiceResult(Of Boolean).Failure(
                    BuildDatabaseErrorMessage("delete", ex))
            End Try
        End Function

        Private Function ValidateRequest(request As DepartmentSaveRequest,
                                         requireOriginalDepartmentCode As Boolean) As String
            If request Is Nothing Then
                Return "Department data is required."
            End If

            If requireOriginalDepartmentCode AndAlso
               String.IsNullOrWhiteSpace(request.OriginalDepartmentCode) Then
                Return "The original department record is required."
            End If

            If String.IsNullOrWhiteSpace(request.DepartmentCode) Then
                Return "Department ID is required."
            End If

            If String.IsNullOrWhiteSpace(request.DepartmentName) Then
                Return "Department Name is required."
            End If

            Return String.Empty
        End Function

        Private Sub EnsureLegacyDepartmentsImported()
            If _legacyImportChecked Then
                Return
            End If

            _legacyImportChecked = True

            If Not File.Exists(_legacyDepartmentsStoragePath) Then
                Return
            End If

            Dim requests As List(Of DepartmentSaveRequest) = ReadLegacyDepartmentRequests()
            If requests.Count = 0 Then
                Return
            End If

            Dim existingCodes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each record As DepartmentRecord In _departmentRepository.GetAll()
                Dim departmentCode As String =
                    If(If(record Is Nothing, Nothing, record.DepartmentCode), String.Empty).Trim()
                If Not String.IsNullOrWhiteSpace(departmentCode) Then
                    existingCodes.Add(departmentCode)
                End If
            Next

            For Each request As DepartmentSaveRequest In requests
                Dim departmentCode As String =
                    If(If(request Is Nothing, Nothing, request.DepartmentCode), String.Empty).Trim()
                Dim departmentName As String =
                    If(If(request Is Nothing, Nothing, request.DepartmentName), String.Empty).Trim()

                If String.IsNullOrWhiteSpace(departmentCode) OrElse
                   String.IsNullOrWhiteSpace(departmentName) OrElse
                   existingCodes.Contains(departmentCode) Then
                    Continue For
                End If

                _departmentRepository.Create(request)
                existingCodes.Add(departmentCode)
            Next

            ArchiveLegacyDepartmentsStorage()
        End Sub

        Private Function ReadLegacyDepartmentRequests() As List(Of DepartmentSaveRequest)
            Dim requests As New List(Of DepartmentSaveRequest)()

            Try
                Dim json As String = File.ReadAllText(_legacyDepartmentsStoragePath)
                If String.IsNullOrWhiteSpace(json) Then
                    Return requests
                End If

                Dim records As List(Of LegacyDepartmentStorageRecord) =
                    JsonSerializer.Deserialize(Of List(Of LegacyDepartmentStorageRecord))(json,
                                                                                          _legacyJsonOptions)
                If records Is Nothing Then
                    Return requests
                End If

                For Each record As LegacyDepartmentStorageRecord In records
                    Dim departmentCode As String =
                        If(If(record Is Nothing, Nothing, record.DepartmentId), String.Empty).Trim()
                    Dim departmentName As String =
                        If(If(record Is Nothing, Nothing, record.DepartmentName), String.Empty).Trim()

                    If String.IsNullOrWhiteSpace(departmentCode) OrElse
                       String.IsNullOrWhiteSpace(departmentName) Then
                        Continue For
                    End If

                    requests.Add(New DepartmentSaveRequest() With {
                        .DepartmentCode = departmentCode,
                        .DepartmentName = departmentName,
                        .HeadName = If(record.Head, String.Empty).Trim()
                    })
                Next
            Catch
                Return New List(Of DepartmentSaveRequest)()
            End Try

            Return requests
        End Function

        Private Sub ArchiveLegacyDepartmentsStorage()
            If Not File.Exists(_legacyDepartmentsStoragePath) Then
                Return
            End If

            Try
                Dim backupPath As String = _legacyDepartmentsStoragePath & ".migrated"
                If File.Exists(backupPath) Then
                    backupPath = _legacyDepartmentsStoragePath & "." &
                        DateTime.UtcNow.ToString("yyyyMMddHHmmss") &
                        ".migrated"
                End If

                File.Move(_legacyDepartmentsStoragePath, backupPath)
            Catch
            End Try
        End Sub

        Private Function BuildDatabaseErrorMessage(operationName As String,
                                                   ex As MySqlException) As String
            If ex Is Nothing Then
                Return "Unable to " & operationName & " department records."
            End If

            If ex.Number = 1062 Then
                Return "A duplicate department ID already exists."
            End If

            Return "Unable to " & operationName & " department records." &
                Environment.NewLine &
                ex.Message
        End Function
    End Class
End Namespace
