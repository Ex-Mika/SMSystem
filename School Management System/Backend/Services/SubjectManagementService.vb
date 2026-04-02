Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text.Json
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class SubjectManagementService
        Private Class LegacySubjectStorageRecord
            Public Property SubjectCode As String
            Public Property SubjectName As String
            Public Property Course As String
            Public Property YearLevel As String
            Public Property Units As String
        End Class

        Private ReadOnly _subjectRepository As SubjectRepository
        Private ReadOnly _legacySubjectsStoragePath As String =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "SchoolManagementSystem",
                         "subjects.json")
        Private ReadOnly _legacyJsonOptions As New JsonSerializerOptions() With {
            .PropertyNameCaseInsensitive = True
        }
        Private _legacyImportChecked As Boolean

        Public Sub New()
            Me.New(New SubjectRepository())
        End Sub

        Public Sub New(subjectRepository As SubjectRepository)
            _subjectRepository = subjectRepository
        End Sub

        Public Function GetSubjects() As ServiceResult(Of List(Of SubjectRecord))
            Try
                EnsureLegacySubjectsImported()

                Dim subjects As List(Of SubjectRecord) = _subjectRepository.GetAll()
                Return ServiceResult(Of List(Of SubjectRecord)).Success(subjects)
            Catch ex As MySqlException
                Return ServiceResult(Of List(Of SubjectRecord)).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            End Try
        End Function

        Public Function CreateSubject(request As SubjectSaveRequest) As ServiceResult(Of SubjectRecord)
            Dim validationMessage As String = ValidateRequest(request, False)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of SubjectRecord).Failure(validationMessage)
            End If

            Try
                EnsureLegacySubjectsImported()

                If _subjectRepository.GetBySubjectCode(request.SubjectCode) IsNot Nothing Then
                    Return ServiceResult(Of SubjectRecord).Failure("Subject Code already exists.")
                End If

                Dim createdRecord As SubjectRecord = _subjectRepository.Create(request)
                Return ServiceResult(Of SubjectRecord).Success(createdRecord, "Subject created.")
            Catch ex As MySqlException
                Return ServiceResult(Of SubjectRecord).Failure(
                    BuildDatabaseErrorMessage("create", ex))
            End Try
        End Function

        Public Function UpdateSubject(request As SubjectSaveRequest) As ServiceResult(Of SubjectRecord)
            Dim validationMessage As String = ValidateRequest(request, True)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of SubjectRecord).Failure(validationMessage)
            End If

            Try
                EnsureLegacySubjectsImported()

                Dim existingRecord As SubjectRecord =
                    _subjectRepository.GetBySubjectCode(request.OriginalSubjectCode)
                If existingRecord Is Nothing Then
                    Return ServiceResult(Of SubjectRecord).Failure(
                        "The selected subject no longer exists.")
                End If

                Dim duplicateSubject As SubjectRecord =
                    _subjectRepository.GetBySubjectCode(request.SubjectCode)
                If duplicateSubject IsNot Nothing AndAlso
                   duplicateSubject.SubjectId <> existingRecord.SubjectId Then
                    Return ServiceResult(Of SubjectRecord).Failure("Subject Code already exists.")
                End If

                Dim updatedRecord As SubjectRecord =
                    _subjectRepository.Update(existingRecord, request)

                Return ServiceResult(Of SubjectRecord).Success(updatedRecord, "Subject updated.")
            Catch ex As MySqlException
                Return ServiceResult(Of SubjectRecord).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            End Try
        End Function

        Public Function DeleteSubject(subjectCode As String) As ServiceResult(Of Boolean)
            If String.IsNullOrWhiteSpace(subjectCode) Then
                Return ServiceResult(Of Boolean).Failure("Subject Code is required.")
            End If

            Try
                EnsureLegacySubjectsImported()

                Dim deleted As Boolean =
                    _subjectRepository.DeleteBySubjectCode(subjectCode.Trim())
                If Not deleted Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected subject no longer exists.")
                End If

                Return ServiceResult(Of Boolean).Success(True, "Subject deleted.")
            Catch ex As MySqlException
                Return ServiceResult(Of Boolean).Failure(
                    BuildDatabaseErrorMessage("delete", ex))
            End Try
        End Function

        Private Function ValidateRequest(request As SubjectSaveRequest,
                                         requireOriginalSubjectCode As Boolean) As String
            If request Is Nothing Then
                Return "Subject data is required."
            End If

            request.OriginalSubjectCode = If(request.OriginalSubjectCode, String.Empty).Trim()
            request.SubjectCode = If(request.SubjectCode, String.Empty).Trim()
            request.SubjectName = If(request.SubjectName, String.Empty).Trim()
            request.CourseText = If(request.CourseText, String.Empty).Trim()
            request.YearLevel = If(request.YearLevel, String.Empty).Trim()

            If requireOriginalSubjectCode AndAlso
               String.IsNullOrWhiteSpace(request.OriginalSubjectCode) Then
                Return "The original subject record is required."
            End If

            If String.IsNullOrWhiteSpace(request.SubjectCode) Then
                Return "Subject Code is required."
            End If

            If String.IsNullOrWhiteSpace(request.SubjectName) Then
                Return "Subject Name is required."
            End If

            Dim normalizedUnits As String = NormalizeUnitsText(request.Units)
            If String.IsNullOrWhiteSpace(normalizedUnits) Then
                Return "Units must be a valid number with up to one decimal place."
            End If

            request.Units = normalizedUnits
            Return String.Empty
        End Function

        Private Sub EnsureLegacySubjectsImported()
            If _legacyImportChecked Then
                Return
            End If

            _legacyImportChecked = True

            If Not File.Exists(_legacySubjectsStoragePath) Then
                Return
            End If

            Dim requests As List(Of SubjectSaveRequest) = ReadLegacySubjectRequests()
            If requests.Count = 0 Then
                Return
            End If

            Dim existingCodes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each record As SubjectRecord In _subjectRepository.GetAll()
                Dim subjectCode As String = If(If(record Is Nothing, Nothing, record.SubjectCode), String.Empty).Trim()
                If Not String.IsNullOrWhiteSpace(subjectCode) Then
                    existingCodes.Add(subjectCode)
                End If
            Next

            For Each request As SubjectSaveRequest In requests
                Dim subjectCode As String = If(If(request Is Nothing, Nothing, request.SubjectCode), String.Empty).Trim()
                Dim subjectName As String = If(If(request Is Nothing, Nothing, request.SubjectName), String.Empty).Trim()

                If String.IsNullOrWhiteSpace(subjectCode) OrElse
                   String.IsNullOrWhiteSpace(subjectName) OrElse
                   existingCodes.Contains(subjectCode) OrElse
                   Not String.IsNullOrWhiteSpace(ValidateRequest(request, False)) Then
                    Continue For
                End If

                _subjectRepository.Create(request)
                existingCodes.Add(subjectCode)
            Next

            ArchiveLegacySubjectsStorage()
        End Sub

        Private Function ReadLegacySubjectRequests() As List(Of SubjectSaveRequest)
            Dim requests As New List(Of SubjectSaveRequest)()

            Try
                Dim json As String = File.ReadAllText(_legacySubjectsStoragePath)
                If String.IsNullOrWhiteSpace(json) Then
                    Return requests
                End If

                Dim records As List(Of LegacySubjectStorageRecord) =
                    JsonSerializer.Deserialize(Of List(Of LegacySubjectStorageRecord))(json, _legacyJsonOptions)
                If records Is Nothing Then
                    Return requests
                End If

                For Each record As LegacySubjectStorageRecord In records
                    Dim subjectCode As String = If(If(record Is Nothing, Nothing, record.SubjectCode), String.Empty).Trim()
                    Dim subjectName As String = If(If(record Is Nothing, Nothing, record.SubjectName), String.Empty).Trim()

                    If String.IsNullOrWhiteSpace(subjectCode) OrElse
                       String.IsNullOrWhiteSpace(subjectName) Then
                        Continue For
                    End If

                    requests.Add(New SubjectSaveRequest() With {
                        .SubjectCode = subjectCode,
                        .SubjectName = subjectName,
                        .CourseText = If(record.Course, String.Empty).Trim(),
                        .YearLevel = If(record.YearLevel, String.Empty).Trim(),
                        .Units = If(record.Units, String.Empty).Trim()
                    })
                Next
            Catch
                Return New List(Of SubjectSaveRequest)()
            End Try

            Return requests
        End Function

        Private Sub ArchiveLegacySubjectsStorage()
            If Not File.Exists(_legacySubjectsStoragePath) Then
                Return
            End If

            Try
                Dim backupPath As String = _legacySubjectsStoragePath & ".migrated"
                If File.Exists(backupPath) Then
                    backupPath = _legacySubjectsStoragePath & "." &
                        DateTime.UtcNow.ToString("yyyyMMddHHmmss") &
                        ".migrated"
                End If

                File.Move(_legacySubjectsStoragePath, backupPath)
            Catch
            End Try
        End Sub

        Private Function NormalizeUnitsText(value As String) As String
            Dim normalizedValue As String = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Return "3"
            End If

            Dim parsedUnits As Decimal
            If Not (Decimal.TryParse(normalizedValue,
                                     NumberStyles.Number,
                                     CultureInfo.InvariantCulture,
                                     parsedUnits) OrElse
                    Decimal.TryParse(normalizedValue,
                                     NumberStyles.Number,
                                     CultureInfo.CurrentCulture,
                                     parsedUnits)) Then
                Return String.Empty
            End If

            If parsedUnits <= 0D Then
                Return String.Empty
            End If

            Dim roundedUnits As Decimal = Decimal.Round(parsedUnits, 1, MidpointRounding.AwayFromZero)
            If roundedUnits <> parsedUnits Then
                Return String.Empty
            End If

            Return roundedUnits.ToString("0.#", CultureInfo.InvariantCulture)
        End Function

        Private Function BuildDatabaseErrorMessage(operationName As String,
                                                   ex As MySqlException) As String
            If ex Is Nothing Then
                Return "Unable to " & operationName & " subject records."
            End If

            If ex.Number = 1062 Then
                Return "A duplicate subject code already exists."
            End If

            Return "Unable to " & operationName & " subject records." &
                Environment.NewLine &
                ex.Message
        End Function
    End Class
End Namespace
