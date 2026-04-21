Imports System.Collections.Generic
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class StudentGradeManagementService
        Private Const PassingGradeThreshold As Decimal = 75D

        Private ReadOnly _studentGradeRepository As StudentGradeRepository
        Private ReadOnly _studentManagementService As StudentManagementService

        Public Sub New()
            Me.New(New StudentGradeRepository(),
                   New StudentManagementService())
        End Sub

        Public Sub New(studentGradeRepository As StudentGradeRepository,
                       studentManagementService As StudentManagementService)
            _studentGradeRepository =
                If(studentGradeRepository, New StudentGradeRepository())
            _studentManagementService =
                If(studentManagementService, New StudentManagementService())
        End Sub

        Public Function GetTeacherGradeRecords(teacherId As String) As ServiceResult(Of List(Of StudentSubjectGradeRecord))
            If String.IsNullOrWhiteSpace(teacherId) Then
                Return ServiceResult(Of List(Of StudentSubjectGradeRecord)).Failure("Teacher ID is required.")
            End If

            Try
                Dim gradeRecords As List(Of StudentSubjectGradeRecord) =
                    _studentGradeRepository.GetTeacherGradeRoster(teacherId.Trim())
                Return ServiceResult(Of List(Of StudentSubjectGradeRecord)).Success(gradeRecords)
            Catch ex As MySqlException
                Return ServiceResult(Of List(Of StudentSubjectGradeRecord)).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            Catch ex As Exception
                Return ServiceResult(Of List(Of StudentSubjectGradeRecord)).Failure(
                    "Unable to load grade records." &
                    Environment.NewLine &
                    ex.Message)
            End Try
        End Function

        Public Function GetStudentGradeSnapshot(studentNumber As String) As ServiceResult(Of StudentGradeSnapshot)
            If String.IsNullOrWhiteSpace(studentNumber) Then
                Return ServiceResult(Of StudentGradeSnapshot).Failure("Student ID is required.")
            End If

            Dim studentResult As ServiceResult(Of StudentRecord) =
                _studentManagementService.GetStudentByStudentNumber(studentNumber.Trim())
            If studentResult Is Nothing OrElse Not studentResult.IsSuccess Then
                Return ServiceResult(Of StudentGradeSnapshot).Failure(
                    If(studentResult Is Nothing,
                       "Unable to load student grades.",
                       studentResult.Message))
            End If

            Try
                Dim snapshot As New StudentGradeSnapshot() With {
                    .Student = studentResult.Data,
                    .Grades = _studentGradeRepository.GetStudentGradesByStudentNumber(studentNumber.Trim())
                }
                snapshot.NoticeMessage = BuildStudentNoticeMessage(snapshot)
                Return ServiceResult(Of StudentGradeSnapshot).Success(snapshot)
            Catch ex As MySqlException
                Return ServiceResult(Of StudentGradeSnapshot).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            Catch ex As Exception
                Return ServiceResult(Of StudentGradeSnapshot).Failure(
                    "Unable to load grade records." &
                    Environment.NewLine &
                    ex.Message)
            End Try
        End Function

        Public Function SaveTeacherGrade(request As StudentGradeSaveRequest) As ServiceResult(Of StudentSubjectGradeRecord)
            Dim validationMessage As String = ValidateSaveRequest(request)
            If validationMessage <> String.Empty Then
                Return ServiceResult(Of StudentSubjectGradeRecord).Failure(validationMessage)
            End If

            Dim teacherGradeRecordsResult As ServiceResult(Of List(Of StudentSubjectGradeRecord)) =
                GetTeacherGradeRecords(request.TeacherId)
            If teacherGradeRecordsResult Is Nothing OrElse
               Not teacherGradeRecordsResult.IsSuccess Then
                Return ServiceResult(Of StudentSubjectGradeRecord).Failure(
                    If(teacherGradeRecordsResult Is Nothing,
                       "Unable to load the teacher grading roster.",
                       teacherGradeRecordsResult.Message))
            End If

            Dim rosterEntry As StudentSubjectGradeRecord =
                FindRosterEntry(teacherGradeRecordsResult.Data,
                                request.StudentNumber,
                                request.SubjectCode)
            If rosterEntry Is Nothing Then
                Return ServiceResult(Of StudentSubjectGradeRecord).Failure(
                    "This student is not part of your current grading roster.")
            End If

            Dim finalGradeValue As Decimal =
                ComputeFinalGrade(request.QuizScore,
                                  request.ProjectScore,
                                  request.MidtermScore,
                                  request.FinalExamScore)
            Dim remarksValue As String = BuildRemarks(finalGradeValue)

            Try
                If Not _studentGradeRepository.SaveGrade(rosterEntry.TeacherRecordId,
                                                         rosterEntry.StudentRecordId,
                                                         rosterEntry.SubjectId,
                                                         rosterEntry.SectionName,
                                                         request.QuizScore,
                                                         request.ProjectScore,
                                                         request.MidtermScore,
                                                         request.FinalExamScore,
                                                         finalGradeValue,
                                                         remarksValue) Then
                    Return ServiceResult(Of StudentSubjectGradeRecord).Failure(
                        "Unable to save the grade right now.")
                End If

                Dim updatedRosterResult As ServiceResult(Of List(Of StudentSubjectGradeRecord)) =
                    GetTeacherGradeRecords(request.TeacherId)
                If updatedRosterResult Is Nothing OrElse
                   Not updatedRosterResult.IsSuccess Then
                    Return ServiceResult(Of StudentSubjectGradeRecord).Success(
                        BuildFallbackUpdatedRecord(rosterEntry,
                                                   request,
                                                   finalGradeValue,
                                                   remarksValue),
                        "Grade saved.")
                End If

                Dim updatedRecord As StudentSubjectGradeRecord =
                    FindRosterEntry(updatedRosterResult.Data,
                                    request.StudentNumber,
                                    request.SubjectCode)
                If updatedRecord Is Nothing Then
                    updatedRecord =
                        BuildFallbackUpdatedRecord(rosterEntry,
                                                   request,
                                                   finalGradeValue,
                                                   remarksValue)
                End If

                Return ServiceResult(Of StudentSubjectGradeRecord).Success(updatedRecord,
                                                                           "Grade saved.")
            Catch ex As MySqlException
                Return ServiceResult(Of StudentSubjectGradeRecord).Failure(
                    BuildDatabaseErrorMessage("save", ex))
            Catch ex As Exception
                Return ServiceResult(Of StudentSubjectGradeRecord).Failure(
                    "Unable to save grade records." &
                    Environment.NewLine &
                    ex.Message)
            End Try
        End Function

        Private Function ValidateSaveRequest(request As StudentGradeSaveRequest) As String
            If request Is Nothing Then
                Return "Grade details are required."
            End If

            If String.IsNullOrWhiteSpace(request.TeacherId) Then
                Return "Teacher ID is required."
            End If

            If String.IsNullOrWhiteSpace(request.StudentNumber) Then
                Return "Student ID is required."
            End If

            If String.IsNullOrWhiteSpace(request.SubjectCode) Then
                Return "Subject Code is required."
            End If

            Dim scoreValidations As Tuple(Of String, Decimal)() = {
                Tuple.Create("Quiz Score", request.QuizScore),
                Tuple.Create("Project Score", request.ProjectScore),
                Tuple.Create("Midterm Score", request.MidtermScore),
                Tuple.Create("Final Exam Score", request.FinalExamScore)
            }

            For Each scoreValidation As Tuple(Of String, Decimal) In scoreValidations
                If scoreValidation Is Nothing Then
                    Continue For
                End If

                If scoreValidation.Item2 < 0D OrElse scoreValidation.Item2 > 100D Then
                    Return scoreValidation.Item1 & " must be between 0 and 100."
                End If
            Next

            Return String.Empty
        End Function

        Private Function FindRosterEntry(records As IEnumerable(Of StudentSubjectGradeRecord),
                                         studentNumber As String,
                                         subjectCode As String) As StudentSubjectGradeRecord
            Dim normalizedStudentNumber As String = NormalizeText(studentNumber)
            Dim normalizedSubjectCode As String = NormalizeText(subjectCode)
            If records Is Nothing OrElse
               normalizedStudentNumber = String.Empty OrElse
               normalizedSubjectCode = String.Empty Then
                Return Nothing
            End If

            For Each record As StudentSubjectGradeRecord In records
                If record Is Nothing Then
                    Continue For
                End If

                If String.Equals(NormalizeText(record.StudentNumber),
                                 normalizedStudentNumber,
                                 StringComparison.OrdinalIgnoreCase) AndAlso
                   String.Equals(NormalizeText(record.SubjectCode),
                                 normalizedSubjectCode,
                                 StringComparison.OrdinalIgnoreCase) Then
                    Return record
                End If
            Next

            Return Nothing
        End Function

        Private Function ComputeFinalGrade(quizScore As Decimal,
                                           projectScore As Decimal,
                                           midtermScore As Decimal,
                                           finalExamScore As Decimal) As Decimal
            Return Decimal.Round((quizScore + projectScore + midtermScore + finalExamScore) / 4D,
                                 2,
                                 MidpointRounding.AwayFromZero)
        End Function

        Private Function BuildRemarks(finalGrade As Decimal) As String
            If finalGrade >= PassingGradeThreshold Then
                Return "Passed"
            End If

            Return "Failed"
        End Function

        Private Function BuildStudentNoticeMessage(snapshot As StudentGradeSnapshot) As String
            If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
                Return "No student grade data loaded."
            End If

            If Not snapshot.HasGrades Then
                Return "No grades have been posted for your enrolled subjects yet."
            End If

            Return "Your posted grades are now available across the current term tabs."
        End Function

        Private Function BuildFallbackUpdatedRecord(sourceRecord As StudentSubjectGradeRecord,
                                                    request As StudentGradeSaveRequest,
                                                    finalGradeValue As Decimal,
                                                    remarksValue As String) As StudentSubjectGradeRecord
            Dim fallbackRecord As StudentSubjectGradeRecord =
                If(sourceRecord, New StudentSubjectGradeRecord())
            fallbackRecord.GradeRecordId = Math.Max(fallbackRecord.GradeRecordId, 1)
            fallbackRecord.QuizScore = request.QuizScore
            fallbackRecord.ProjectScore = request.ProjectScore
            fallbackRecord.MidtermScore = request.MidtermScore
            fallbackRecord.FinalExamScore = request.FinalExamScore
            fallbackRecord.FinalGrade = finalGradeValue
            fallbackRecord.Remarks = remarksValue
            fallbackRecord.UpdatedAt = DateTime.Now
            Return fallbackRecord
        End Function

        Private Function NormalizeText(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function

        Private Function BuildDatabaseErrorMessage(operationName As String,
                                                   ex As MySqlException) As String
            If ex Is Nothing Then
                Return "Unable to " & operationName & " grade records."
            End If

            Return "Unable to " & operationName & " grade records." &
                Environment.NewLine &
                ex.Message
        End Function
    End Class
End Namespace
