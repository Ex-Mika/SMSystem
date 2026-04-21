Imports System.Collections.Generic
Imports System.Text
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class StudentEnrollmentManagementService
        Private Class EnrollmentContext
            Public Property Student As StudentRecord
            Public Property Subjects As New List(Of SubjectRecord)()
        End Class

        Private ReadOnly _studentManagementService As StudentManagementService
        Private ReadOnly _subjectManagementService As SubjectManagementService
        Private ReadOnly _teacherScheduleManagementService As TeacherScheduleManagementService
        Private ReadOnly _enrollmentRepository As StudentSubjectEnrollmentRepository

        Public Sub New()
            Me.New(New StudentManagementService(),
                   New SubjectManagementService(),
                   New TeacherScheduleManagementService(),
                   New StudentSubjectEnrollmentRepository())
        End Sub

        Public Sub New(studentManagementService As StudentManagementService,
                       subjectManagementService As SubjectManagementService,
                       enrollmentRepository As StudentSubjectEnrollmentRepository)
            Me.New(studentManagementService,
                   subjectManagementService,
                   New TeacherScheduleManagementService(),
                   enrollmentRepository)
        End Sub

        Public Sub New(studentManagementService As StudentManagementService,
                       subjectManagementService As SubjectManagementService,
                       teacherScheduleManagementService As TeacherScheduleManagementService,
                       enrollmentRepository As StudentSubjectEnrollmentRepository)
            _studentManagementService = studentManagementService
            _subjectManagementService = subjectManagementService
            _teacherScheduleManagementService = teacherScheduleManagementService
            _enrollmentRepository = enrollmentRepository
        End Sub

        Public Function GetEnrollmentSnapshot(studentNumber As String) As ServiceResult(Of StudentEnrollmentSnapshot)
            Dim contextResult As ServiceResult(Of EnrollmentContext) = LoadContext(studentNumber)
            If Not contextResult.IsSuccess Then
                Return ServiceResult(Of StudentEnrollmentSnapshot).Failure(contextResult.Message)
            End If

            Try
                Dim snapshot As StudentEnrollmentSnapshot = BuildSnapshot(contextResult.Data)
                Return ServiceResult(Of StudentEnrollmentSnapshot).Success(snapshot)
            Catch ex As MySqlException
                Return ServiceResult(Of StudentEnrollmentSnapshot).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            Catch ex As Exception
                Return ServiceResult(Of StudentEnrollmentSnapshot).Failure(
                    "Unable to load enrollment records." &
                    Environment.NewLine &
                    ex.Message)
            End Try
        End Function

        Public Function EnrollSubject(studentNumber As String,
                                      subjectCode As String) As ServiceResult(Of Boolean)
            Dim normalizedSubjectCode As String = NormalizeText(subjectCode)
            If String.IsNullOrWhiteSpace(normalizedSubjectCode) Then
                Return ServiceResult(Of Boolean).Failure("Subject Code is required.")
            End If

            Dim contextResult As ServiceResult(Of EnrollmentContext) = LoadContext(studentNumber)
            If Not contextResult.IsSuccess Then
                Return ServiceResult(Of Boolean).Failure(contextResult.Message)
            End If

            Dim context As EnrollmentContext = contextResult.Data
            Dim subject As SubjectRecord = FindSubjectByCode(context.Subjects, normalizedSubjectCode)
            If subject Is Nothing Then
                Return ServiceResult(Of Boolean).Failure(
                    "The selected subject no longer exists.")
            End If

            If Not IsSubjectAvailableToStudent(context.Student, subject) Then
                Return ServiceResult(Of Boolean).Failure(
                    "This subject is not available for your year level.")
            End If

            Try
                If _enrollmentRepository.Exists(context.Student.StudentRecordId, subject.SubjectId) Then
                    Return ServiceResult(Of Boolean).Failure(
                        "You are already enrolled in this subject.")
                End If

                If Not _enrollmentRepository.Create(context.Student.StudentRecordId, subject.SubjectId) Then
                    Return ServiceResult(Of Boolean).Failure(
                        "Unable to add the selected subject right now.")
                End If

                Return ServiceResult(Of Boolean).Success(True,
                                                         "Subject added to your selected load.")
            Catch ex As MySqlException
                Return ServiceResult(Of Boolean).Failure(
                    BuildDatabaseErrorMessage("save", ex))
            End Try
        End Function

        Public Function RemoveSubject(studentNumber As String,
                                      subjectCode As String) As ServiceResult(Of Boolean)
            Dim normalizedSubjectCode As String = NormalizeText(subjectCode)
            If String.IsNullOrWhiteSpace(normalizedSubjectCode) Then
                Return ServiceResult(Of Boolean).Failure("Subject Code is required.")
            End If

            Dim contextResult As ServiceResult(Of EnrollmentContext) = LoadContext(studentNumber)
            If Not contextResult.IsSuccess Then
                Return ServiceResult(Of Boolean).Failure(contextResult.Message)
            End If

            Dim context As EnrollmentContext = contextResult.Data
            Dim subject As SubjectRecord = FindSubjectByCode(context.Subjects, normalizedSubjectCode)
            If subject Is Nothing Then
                Return ServiceResult(Of Boolean).Failure(
                    "The selected subject no longer exists.")
            End If

            Try
                If Not _enrollmentRepository.Delete(context.Student.StudentRecordId, subject.SubjectId) Then
                    Return ServiceResult(Of Boolean).Failure(
                        "The selected subject is not in your current load.")
                End If

                Return ServiceResult(Of Boolean).Success(True,
                                                         "Subject removed from your selected load.")
            Catch ex As MySqlException
                Return ServiceResult(Of Boolean).Failure(
                    BuildDatabaseErrorMessage("update", ex))
            End Try
        End Function

        Public Function UpdateStudentSection(studentNumber As String,
                                             sectionName As String) As ServiceResult(Of StudentRecord)
            Dim normalizedSectionName As String = NormalizeText(sectionName)
            If String.IsNullOrWhiteSpace(normalizedSectionName) Then
                Return ServiceResult(Of StudentRecord).Failure("Section is required.")
            End If

            Dim contextResult As ServiceResult(Of EnrollmentContext) = LoadContext(studentNumber)
            If Not contextResult.IsSuccess Then
                Return ServiceResult(Of StudentRecord).Failure(contextResult.Message)
            End If

            Dim context As EnrollmentContext = contextResult.Data
            Dim availableSections As List(Of StudentSectionOption) =
                BuildAvailableSections(context.Student, context.Subjects)
            If Not ContainsSectionOption(availableSections,
                                         normalizedSectionName,
                                         context.Student) Then
                Return ServiceResult(Of StudentRecord).Failure(
                    "This section is not available for your current schedules.")
            End If

            Return _studentManagementService.UpdateStudentSection(studentNumber,
                                                                  normalizedSectionName)
        End Function

        Private Function LoadContext(studentNumber As String) As ServiceResult(Of EnrollmentContext)
            If String.IsNullOrWhiteSpace(studentNumber) Then
                Return ServiceResult(Of EnrollmentContext).Failure("Student ID is required.")
            End If

            Dim studentResult As ServiceResult(Of StudentRecord) =
                _studentManagementService.GetStudentByStudentNumber(studentNumber.Trim())
            If Not studentResult.IsSuccess Then
                Return ServiceResult(Of EnrollmentContext).Failure(studentResult.Message)
            End If

            Dim subjectResult As ServiceResult(Of List(Of SubjectRecord)) =
                _subjectManagementService.GetSubjects()
            If Not subjectResult.IsSuccess Then
                Return ServiceResult(Of EnrollmentContext).Failure(subjectResult.Message)
            End If

            Return ServiceResult(Of EnrollmentContext).Success(New EnrollmentContext() With {
                .Student = studentResult.Data,
                .Subjects = If(subjectResult.Data, New List(Of SubjectRecord)())
            })
        End Function

        Private Function BuildSnapshot(context As EnrollmentContext) As StudentEnrollmentSnapshot
            Dim snapshot As New StudentEnrollmentSnapshot()
            If context Is Nothing OrElse context.Student Is Nothing Then
                snapshot.NoticeMessage = "No enrollment data loaded."
                Return snapshot
            End If

            snapshot.Student = context.Student

            Dim enrolledSubjectIds As New HashSet(Of Integer)()
            For Each subjectId As Integer In _enrollmentRepository.GetSubjectIdsByStudentId(context.Student.StudentRecordId)
                enrolledSubjectIds.Add(subjectId)
            Next

            For Each subject As SubjectRecord In SortSubjects(context.Subjects)
                If subject Is Nothing OrElse subject.SubjectId <= 0 Then
                    Continue For
                End If

                If enrolledSubjectIds.Contains(subject.SubjectId) Then
                    snapshot.SelectedSubjects.Add(subject)
                    Continue For
                End If

                If IsSubjectAvailableToStudent(context.Student, subject) Then
                    snapshot.AvailableSubjects.Add(subject)
                End If
            Next

            snapshot.AvailableSections = BuildAvailableSections(context.Student,
                                                                context.Subjects,
                                                                enrolledSubjectIds)
            snapshot.NoticeMessage = BuildNoticeMessage(snapshot)
            Return snapshot
        End Function

        Private Function SortSubjects(subjects As IEnumerable(Of SubjectRecord)) As List(Of SubjectRecord)
            Dim sortedSubjects As New List(Of SubjectRecord)()

            If subjects Is Nothing Then
                Return sortedSubjects
            End If

            For Each subject As SubjectRecord In subjects
                If subject IsNot Nothing Then
                    sortedSubjects.Add(subject)
                End If
            Next

            sortedSubjects.Sort(AddressOf CompareSubjects)
            Return sortedSubjects
        End Function

        Private Function CompareSubjects(left As SubjectRecord,
                                         right As SubjectRecord) As Integer
            Dim codeComparison As Integer =
                StringComparer.OrdinalIgnoreCase.Compare(NormalizeText(If(left Is Nothing,
                                                                          String.Empty,
                                                                          left.SubjectCode)),
                                                        NormalizeText(If(right Is Nothing,
                                                                          String.Empty,
                                                                          right.SubjectCode)))
            If codeComparison <> 0 Then
                Return codeComparison
            End If

            Return StringComparer.OrdinalIgnoreCase.Compare(NormalizeText(If(left Is Nothing,
                                                                             String.Empty,
                                                                             left.SubjectName)),
                                                           NormalizeText(If(right Is Nothing,
                                                                             String.Empty,
                                                                             right.SubjectName)))
        End Function

        Private Function FindSubjectByCode(subjects As IEnumerable(Of SubjectRecord),
                                           subjectCode As String) As SubjectRecord
            Dim normalizedCode As String = NormalizeText(subjectCode)
            If subjects Is Nothing OrElse String.IsNullOrWhiteSpace(normalizedCode) Then
                Return Nothing
            End If

            For Each subject As SubjectRecord In subjects
                If subject Is Nothing Then
                    Continue For
                End If

                If String.Equals(NormalizeText(subject.SubjectCode),
                                 normalizedCode,
                                 StringComparison.OrdinalIgnoreCase) Then
                    Return subject
                End If
            Next

            Return Nothing
        End Function

        Private Function BuildNoticeMessage(snapshot As StudentEnrollmentSnapshot) As String
            If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
                Return "No enrollment data loaded."
            End If

            If Not snapshot.Student.YearLevel.HasValue Then
                Return "Your year level is not set. Contact the registrar to unlock eligible subjects."
            End If

            Dim yearLevelValue As String =
                StudentScheduleHelper.BuildStudentYearLevelValue(snapshot.Student,
                                                                 "--")
            If snapshot.AvailableSubjectCount = 0 AndAlso snapshot.SelectedSubjectCount = 0 Then
                Return "No subjects are currently available for year level " &
                    yearLevelValue & "."
            End If

            If snapshot.AvailableSubjectCount = 0 Then
                Return "All eligible subjects for year level " &
                    yearLevelValue &
                    " are already in your selected load."
            End If

            Return "Choose from the subjects tagged for year level " &
                yearLevelValue &
                "."
        End Function

        Private Function IsSubjectAvailableToStudent(student As StudentRecord,
                                                     subject As SubjectRecord) As Boolean
            If student Is Nothing OrElse subject Is Nothing Then
                Return False
            End If

            Return MatchesStudentYearLevel(student, subject)
        End Function

        Private Function MatchesStudentYearLevel(student As StudentRecord,
                                                 subject As SubjectRecord) As Boolean
            If student Is Nothing OrElse
               subject Is Nothing OrElse
               Not student.YearLevel.HasValue Then
                Return False
            End If

            Dim subjectYearLevels As List(Of Integer) = ParseSubjectYearLevels(subject.YearLevel)
            If subjectYearLevels.Count = 0 Then
                Return False
            End If

            For Each subjectYear As Integer In subjectYearLevels
                If subjectYear = student.YearLevel.Value Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Function ParseSubjectYearLevels(value As String) As List(Of Integer)
            Dim levels As New List(Of Integer)()
            Dim normalizedValue As String = NormalizeText(value)
            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Return levels
            End If

            Dim separators As Char() = {","c, "/"c, ";"c, "&"c}
            Dim segments As String() =
                normalizedValue.Split(separators, StringSplitOptions.RemoveEmptyEntries)

            For Each segment As String In segments
                Dim parsedLevel As Integer? = ParseYearLevelValue(segment)
                If parsedLevel.HasValue AndAlso Not levels.Contains(parsedLevel.Value) Then
                    levels.Add(parsedLevel.Value)
                End If
            Next

            If levels.Count > 0 Then
                Return levels
            End If

            Dim singleLevel As Integer? = ParseYearLevelValue(normalizedValue)
            If singleLevel.HasValue Then
                levels.Add(singleLevel.Value)
            End If

            Return levels
        End Function

        Private Function ParseYearLevelValue(value As String) As Integer?
            Dim normalizedValue As String = NormalizeText(value).ToLowerInvariant()
            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Return Nothing
            End If

            Select Case normalizedValue
                Case "1", "1st", "1st year", "first year"
                    Return 1
                Case "2", "2nd", "2nd year", "second year"
                    Return 2
                Case "3", "3rd", "3rd year", "third year"
                    Return 3
                Case "4", "4th", "4th year", "fourth year"
                    Return 4
                Case "5", "5th", "5th year", "fifth year"
                    Return 5
                Case "6", "6th", "6th year", "sixth year"
                    Return 6
            End Select

            Dim digits As New StringBuilder()

            For Each currentCharacter As Char In normalizedValue
                If Char.IsDigit(currentCharacter) Then
                    digits.Append(currentCharacter)
                ElseIf digits.Length > 0 Then
                    Exit For
                End If
            Next

            Dim parsedNumericValue As Integer
            If Integer.TryParse(digits.ToString(), parsedNumericValue) Then
                Return parsedNumericValue
            End If

            Return Nothing
        End Function

        Private Function BuildAvailableSections(student As StudentRecord,
                                                subjects As IEnumerable(Of SubjectRecord),
                                                Optional enrolledSubjectIds As HashSet(Of Integer) = Nothing) As List(Of StudentSectionOption)
            Dim options As New List(Of StudentSectionOption)()
            If student Is Nothing Then
                Return options
            End If

            Dim resolvedEnrolledSubjectIds As HashSet(Of Integer) = enrolledSubjectIds
            If resolvedEnrolledSubjectIds Is Nothing Then
                resolvedEnrolledSubjectIds = New HashSet(Of Integer)()

                For Each subjectId As Integer In _enrollmentRepository.GetSubjectIdsByStudentId(student.StudentRecordId)
                    resolvedEnrolledSubjectIds.Add(subjectId)
                Next
            End If

            Dim referenceSubjects As List(Of SubjectRecord) =
                CollectSectionReferenceSubjects(student,
                                               subjects,
                                               resolvedEnrolledSubjectIds)
            If referenceSubjects.Count = 0 Then
                AddCurrentSectionOption(options, student)
                SortSectionOptions(options)
                Return options
            End If

            Dim scheduleResult As ServiceResult(Of List(Of TeacherScheduleRecord)) =
                _teacherScheduleManagementService.GetSchedules()
            If scheduleResult IsNot Nothing AndAlso scheduleResult.IsSuccess Then
                For Each schedule As TeacherScheduleRecord In scheduleResult.Data
                    If schedule Is Nothing OrElse
                       Not ScheduleMatchesReferenceSubjects(schedule, referenceSubjects) Then
                        Continue For
                    End If

                    AddSectionOption(options,
                                     NormalizeText(schedule.Section),
                                     student)
                Next
            End If

            AddCurrentSectionOption(options, student)
            SortSectionOptions(options)
            Return options
        End Function

        Private Function CollectSectionReferenceSubjects(student As StudentRecord,
                                                         subjects As IEnumerable(Of SubjectRecord),
                                                         enrolledSubjectIds As HashSet(Of Integer)) As List(Of SubjectRecord)
            Dim eligibleSubjects As New List(Of SubjectRecord)()
            Dim selectedSubjects As New List(Of SubjectRecord)()

            If student Is Nothing OrElse subjects Is Nothing Then
                Return eligibleSubjects
            End If

            For Each subject As SubjectRecord In subjects
                If subject Is Nothing OrElse subject.SubjectId <= 0 Then
                    Continue For
                End If

                If enrolledSubjectIds IsNot Nothing AndAlso
                   enrolledSubjectIds.Contains(subject.SubjectId) Then
                    selectedSubjects.Add(subject)
                    Continue For
                End If

                If IsSubjectAvailableToStudent(student, subject) Then
                    eligibleSubjects.Add(subject)
                End If
            Next

            If selectedSubjects.Count > 0 Then
                Return selectedSubjects
            End If

            Return eligibleSubjects
        End Function

        Private Function ScheduleMatchesReferenceSubjects(schedule As TeacherScheduleRecord,
                                                          referenceSubjects As IEnumerable(Of SubjectRecord)) As Boolean
            If schedule Is Nothing OrElse referenceSubjects Is Nothing Then
                Return False
            End If

            For Each subject As SubjectRecord In referenceSubjects
                If ScheduleMatchesSubject(schedule, subject) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Function ScheduleMatchesSubject(schedule As TeacherScheduleRecord,
                                                subject As SubjectRecord) As Boolean
            If schedule Is Nothing OrElse subject Is Nothing Then
                Return False
            End If

            Dim scheduleSubjectCode As String = NormalizeText(schedule.SubjectCode)
            Dim scheduleSubjectName As String = NormalizeText(schedule.SubjectName)
            Dim subjectCode As String = NormalizeText(subject.SubjectCode)
            Dim subjectName As String = NormalizeText(subject.SubjectName)

            If Not String.IsNullOrWhiteSpace(scheduleSubjectCode) AndAlso
               Not String.IsNullOrWhiteSpace(subjectCode) AndAlso
               String.Equals(scheduleSubjectCode,
                             subjectCode,
                             StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            If String.IsNullOrWhiteSpace(scheduleSubjectCode) AndAlso
               Not String.IsNullOrWhiteSpace(scheduleSubjectName) AndAlso
               Not String.IsNullOrWhiteSpace(subjectName) AndAlso
               String.Equals(scheduleSubjectName,
                             subjectName,
                             StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Return False
        End Function

        Private Sub AddCurrentSectionOption(options As List(Of StudentSectionOption),
                                            student As StudentRecord)
            If student Is Nothing Then
                Return
            End If

            AddSectionOption(options,
                             NormalizeText(student.SectionName),
                             student)
        End Sub

        Private Sub AddSectionOption(options As List(Of StudentSectionOption),
                                     sectionValue As String,
                                     student As StudentRecord)
            Dim normalizedSectionValue As String = NormalizeText(sectionValue)
            If String.IsNullOrWhiteSpace(normalizedSectionValue) Then
                Return
            End If

            If ContainsSectionOption(options, normalizedSectionValue, student) Then
                Return
            End If

            options.Add(New StudentSectionOption() With {
                .SectionValue = normalizedSectionValue,
                .SectionLabel = BuildSectionLabel(normalizedSectionValue, student),
                .ComparisonToken = BuildSectionComparisonToken(normalizedSectionValue,
                                                               student)
            })
        End Sub

        Private Function ContainsSectionOption(options As IEnumerable(Of StudentSectionOption),
                                               sectionValue As String,
                                               student As StudentRecord) As Boolean
            If options Is Nothing Then
                Return False
            End If

            Dim comparisonToken As String = BuildSectionComparisonToken(sectionValue, student)
            If String.IsNullOrWhiteSpace(comparisonToken) Then
                Return False
            End If

            For Each sectionOption As StudentSectionOption In options
                If sectionOption Is Nothing Then
                    Continue For
                End If

                If String.Equals(sectionOption.ComparisonToken,
                                 comparisonToken,
                                 StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Sub SortSectionOptions(options As List(Of StudentSectionOption))
            If options Is Nothing Then
                Return
            End If

            options.Sort(Function(left, right)
                             Return StringComparer.OrdinalIgnoreCase.Compare(
                                 NormalizeText(If(left Is Nothing,
                                                  String.Empty,
                                                  left.SectionLabel)),
                                 NormalizeText(If(right Is Nothing,
                                                  String.Empty,
                                                  right.SectionLabel)))
                         End Function)
        End Sub

        Private Function BuildSectionLabel(sectionValue As String,
                                           student As StudentRecord) As String
            Return StudentScheduleHelper.BuildCompactSectionValue(
                sectionValue,
                ResolveStudentYearToken(student))
        End Function

        Private Function BuildSectionComparisonToken(sectionValue As String,
                                                     student As StudentRecord) As String
            Dim normalizedSection As String = NormalizeSectionToken(sectionValue)
            If String.IsNullOrWhiteSpace(normalizedSection) Then
                Return String.Empty
            End If

            Dim yearToken As String = ResolveStudentYearToken(student)
            If String.IsNullOrWhiteSpace(yearToken) OrElse
               normalizedSection.StartsWith(yearToken,
                                            StringComparison.OrdinalIgnoreCase) Then
                Return normalizedSection.ToUpperInvariant()
            End If

            Return (yearToken & normalizedSection).ToUpperInvariant()
        End Function

        Private Function ResolveStudentYearToken(student As StudentRecord) As String
            If student Is Nothing OrElse Not student.YearLevel.HasValue Then
                Return String.Empty
            End If

            Return student.YearLevel.Value.ToString()
        End Function

        Private Function NormalizeSectionToken(sectionValue As String) As String
            Dim normalizedSection As String = NormalizeText(sectionValue)
            If String.IsNullOrWhiteSpace(normalizedSection) Then
                Return String.Empty
            End If

            If normalizedSection.StartsWith("Section:",
                                            StringComparison.OrdinalIgnoreCase) Then
                normalizedSection =
                    NormalizeText(normalizedSection.Substring("Section:".Length))
            ElseIf normalizedSection.StartsWith("Section ",
                                                StringComparison.OrdinalIgnoreCase) Then
                normalizedSection =
                    NormalizeText(normalizedSection.Substring("Section ".Length))
            End If

            Return normalizedSection.Replace(" ", String.Empty)
        End Function

        Private Function NormalizeText(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function

        Private Function BuildDatabaseErrorMessage(operationName As String,
                                                   ex As MySqlException) As String
            If ex Is Nothing Then
                Return "Unable to " & operationName & " enrollment records."
            End If

            If ex.Number = 1062 Then
                Return "That subject is already in your selected load."
            End If

            Return "Unable to " & operationName & " enrollment records." &
                Environment.NewLine &
                ex.Message
        End Function
    End Class
End Namespace
