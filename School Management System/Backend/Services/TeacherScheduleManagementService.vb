Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text.Json
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class TeacherScheduleManagementService
        Private Class LegacyScheduleStorageRecord
            Public Property TeacherId As String
            Public Property TeacherName As String
            Public Property Day As String
            Public Property Session As String
            Public Property Section As String
            Public Property SubjectCode As String
            Public Property SubjectName As String
            Public Property Room As String
        End Class

        Private ReadOnly _scheduleRepository As TeacherScheduleRepository
        Private ReadOnly _legacySchedulesStoragePath As String =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "SchoolManagementSystem",
                         "professor-schedules.json")
        Private ReadOnly _legacyJsonOptions As New JsonSerializerOptions() With {
            .PropertyNameCaseInsensitive = True
        }
        Private _legacyImportChecked As Boolean

        Public Sub New()
            Me.New(New TeacherScheduleRepository())
        End Sub

        Public Sub New(scheduleRepository As TeacherScheduleRepository)
            _scheduleRepository = scheduleRepository
        End Sub

        Public Function GetSchedules() As ServiceResult(Of List(Of TeacherScheduleRecord))
            Try
                EnsureLegacySchedulesImported()

                Dim schedules As List(Of TeacherScheduleRecord) = _scheduleRepository.GetAll()
                Return ServiceResult(Of List(Of TeacherScheduleRecord)).Success(schedules)
            Catch ex As MySqlException
                Return ServiceResult(Of List(Of TeacherScheduleRecord)).Failure(
                    BuildDatabaseErrorMessage("load", ex))
            Catch ex As Exception
                Return ServiceResult(Of List(Of TeacherScheduleRecord)).Failure(
                    "Unable to load schedule records." &
                    Environment.NewLine &
                    ex.Message)
            End Try
        End Function

        Public Function SaveSchedule(request As TeacherScheduleSaveRequest) As ServiceResult(Of TeacherScheduleRecord)
            Dim normalizedRequest As TeacherScheduleSaveRequest = NormalizeRequest(request)
            Dim validationMessage As String = ValidateRequest(normalizedRequest)
            If Not String.IsNullOrWhiteSpace(validationMessage) Then
                Return ServiceResult(Of TeacherScheduleRecord).Failure(validationMessage)
            End If

            Try
                EnsureLegacySchedulesImported()

                Dim conflictMessage As String = GetConflictMessage(normalizedRequest)
                If Not String.IsNullOrWhiteSpace(conflictMessage) Then
                    Return ServiceResult(Of TeacherScheduleRecord).Failure(conflictMessage)
                End If

                Dim savedRecord As TeacherScheduleRecord = _scheduleRepository.Save(normalizedRequest)
                If savedRecord Is Nothing Then
                    Return ServiceResult(Of TeacherScheduleRecord).Failure(
                        "The selected professor no longer exists.")
                End If

                Return ServiceResult(Of TeacherScheduleRecord).Success(savedRecord,
                                                                       "Schedule slot saved.")
            Catch ex As MySqlException
                Return ServiceResult(Of TeacherScheduleRecord).Failure(
                    BuildDatabaseErrorMessage("save", ex))
            End Try
        End Function

        Public Function DeleteSchedule(teacherId As String,
                                       dayValue As String,
                                       sessionValue As String) As ServiceResult(Of Boolean)
            Dim normalizedTeacherId As String = NormalizeText(teacherId)
            Dim normalizedDay As String = NormalizeDay(dayValue)
            Dim normalizedSession As String = NormalizeSession(sessionValue)

            If String.IsNullOrWhiteSpace(normalizedTeacherId) Then
                Return ServiceResult(Of Boolean).Failure("Select a professor first.")
            End If

            If String.IsNullOrWhiteSpace(normalizedDay) Then
                Return ServiceResult(Of Boolean).Failure("Select a valid day.")
            End If

            If String.IsNullOrWhiteSpace(normalizedSession) Then
                Return ServiceResult(Of Boolean).Failure("Select a valid session.")
            End If

            Try
                EnsureLegacySchedulesImported()

                Dim deleted As Boolean =
                    _scheduleRepository.DeleteByTeacherSlot(normalizedTeacherId,
                                                            normalizedDay,
                                                            normalizedSession)
                If Not deleted Then
                    Return ServiceResult(Of Boolean).Failure(
                        "No matching slot found for this professor.")
                End If

                Return ServiceResult(Of Boolean).Success(True, "Schedule slot removed.")
            Catch ex As MySqlException
                Return ServiceResult(Of Boolean).Failure(
                    BuildDatabaseErrorMessage("delete", ex))
            End Try
        End Function

        Private Function ValidateRequest(request As TeacherScheduleSaveRequest) As String
            If request Is Nothing Then
                Return "Schedule data is required."
            End If

            If String.IsNullOrWhiteSpace(request.TeacherId) Then
                Return "Select a professor first."
            End If

            If String.IsNullOrWhiteSpace(request.Day) Then
                Return "Select a valid day."
            End If

            If String.IsNullOrWhiteSpace(request.Session) Then
                Return "Session is required in 24-hour format."
            End If

            Dim startMinutes As Integer
            Dim endMinutes As Integer
            If Not TryParseSessionRange(request.Session, startMinutes, endMinutes) Then
                Return "Session is required in 24-hour format."
            End If

            If String.IsNullOrWhiteSpace(request.SubjectCode) AndAlso
               String.IsNullOrWhiteSpace(request.SubjectName) Then
                Return "Select a subject before saving a slot."
            End If

            Return String.Empty
        End Function

        Private Function NormalizeRequest(request As TeacherScheduleSaveRequest) As TeacherScheduleSaveRequest
            If request Is Nothing Then
                Return Nothing
            End If

            Return New TeacherScheduleSaveRequest() With {
                .TeacherId = NormalizeText(request.TeacherId),
                .Day = NormalizeDay(request.Day),
                .Session = NormalizeSession(request.Session),
                .Section = NormalizeText(request.Section),
                .SubjectCode = NormalizeText(request.SubjectCode),
                .SubjectName = NormalizeText(request.SubjectName),
                .Room = NormalizeText(request.Room)
            }
        End Function

        Private Sub EnsureLegacySchedulesImported()
            If _legacyImportChecked Then
                Return
            End If

            _legacyImportChecked = True

            If Not File.Exists(_legacySchedulesStoragePath) Then
                Return
            End If

            Dim requests As List(Of TeacherScheduleSaveRequest) = ReadLegacyScheduleRequests()
            If requests.Count = 0 Then
                Return
            End If

            Dim skippedCount As Integer = 0
            Dim importedCount As Integer = 0

            For Each request As TeacherScheduleSaveRequest In requests
                Dim validationMessage As String = ValidateRequest(request)
                If Not String.IsNullOrWhiteSpace(validationMessage) Then
                    skippedCount += 1
                    Continue For
                End If

                Dim conflictMessage As String = GetConflictMessage(request)
                If Not String.IsNullOrWhiteSpace(conflictMessage) Then
                    skippedCount += 1
                    Continue For
                End If

                Dim savedRecord As TeacherScheduleRecord = _scheduleRepository.Save(request)
                If savedRecord Is Nothing Then
                    skippedCount += 1
                    Continue For
                End If

                importedCount += 1
            Next

            If importedCount > 0 AndAlso skippedCount = 0 Then
                ArchiveLegacySchedulesStorage()
            End If
        End Sub

        Private Function ReadLegacyScheduleRequests() As List(Of TeacherScheduleSaveRequest)
            Dim requests As New List(Of TeacherScheduleSaveRequest)()

            Try
                Dim json As String = File.ReadAllText(_legacySchedulesStoragePath)
                If String.IsNullOrWhiteSpace(json) Then
                    Return requests
                End If

                Dim records As List(Of LegacyScheduleStorageRecord) =
                    JsonSerializer.Deserialize(Of List(Of LegacyScheduleStorageRecord))(json,
                                                                                       _legacyJsonOptions)
                If records Is Nothing Then
                    Return requests
                End If

                For Each record As LegacyScheduleStorageRecord In records
                    Dim normalizedRequest As TeacherScheduleSaveRequest =
                        NormalizeRequest(New TeacherScheduleSaveRequest() With {
                            .TeacherId = If(If(record Is Nothing, Nothing, record.TeacherId), String.Empty),
                            .Day = If(If(record Is Nothing, Nothing, record.Day), String.Empty),
                            .Session = If(If(record Is Nothing, Nothing, record.Session), String.Empty),
                            .Section = If(If(record Is Nothing, Nothing, record.Section), String.Empty),
                            .SubjectCode = If(If(record Is Nothing, Nothing, record.SubjectCode), String.Empty),
                            .SubjectName = If(If(record Is Nothing, Nothing, record.SubjectName), String.Empty),
                            .Room = If(If(record Is Nothing, Nothing, record.Room), String.Empty)
                        })

                    If Not String.IsNullOrWhiteSpace(ValidateRequest(normalizedRequest)) Then
                        Continue For
                    End If

                    requests.Add(normalizedRequest)
                Next
            Catch
                Return New List(Of TeacherScheduleSaveRequest)()
            End Try

            Return requests
        End Function

        Private Sub ArchiveLegacySchedulesStorage()
            If Not File.Exists(_legacySchedulesStoragePath) Then
                Return
            End If

            Try
                Dim backupPath As String = _legacySchedulesStoragePath & ".migrated"
                If File.Exists(backupPath) Then
                    backupPath = _legacySchedulesStoragePath & "." &
                        DateTime.UtcNow.ToString("yyyyMMddHHmmss") &
                        ".migrated"
                End If

                File.Move(_legacySchedulesStoragePath, backupPath)
            Catch
            End Try
        End Sub

        Private Function BuildDatabaseErrorMessage(operationName As String,
                                                   ex As MySqlException) As String
            If ex Is Nothing Then
                Return "Unable to " & operationName & " schedule records."
            End If

            If ex.Number = 1452 Then
                Return "The selected professor no longer exists."
            End If

            Return "Unable to " & operationName & " schedule records." &
                Environment.NewLine &
                ex.Message
        End Function

        Private Function GetConflictMessage(request As TeacherScheduleSaveRequest) As String
            If request Is Nothing Then
                Return String.Empty
            End If

            Dim schedules As List(Of TeacherScheduleRecord) = _scheduleRepository.GetAll()
            Dim teacherConflict As TeacherScheduleRecord = FindTeacherConflict(request, schedules)
            If teacherConflict IsNot Nothing Then
                Return "Time conflict: " &
                    GetTeacherReference(teacherConflict) &
                    " already has " &
                    GetSubjectReference(teacherConflict) &
                    " on " &
                    request.Day &
                    " at " &
                    NormalizeSession(teacherConflict.Session) &
                    "."
            End If

            Dim roomConflict As TeacherScheduleRecord = FindRoomConflict(request, schedules)
            If roomConflict IsNot Nothing Then
                Return "Time conflict: room " &
                    request.Room &
                    " is already assigned to " &
                    GetTeacherReference(roomConflict) &
                    " for " &
                    GetSubjectReference(roomConflict) &
                    " on " &
                    request.Day &
                    " at " &
                    NormalizeSession(roomConflict.Session) &
                    "."
            End If

            Dim sectionConflict As TeacherScheduleRecord = FindSectionConflict(request, schedules)
            If sectionConflict IsNot Nothing Then
                Return "Time conflict: section " &
                    request.Section &
                    " already has " &
                    GetSubjectReference(sectionConflict) &
                    " with " &
                    GetTeacherReference(sectionConflict) &
                    " on " &
                    request.Day &
                    " at " &
                    NormalizeSession(sectionConflict.Session) &
                    "."
            End If

            Return String.Empty
        End Function

        Private Function FindTeacherConflict(request As TeacherScheduleSaveRequest,
                                             schedules As IEnumerable(Of TeacherScheduleRecord)) As TeacherScheduleRecord
            If request Is Nothing OrElse schedules Is Nothing Then
                Return Nothing
            End If

            For Each schedule As TeacherScheduleRecord In schedules
                If schedule Is Nothing OrElse ShouldSkipConflictCheck(schedule, request) Then
                    Continue For
                End If

                If Not String.Equals(NormalizeText(schedule.TeacherId),
                                     request.TeacherId,
                                     StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                If Not IsSameNormalizedDay(schedule.Day, request.Day) Then
                    Continue For
                End If

                If SessionRangesOverlap(schedule.Session, request.Session) Then
                    Return schedule
                End If
            Next

            Return Nothing
        End Function

        Private Function FindRoomConflict(request As TeacherScheduleSaveRequest,
                                          schedules As IEnumerable(Of TeacherScheduleRecord)) As TeacherScheduleRecord
            If request Is Nothing OrElse
               schedules Is Nothing OrElse
               String.IsNullOrWhiteSpace(request.Room) Then
                Return Nothing
            End If

            For Each schedule As TeacherScheduleRecord In schedules
                If schedule Is Nothing OrElse ShouldSkipConflictCheck(schedule, request) Then
                    Continue For
                End If

                If Not IsSameNormalizedDay(schedule.Day, request.Day) Then
                    Continue For
                End If

                If Not String.Equals(NormalizeText(schedule.Room),
                                     request.Room,
                                     StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                If SessionRangesOverlap(schedule.Session, request.Session) Then
                    Return schedule
                End If
            Next

            Return Nothing
        End Function

        Private Function FindSectionConflict(request As TeacherScheduleSaveRequest,
                                             schedules As IEnumerable(Of TeacherScheduleRecord)) As TeacherScheduleRecord
            If request Is Nothing OrElse
               schedules Is Nothing OrElse
               String.IsNullOrWhiteSpace(request.Section) Then
                Return Nothing
            End If

            For Each schedule As TeacherScheduleRecord In schedules
                If schedule Is Nothing OrElse ShouldSkipConflictCheck(schedule, request) Then
                    Continue For
                End If

                If Not IsSameNormalizedDay(schedule.Day, request.Day) Then
                    Continue For
                End If

                If Not String.Equals(NormalizeText(schedule.Section),
                                     request.Section,
                                     StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                If SessionRangesOverlap(schedule.Session, request.Session) Then
                    Return schedule
                End If
            Next

            Return Nothing
        End Function

        Private Function ShouldSkipConflictCheck(schedule As TeacherScheduleRecord,
                                                 request As TeacherScheduleSaveRequest) As Boolean
            If schedule Is Nothing OrElse request Is Nothing Then
                Return True
            End If

            Return String.Equals(NormalizeText(schedule.TeacherId),
                                 request.TeacherId,
                                 StringComparison.OrdinalIgnoreCase) AndAlso
                   IsSameNormalizedDay(schedule.Day, request.Day) AndAlso
                   String.Equals(NormalizeSession(schedule.Session),
                                 request.Session,
                                 StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function IsSameNormalizedDay(leftDay As String, rightDay As String) As Boolean
            Return String.Equals(NormalizeDay(leftDay),
                                 NormalizeDay(rightDay),
                                 StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function SessionRangesOverlap(leftSession As String, rightSession As String) As Boolean
            Dim leftStart As Integer
            Dim leftEnd As Integer
            Dim rightStart As Integer
            Dim rightEnd As Integer
            Dim leftParsed As Boolean = TryParseSessionRange(leftSession, leftStart, leftEnd)
            Dim rightParsed As Boolean = TryParseSessionRange(rightSession, rightStart, rightEnd)

            If leftParsed AndAlso rightParsed Then
                Return leftStart < rightEnd AndAlso rightStart < leftEnd
            End If

            Return String.Equals(NormalizeSession(leftSession),
                                 NormalizeSession(rightSession),
                                 StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function GetTeacherReference(schedule As TeacherScheduleRecord) As String
            If schedule Is Nothing Then
                Return "this professor"
            End If

            Dim teacherName As String = NormalizeText(schedule.TeacherName)
            If Not String.IsNullOrWhiteSpace(teacherName) Then
                Return teacherName
            End If

            Dim teacherId As String = NormalizeText(schedule.TeacherId)
            If Not String.IsNullOrWhiteSpace(teacherId) Then
                Return teacherId
            End If

            Return "this professor"
        End Function

        Private Function GetSubjectReference(schedule As TeacherScheduleRecord) As String
            If schedule Is Nothing Then
                Return "another class"
            End If

            Dim subjectCode As String = NormalizeText(schedule.SubjectCode)
            Dim subjectName As String = NormalizeText(schedule.SubjectName)
            If Not String.IsNullOrWhiteSpace(subjectCode) AndAlso
               Not String.IsNullOrWhiteSpace(subjectName) AndAlso
               Not String.Equals(subjectCode, subjectName, StringComparison.OrdinalIgnoreCase) Then
                Return subjectCode & " - " & subjectName
            End If

            If Not String.IsNullOrWhiteSpace(subjectCode) Then
                Return subjectCode
            End If

            If Not String.IsNullOrWhiteSpace(subjectName) Then
                Return subjectName
            End If

            Return "another class"
        End Function

        Private Function NormalizeDay(dayValue As String) As String
            Dim normalized As String = NormalizeText(dayValue)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return String.Empty
            End If

            Select Case normalized.ToLowerInvariant()
                Case "monday", "mon"
                    Return "Monday"
                Case "tuesday", "tue", "tues"
                    Return "Tuesday"
                Case "wednesday", "wed"
                    Return "Wednesday"
                Case "thursday", "thu", "thur", "thurs"
                    Return "Thursday"
                Case "friday", "fri"
                    Return "Friday"
                Case Else
                    Return String.Empty
            End Select
        End Function

        Private Function NormalizeSession(sessionValue As String) As String
            Dim startMinutes As Integer
            Dim endMinutes As Integer
            If TryParseSessionRange(sessionValue, startMinutes, endMinutes) Then
                Return FormatSessionRange(startMinutes, endMinutes)
            End If

            Return NormalizeText(sessionValue)
        End Function

        Private Function TryParseSessionRange(sessionValue As String,
                                              ByRef parsedStartMinutes As Integer,
                                              ByRef parsedEndMinutes As Integer) As Boolean
            Dim normalizedSession As String = NormalizeText(sessionValue)
            If String.IsNullOrWhiteSpace(normalizedSession) Then
                Return False
            End If

            Dim startToken As String = String.Empty
            Dim endToken As String = String.Empty
            If Not TrySplitSessionRange(normalizedSession, startToken, endToken) Then
                Return False
            End If

            If Not TryParseClockToken(startToken, parsedStartMinutes) OrElse
               Not TryParseClockToken(endToken, parsedEndMinutes) Then
                Return False
            End If

            If parsedEndMinutes <= parsedStartMinutes Then
                Return False
            End If

            Return True
        End Function

        Private Function TrySplitSessionRange(sessionValue As String,
                                              ByRef startToken As String,
                                              ByRef endToken As String) As Boolean
            startToken = String.Empty
            endToken = String.Empty

            Dim separator As String = " - "
            Dim separatorIndex As Integer = sessionValue.IndexOf(separator, StringComparison.Ordinal)
            If separatorIndex <= 0 Then
                separator = " to "
                separatorIndex = sessionValue.IndexOf(separator, StringComparison.OrdinalIgnoreCase)
            End If

            If separatorIndex <= 0 Then
                separator = "-"
                separatorIndex = sessionValue.IndexOf(separator, StringComparison.Ordinal)
            End If

            If separatorIndex <= 0 Then
                Return False
            End If

            startToken = NormalizeText(sessionValue.Substring(0, separatorIndex))
            endToken = NormalizeText(sessionValue.Substring(separatorIndex + separator.Length))

            If String.IsNullOrWhiteSpace(startToken) OrElse
               String.IsNullOrWhiteSpace(endToken) Then
                Return False
            End If

            Return True
        End Function

        Private Function TryParseClockToken(value As String,
                                            ByRef parsedMinutes As Integer) As Boolean
            parsedMinutes = 0

            Dim normalized As String = NormalizeText(value)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            Dim splitTokens As String() = normalized.Split(":"c)
            If splitTokens.Length = 2 Then
                Dim parsedHour As Integer
                Dim parsedMinute As Integer

                If Integer.TryParse(NormalizeText(splitTokens(0)),
                                    NumberStyles.Integer,
                                    CultureInfo.InvariantCulture,
                                    parsedHour) AndAlso
                   Integer.TryParse(NormalizeText(splitTokens(1)),
                                    NumberStyles.Integer,
                                    CultureInfo.InvariantCulture,
                                    parsedMinute) Then
                    If parsedHour >= 0 AndAlso parsedHour <= 24 AndAlso
                       parsedMinute >= 0 AndAlso parsedMinute <= 59 AndAlso
                       (parsedHour < 24 OrElse parsedMinute = 0) Then
                        parsedMinutes = (parsedHour * 60) + parsedMinute
                        Return True
                    End If
                End If
            End If

            Dim supportedFormats As String() = New String() {
                "HH:mm",
                "H:mm",
                "hh:mm tt",
                "h:mm tt",
                "hh:mmtt",
                "h:mmtt"
            }

            Dim parsedTime As DateTime
            If DateTime.TryParseExact(normalized,
                                      supportedFormats,
                                      CultureInfo.InvariantCulture,
                                      DateTimeStyles.AllowWhiteSpaces,
                                      parsedTime) Then
                parsedMinutes = (parsedTime.Hour * 60) + parsedTime.Minute
                Return True
            End If

            If DateTime.TryParse(normalized,
                                 CultureInfo.CurrentCulture,
                                 DateTimeStyles.NoCurrentDateDefault,
                                 parsedTime) Then
                parsedMinutes = (parsedTime.Hour * 60) + parsedTime.Minute
                Return True
            End If

            If DateTime.TryParse(normalized,
                                 CultureInfo.InvariantCulture,
                                 DateTimeStyles.NoCurrentDateDefault,
                                 parsedTime) Then
                parsedMinutes = (parsedTime.Hour * 60) + parsedTime.Minute
                Return True
            End If

            Return False
        End Function

        Private Function FormatSessionRange(startMinutes As Integer,
                                            endMinutes As Integer) As String
            Return FormatClockMinutes(startMinutes) &
                   " - " &
                   FormatClockMinutes(endMinutes)
        End Function

        Private Function FormatClockMinutes(totalMinutes As Integer) As String
            Dim normalizedTotal As Integer = Math.Max(0, Math.Min(1440, totalMinutes))
            Dim hourValue As Integer = normalizedTotal \ 60
            Dim minuteValue As Integer = normalizedTotal Mod 60

            Return hourValue.ToString("00", CultureInfo.InvariantCulture) &
                   ":" &
                   minuteValue.ToString("00", CultureInfo.InvariantCulture)
        End Function

        Private Function NormalizeText(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function
    End Class
End Namespace
