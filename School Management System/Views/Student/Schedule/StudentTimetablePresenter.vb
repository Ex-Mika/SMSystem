Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models

Friend NotInheritable Class StudentTimetablePresenter
    Public Shared ReadOnly DayHeaders As String() = New String() {
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday"
    }

    Private Const TimetableSessionColumnName As String = "__Session"

    Private Sub New()
    End Sub

    Public Shared Function CreateEmptyTable() As DataTable
        Dim table As DataTable = CreateTimetableStructure()
        AddPlaceholderTimetableRow(table)
        Return table
    End Function

    Public Shared Function BuildTable(schedules As IEnumerable(Of TeacherScheduleRecord),
                                      subjectRecords As IEnumerable(Of SubjectRecord)) As DataTable
        If schedules Is Nothing Then
            Return CreateEmptyTable()
        End If

        Dim orderedSessions As List(Of String) = GetOrderedSessions(schedules)
        If orderedSessions.Count = 0 Then
            Return CreateEmptyTable()
        End If

        Dim table As DataTable = CreateTimetableStructure()
        For Each sessionEntry As String In orderedSessions
            AddTimetableRow(table, sessionEntry)
        Next

        For Each schedule As TeacherScheduleRecord In schedules
            Dim normalizedDay As String = NormalizeDayLabel(If(schedule Is Nothing, String.Empty, schedule.Day))
            Dim normalizedSession As String = NormalizeSession(If(schedule Is Nothing, String.Empty, schedule.Session))
            If String.IsNullOrWhiteSpace(normalizedDay) OrElse String.IsNullOrWhiteSpace(normalizedSession) Then
                Continue For
            End If

            Dim targetRow As DataRow = FindSessionRow(table, normalizedSession)
            If targetRow Is Nothing Then
                AddTimetableRow(table, normalizedSession)
                targetRow = FindSessionRow(table, normalizedSession)
            End If

            If targetRow IsNot Nothing Then
                targetRow(normalizedDay) = BuildTimetableCellDisplay(schedule, subjectRecords)
            End If
        Next

        If table.Rows.Count = 0 Then
            AddPlaceholderTimetableRow(table)
        End If

        Return table
    End Function

    Public Shared Sub ConfigureDataGrid(targetGrid As DataGrid, sourceTable As DataTable)
        If targetGrid Is Nothing OrElse sourceTable Is Nothing Then
            Return
        End If

        targetGrid.Columns.Clear()

        Dim centeredTextStyle As New Style(GetType(TextBlock))
        centeredTextStyle.Setters.Add(New Setter(TextBlock.TextAlignmentProperty, TextAlignment.Center))
        centeredTextStyle.Setters.Add(New Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center))
        centeredTextStyle.Setters.Add(New Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap))

        For Each dayHeader As String In DayHeaders
            If Not sourceTable.Columns.Contains(dayHeader) Then
                Continue For
            End If

            Dim dayColumn As New DataGridTextColumn() With {
                .Header = dayHeader,
                .Binding = New Binding(BuildDataTableBindingPath(dayHeader)),
                .IsReadOnly = True,
                .CanUserSort = False,
                .MinWidth = 120,
                .Width = New DataGridLength(1, DataGridLengthUnitType.Star),
                .ElementStyle = centeredTextStyle
            }
            targetGrid.Columns.Add(dayColumn)
        Next
    End Sub

    Public Shared Function MatchesStudentSchedule(record As TeacherScheduleRecord,
                                                  student As StudentRecord,
                                                  subjectRecords As IEnumerable(Of SubjectRecord)) As Boolean
        Return StudentScheduleHelper.MatchesStudentSchedule(record,
                                                            student,
                                                            subjectRecords)
    End Function

    Public Shared Function BuildStudentSectionLabel(student As StudentRecord) As String
        Return StudentScheduleHelper.BuildStudentSectionLabel(student)
    End Function

    Public Shared Function NormalizeDayLabel(dayValue As String) As String
        Return StudentScheduleHelper.NormalizeDayLabel(dayValue)
    End Function

    Private Shared Function CreateTimetableStructure() As DataTable
        Dim table As New DataTable()
        table.Columns.Add(TimetableSessionColumnName, GetType(String))

        For Each dayHeader As String In DayHeaders
            table.Columns.Add(dayHeader, GetType(String))
        Next

        Return table
    End Function

    Private Shared Sub AddPlaceholderTimetableRow(table As DataTable)
        If table Is Nothing Then
            Return
        End If

        AddTimetableRow(table, "--")
    End Sub

    Private Shared Sub AddTimetableRow(table As DataTable, sessionValue As String)
        If table Is Nothing Then
            Return
        End If

        Dim normalizedSession As String = NormalizeSession(sessionValue)
        If String.IsNullOrWhiteSpace(normalizedSession) Then
            normalizedSession = "--"
        End If

        Dim row As DataRow = table.NewRow()
        row(TimetableSessionColumnName) = normalizedSession

        For Each dayHeader As String In DayHeaders
            row(dayHeader) = "--"
        Next

        table.Rows.Add(row)
    End Sub

    Private Shared Function FindSessionRow(table As DataTable, sessionValue As String) As DataRow
        If table Is Nothing OrElse String.IsNullOrWhiteSpace(sessionValue) Then
            Return Nothing
        End If

        Dim normalizedSession As String = NormalizeSession(sessionValue)
        For Each row As DataRow In table.Rows
            Dim candidateSession As String = NormalizeSession(ReadRowValue(row, TimetableSessionColumnName))
            If String.Equals(candidateSession, normalizedSession, StringComparison.OrdinalIgnoreCase) Then
                Return row
            End If
        Next

        Return Nothing
    End Function

    Private Shared Function GetOrderedSessions(schedules As IEnumerable(Of TeacherScheduleRecord)) As List(Of String)
        Dim orderedSessions As New List(Of String)()
        If schedules Is Nothing Then
            Return orderedSessions
        End If

        For Each schedule As TeacherScheduleRecord In schedules
            Dim normalizedSession As String = NormalizeSession(If(schedule Is Nothing, String.Empty, schedule.Session))
            If String.IsNullOrWhiteSpace(normalizedSession) Then
                Continue For
            End If

            If Not ContainsIgnoreCase(orderedSessions, normalizedSession) Then
                orderedSessions.Add(normalizedSession)
            End If
        Next

        orderedSessions.Sort(AddressOf CompareSessions)
        Return orderedSessions
    End Function

    Private Shared Function CompareSessions(left As String, right As String) As Integer
        Dim leftStart As Integer
        Dim rightStart As Integer
        Dim leftHasTime As Boolean = TryParseSessionStart(left, leftStart)
        Dim rightHasTime As Boolean = TryParseSessionStart(right, rightStart)

        If leftHasTime AndAlso rightHasTime Then
            Return leftStart.CompareTo(rightStart)
        End If

        If leftHasTime Then
            Return -1
        End If

        If rightHasTime Then
            Return 1
        End If

        Return String.Compare(left, right, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function TryParseSessionStart(sessionValue As String, ByRef parsedStart As Integer) As Boolean
        Dim parsedEnd As Integer
        If TryParseSessionRange(sessionValue, parsedStart, parsedEnd) Then
            Return True
        End If

        Return False
    End Function

    Private Shared Function BuildTimetableCellDisplay(record As TeacherScheduleRecord,
                                                      subjectRecords As IEnumerable(Of SubjectRecord)) As String
        If record Is Nothing Then
            Return "--"
        End If

        Dim sessionValue As String = NormalizeSession(record.Session)
        Dim subjectCode As String = NormalizeText(record.SubjectCode)
        Dim subjectName As String = NormalizeText(record.SubjectName)
        Dim sectionValue As String = NormalizeText(record.Section)
        Dim yearLevelValue As String = ResolveSubjectYearLevel(record, subjectRecords)
        Dim sectionLine As String = BuildCompactSectionLabel(sectionValue, yearLevelValue)
        Dim roomValue As String = NormalizeText(record.Room)

        Dim subjectValue As String = subjectCode
        If Not String.IsNullOrWhiteSpace(subjectCode) AndAlso
           Not String.IsNullOrWhiteSpace(subjectName) AndAlso
           Not String.Equals(subjectCode, subjectName, StringComparison.OrdinalIgnoreCase) Then
            subjectValue = subjectCode & " - " & subjectName
        ElseIf String.IsNullOrWhiteSpace(subjectValue) Then
            subjectValue = subjectName
        End If

        If String.IsNullOrWhiteSpace(sessionValue) Then
            sessionValue = "--"
        End If
        If String.IsNullOrWhiteSpace(subjectValue) Then
            subjectValue = "--"
        End If
        If String.IsNullOrWhiteSpace(roomValue) Then
            roomValue = "--"
        End If

        Return String.Join(Environment.NewLine, New String() {
            sessionValue,
            subjectValue,
            sectionLine,
            roomValue
        })
    End Function

    Private Shared Function ResolveSubjectYearLevel(record As TeacherScheduleRecord,
                                                    subjectRecords As IEnumerable(Of SubjectRecord)) As String
        If record Is Nothing OrElse subjectRecords Is Nothing Then
            Return String.Empty
        End If

        Dim subjectCode As String = NormalizeText(record.SubjectCode)
        Dim subjectName As String = NormalizeText(record.SubjectName)

        For Each subjectRecord As SubjectRecord In subjectRecords
            If subjectRecord Is Nothing Then
                Continue For
            End If

            If Not String.IsNullOrWhiteSpace(subjectCode) AndAlso
               String.Equals(NormalizeText(subjectRecord.SubjectCode),
                             subjectCode,
                             StringComparison.OrdinalIgnoreCase) Then
                Return NormalizeText(subjectRecord.YearLevel)
            End If

            If String.IsNullOrWhiteSpace(subjectCode) AndAlso
               Not String.IsNullOrWhiteSpace(subjectName) AndAlso
               String.Equals(NormalizeText(subjectRecord.SubjectName),
                             subjectName,
                             StringComparison.OrdinalIgnoreCase) Then
                Return NormalizeText(subjectRecord.YearLevel)
            End If
        Next

        Return String.Empty
    End Function

    Private Shared Function ResolveStudentYearLevel(student As StudentRecord) As String
        If student Is Nothing Then
            Return String.Empty
        End If

        If student.YearLevel.HasValue Then
            Return student.YearLevel.Value.ToString(CultureInfo.InvariantCulture)
        End If

        Return NormalizeText(student.YearLevelLabel)
    End Function

    Private Shared Function BuildCompactSectionLabel(sectionValue As String, yearLevelValue As String) As String
        Return StudentScheduleHelper.BuildCompactSectionValue(sectionValue,
                                                              yearLevelValue)
    End Function

    Private Shared Function NormalizeSectionToken(sectionValue As String) As String
        Dim normalizedSection As String = NormalizeText(sectionValue)
        If String.IsNullOrWhiteSpace(normalizedSection) Then
            Return String.Empty
        End If

        If normalizedSection.StartsWith("Section:", StringComparison.OrdinalIgnoreCase) Then
            normalizedSection = NormalizeText(normalizedSection.Substring("Section:".Length))
        ElseIf normalizedSection.StartsWith("Section ", StringComparison.OrdinalIgnoreCase) Then
            normalizedSection = NormalizeText(normalizedSection.Substring("Section ".Length))
        End If

        Return normalizedSection.Replace(" ", String.Empty)
    End Function

    Private Shared Function NormalizeComparisonToken(sectionValue As String) As String
        Dim compactSection As String = NormalizeSectionToken(sectionValue)
        If String.IsNullOrWhiteSpace(compactSection) Then
            Return String.Empty
        End If

        Dim builder As New System.Text.StringBuilder()
        For Each characterValue As Char In compactSection
            If Char.IsLetterOrDigit(characterValue) Then
                builder.Append(Char.ToUpperInvariant(characterValue))
            End If
        Next

        Return builder.ToString()
    End Function

    Private Shared Function NormalizeYearLevelToken(yearLevelValue As String) As String
        Dim normalizedYearLevel As String = NormalizeText(yearLevelValue)
        If String.IsNullOrWhiteSpace(normalizedYearLevel) Then
            Return String.Empty
        End If

        Dim compactDigits As String = String.Empty
        For Each characterValue As Char In normalizedYearLevel
            If Char.IsDigit(characterValue) Then
                compactDigits &= characterValue
            ElseIf compactDigits.Length > 0 Then
                Exit For
            End If
        Next

        If compactDigits.Length > 0 Then
            Return compactDigits
        End If

        Dim lowerYearLevel As String = normalizedYearLevel.ToLowerInvariant()
        If lowerYearLevel.Contains("first") Then
            Return "1"
        End If
        If lowerYearLevel.Contains("second") Then
            Return "2"
        End If
        If lowerYearLevel.Contains("third") Then
            Return "3"
        End If
        If lowerYearLevel.Contains("fourth") Then
            Return "4"
        End If
        If lowerYearLevel.Contains("fifth") Then
            Return "5"
        End If
        If lowerYearLevel.Contains("sixth") Then
            Return "6"
        End If

        Return normalizedYearLevel.Replace(" ", String.Empty)
    End Function

    Private Shared Function NormalizeSession(sessionValue As String) As String
        Return StudentScheduleHelper.NormalizeSession(sessionValue)
    End Function

    Private Shared Function TryParseSessionRange(sessionValue As String,
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

    Private Shared Function TrySplitSessionRange(sessionValue As String,
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
        If String.IsNullOrWhiteSpace(startToken) OrElse String.IsNullOrWhiteSpace(endToken) Then
            Return False
        End If

        Return True
    End Function

    Private Shared Function TryParseClockToken(value As String, ByRef parsedMinutes As Integer) As Boolean
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

    Private Shared Function FormatSessionRange(startMinutes As Integer, endMinutes As Integer) As String
        Return FormatClockMinutes(startMinutes) & " - " & FormatClockMinutes(endMinutes)
    End Function

    Private Shared Function FormatClockMinutes(totalMinutes As Integer) As String
        Dim normalizedTotal As Integer = Math.Max(0, Math.Min(1440, totalMinutes))
        Dim hourValue As Integer = normalizedTotal \ 60
        Dim minuteValue As Integer = normalizedTotal Mod 60
        Return hourValue.ToString("00", CultureInfo.InvariantCulture) &
               ":" &
               minuteValue.ToString("00", CultureInfo.InvariantCulture)
    End Function

    Private Shared Function ReadRowValue(row As DataRow, columnName As String) As String
        If row Is Nothing OrElse row.Table Is Nothing OrElse Not row.Table.Columns.Contains(columnName) Then
            Return String.Empty
        End If

        If row.IsNull(columnName) Then
            Return String.Empty
        End If

        Return Convert.ToString(row(columnName)).Trim()
    End Function

    Private Shared Function BuildDataTableBindingPath(columnName As String) As String
        Dim safeColumnName As String = If(columnName, String.Empty)
        Return "[" & safeColumnName.Replace("]", "]]") & "]"
    End Function

    Private Shared Function ContainsIgnoreCase(values As IEnumerable(Of String), candidate As String) As Boolean
        If values Is Nothing Then
            Return False
        End If

        For Each value As String In values
            If String.Equals(value, candidate, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Shared Function NormalizeText(value As String) As String
        Return If(value, String.Empty).Trim()
    End Function
End Class
