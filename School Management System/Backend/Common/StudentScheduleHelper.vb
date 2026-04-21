Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text
Imports School_Management_System.Backend.Models

Namespace Backend.Common
    Public NotInheritable Class StudentScheduleHelper
        Private Sub New()
        End Sub

        Public Shared Function MatchesStudentSchedule(record As TeacherScheduleRecord,
                                                      student As StudentRecord,
                                                      subjectRecords As IEnumerable(Of SubjectRecord)) As Boolean
            If record Is Nothing OrElse student Is Nothing Then
                Return False
            End If

            Return SectionMatches(record.Section,
                                  student.SectionName,
                                  ResolveStudentYearLevel(student),
                                  ResolveSubjectYearLevel(record, subjectRecords))
        End Function

        Public Shared Function SectionMatches(scheduleSectionValue As String,
                                              studentSectionValue As String,
                                              Optional studentYearLevelValue As String = "",
                                              Optional subjectYearLevelValue As String = "") As Boolean
            Dim scheduleSectionToken As String =
                NormalizeComparisonToken(scheduleSectionValue)
            If String.IsNullOrWhiteSpace(scheduleSectionToken) Then
                Return False
            End If

            Dim studentBaseSectionValue As String =
                BuildSectionValue(studentSectionValue,
                                  studentYearLevelValue,
                                  String.Empty)
            Dim studentSectionToken As String =
                NormalizeComparisonToken(studentBaseSectionValue)
            If String.IsNullOrWhiteSpace(studentSectionToken) Then
                Return False
            End If

            Dim studentYearToken As String =
                NormalizeYearLevelToken(studentYearLevelValue)
            Dim subjectYearToken As String =
                NormalizeYearLevelToken(subjectYearLevelValue)
            Dim compactStudentSection As String = studentSectionToken

            If Not String.IsNullOrWhiteSpace(studentYearToken) Then
                compactStudentSection = studentYearToken & studentSectionToken
            End If

            Dim isSectionMatch As Boolean =
                String.Equals(scheduleSectionToken,
                              studentSectionToken,
                              StringComparison.OrdinalIgnoreCase) OrElse
                String.Equals(scheduleSectionToken,
                              compactStudentSection,
                              StringComparison.OrdinalIgnoreCase)
            If Not isSectionMatch Then
                Return False
            End If

            If String.IsNullOrWhiteSpace(studentYearToken) OrElse
               String.IsNullOrWhiteSpace(subjectYearToken) Then
                Return True
            End If

            Return String.Equals(studentYearToken,
                                 subjectYearToken,
                                 StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function BuildStudentSectionLabel(student As StudentRecord) As String
            Return BuildStudentSectionLabel(student, "--")
        End Function

        Public Shared Function BuildStudentSectionLabel(student As StudentRecord,
                                                        fallbackValue As String) As String
            If student Is Nothing Then
                Return fallbackValue
            End If

            Return BuildCompactSectionValue(student.SectionName,
                                            ResolveStudentYearLevel(student),
                                            fallbackValue)
        End Function

        Public Shared Function BuildStudentSectionValue(student As StudentRecord,
                                                        Optional fallbackValue As String = "--") As String
            If student Is Nothing Then
                Return fallbackValue
            End If

            Return BuildSectionValue(student.SectionName,
                                     ResolveStudentYearLevel(student),
                                     fallbackValue)
        End Function

        Public Shared Function BuildStudentYearLevelValue(student As StudentRecord,
                                                          Optional fallbackValue As String = "--") As String
            Return BuildYearLevelValue(ResolveStudentYearLevel(student), fallbackValue)
        End Function

        Public Shared Function BuildSectionValue(sectionValue As String,
                                                 Optional yearLevelValue As String = "",
                                                 Optional fallbackValue As String = "--") As String
            Dim compactSection As String = NormalizeSectionToken(sectionValue)
            Dim compactYearLevel As String = NormalizeYearLevelToken(yearLevelValue)

            If String.IsNullOrWhiteSpace(compactSection) Then
                Return fallbackValue
            End If

            If Not String.IsNullOrWhiteSpace(compactYearLevel) AndAlso
               compactSection.Length > compactYearLevel.Length AndAlso
               compactSection.StartsWith(compactYearLevel,
                                         StringComparison.OrdinalIgnoreCase) AndAlso
               Char.IsLetter(compactSection(compactYearLevel.Length)) Then
                compactSection = compactSection.Substring(compactYearLevel.Length).Trim()
            End If

            If String.IsNullOrWhiteSpace(compactSection) Then
                Return fallbackValue
            End If

            Return compactSection
        End Function

        Public Shared Function BuildCompactSectionValue(sectionValue As String,
                                                        yearLevelValue As String,
                                                        Optional fallbackValue As String = "--") As String
            Dim compactSection As String =
                BuildSectionValue(sectionValue, yearLevelValue, String.Empty)
            Dim compactYearLevel As String =
                BuildYearLevelValue(yearLevelValue, String.Empty)

            If String.IsNullOrWhiteSpace(compactSection) AndAlso
               String.IsNullOrWhiteSpace(compactYearLevel) Then
                Return fallbackValue
            End If

            If String.IsNullOrWhiteSpace(compactSection) Then
                Return compactYearLevel
            End If

            If String.IsNullOrWhiteSpace(compactYearLevel) Then
                Return compactSection
            End If

            Return compactYearLevel & compactSection
        End Function

        Public Shared Function BuildYearLevelValue(yearLevelValue As String,
                                                   Optional fallbackValue As String = "--") As String
            Dim compactYearLevel As String = NormalizeYearLevelToken(yearLevelValue)
            If String.IsNullOrWhiteSpace(compactYearLevel) Then
                Return fallbackValue
            End If

            Return compactYearLevel
        End Function

        Public Shared Function NormalizeDayLabel(dayValue As String) As String
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

        Public Shared Function GetDaySortOrder(dayValue As String) As Integer
            Select Case NormalizeDayLabel(dayValue)
                Case "Monday"
                    Return 1
                Case "Tuesday"
                    Return 2
                Case "Wednesday"
                    Return 3
                Case "Thursday"
                    Return 4
                Case "Friday"
                    Return 5
                Case Else
                    Return Integer.MaxValue
            End Select
        End Function

        Public Shared Function NormalizeSession(sessionValue As String) As String
            Dim startMinutes As Integer
            Dim endMinutes As Integer

            If TryParseSessionRange(sessionValue,
                                    startMinutes,
                                    endMinutes) Then
                Return FormatSessionRange(startMinutes, endMinutes)
            End If

            Return NormalizeText(sessionValue)
        End Function

        Public Shared Function TryGetSessionStartMinutes(sessionValue As String,
                                                         ByRef parsedStartMinutes As Integer) As Boolean
            Dim parsedEndMinutes As Integer
            Return TryParseSessionRange(sessionValue,
                                        parsedStartMinutes,
                                        parsedEndMinutes)
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

        Private Shared Function NormalizeSectionToken(sectionValue As String) As String
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

        Private Shared Function NormalizeComparisonToken(sectionValue As String) As String
            Dim compactSection As String = NormalizeSectionToken(sectionValue)
            If String.IsNullOrWhiteSpace(compactSection) Then
                Return String.Empty
            End If

            Dim builder As New StringBuilder()

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

        Private Shared Function TryParseSessionRange(sessionValue As String,
                                                     ByRef parsedStartMinutes As Integer,
                                                     ByRef parsedEndMinutes As Integer) As Boolean
            Dim normalizedSession As String = NormalizeText(sessionValue)
            If String.IsNullOrWhiteSpace(normalizedSession) Then
                Return False
            End If

            Dim startToken As String = String.Empty
            Dim endToken As String = String.Empty

            If Not TrySplitSessionRange(normalizedSession,
                                        startToken,
                                        endToken) Then
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
            Dim separatorIndex As Integer =
                sessionValue.IndexOf(separator, StringComparison.Ordinal)
            If separatorIndex <= 0 Then
                separator = " to "
                separatorIndex =
                    sessionValue.IndexOf(separator, StringComparison.OrdinalIgnoreCase)
            End If

            If separatorIndex <= 0 Then
                separator = "-"
                separatorIndex = sessionValue.IndexOf(separator, StringComparison.Ordinal)
            End If

            If separatorIndex <= 0 Then
                Return False
            End If

            startToken = NormalizeText(sessionValue.Substring(0, separatorIndex))
            endToken =
                NormalizeText(sessionValue.Substring(separatorIndex + separator.Length))
            If String.IsNullOrWhiteSpace(startToken) OrElse
               String.IsNullOrWhiteSpace(endToken) Then
                Return False
            End If

            Return True
        End Function

        Private Shared Function TryParseClockToken(value As String,
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

        Private Shared Function FormatSessionRange(startMinutes As Integer,
                                                   endMinutes As Integer) As String
            Return FormatClockMinutes(startMinutes) &
                   " - " &
                   FormatClockMinutes(endMinutes)
        End Function

        Private Shared Function FormatClockMinutes(totalMinutes As Integer) As String
            Dim normalizedTotal As Integer =
                Math.Max(0, Math.Min(1440, totalMinutes))
            Dim hourValue As Integer = normalizedTotal \ 60
            Dim minuteValue As Integer = normalizedTotal Mod 60

            Return hourValue.ToString("00", CultureInfo.InvariantCulture) &
                   ":" &
                   minuteValue.ToString("00", CultureInfo.InvariantCulture)
        End Function

        Private Shared Function NormalizeText(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function
    End Class
End Namespace
