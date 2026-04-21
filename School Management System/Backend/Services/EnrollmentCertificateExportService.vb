Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models

Namespace Backend.Services
    Public Class EnrollmentCertificateExportService
        Private Class CertificateScheduleEntry
            Public Property DayLabel As String = String.Empty
            Public Property DaySortOrder As Integer = Integer.MaxValue
            Public Property SessionLabel As String = String.Empty
            Public Property SessionSortOrder As Integer = Integer.MaxValue
            Public Property SubjectLabel As String = String.Empty
            Public Property RoomLabel As String = String.Empty
            Public Property TeacherLabel As String = String.Empty
            Public Property SectionYearLabel As String = String.Empty
        End Class

        Private Class CertificateTimetableRow
            Public Property SessionLabel As String = String.Empty
            Public Property SessionSortOrder As Integer = Integer.MaxValue
            Public Property DayCells As New Dictionary(Of String, String)(
                StringComparer.OrdinalIgnoreCase)
        End Class

        Private Class PdfImageResource
            Public Property ResourceName As String = String.Empty
            Public Property Width As Integer
            Public Property Height As Integer
            Public Property ImageBytes As Byte()
        End Class

        Private NotInheritable Class PdfPageBuilder
            Private Shared ReadOnly Invariant As CultureInfo =
                CultureInfo.InvariantCulture

            Public Const PageWidth As Double = 595D
            Public Const PageHeight As Double = 842D

            Private ReadOnly _content As New StringBuilder()

            Public Sub FillRectangle(left As Double,
                                     top As Double,
                                     width As Double,
                                     height As Double,
                                     fillColor As String)
                If width <= 0D OrElse height <= 0D Then
                    Return
                End If

                _content.AppendLine(BuildColorCommand(fillColor, isStroke:=False))
                _content.AppendLine(String.Format(Invariant,
                                                  "{0} {1} {2} {3} re f",
                                                  FormatNumber(left),
                                                  FormatNumber(ToPdfY(top + height)),
                                                  FormatNumber(width),
                                                  FormatNumber(height)))
            End Sub

            Public Sub StrokeRectangle(left As Double,
                                       top As Double,
                                       width As Double,
                                       height As Double,
                                       strokeColor As String,
                                       lineWidth As Double)
                If width <= 0D OrElse height <= 0D Then
                    Return
                End If

                _content.AppendLine(String.Format(Invariant,
                                                  "{0} w",
                                                  FormatNumber(Math.Max(0.25D,
                                                                        lineWidth))))
                _content.AppendLine(BuildColorCommand(strokeColor, isStroke:=True))
                _content.AppendLine(String.Format(Invariant,
                                                  "{0} {1} {2} {3} re S",
                                                  FormatNumber(left),
                                                  FormatNumber(ToPdfY(top + height)),
                                                  FormatNumber(width),
                                                  FormatNumber(height)))
            End Sub

            Public Sub DrawLine(startX As Double,
                                startTop As Double,
                                endX As Double,
                                endTop As Double,
                                strokeColor As String,
                                lineWidth As Double)
                _content.AppendLine(String.Format(Invariant,
                                                  "{0} w",
                                                  FormatNumber(Math.Max(0.25D,
                                                                        lineWidth))))
                _content.AppendLine(BuildColorCommand(strokeColor, isStroke:=True))
                _content.AppendLine(String.Format(Invariant,
                                                  "{0} {1} m",
                                                  FormatNumber(startX),
                                                  FormatNumber(ToPdfY(startTop))))
                _content.AppendLine(String.Format(Invariant,
                                                  "{0} {1} l S",
                                                  FormatNumber(endX),
                                                  FormatNumber(ToPdfY(endTop))))
            End Sub

            Public Sub DrawText(text As String,
                                left As Double,
                                top As Double,
                                fontResource As String,
                                fontSize As Double,
                                fillColor As String)
                If String.IsNullOrWhiteSpace(text) Then
                    Return
                End If

                Dim normalizedFont As String = If(fontResource, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalizedFont) Then
                    normalizedFont = "F1"
                End If

                _content.AppendLine("BT")
                _content.AppendLine(BuildColorCommand(fillColor, isStroke:=False))
                _content.AppendLine(String.Format(Invariant,
                                                  "/{0} {1} Tf",
                                                  normalizedFont,
                                                  FormatNumber(fontSize)))
                _content.AppendLine(String.Format(Invariant,
                                                  "1 0 0 1 {0} {1} Tm",
                                                  FormatNumber(left),
                                                  FormatNumber(ToPdfY(top + fontSize))))
                _content.AppendLine(EncodeText(text) & " Tj")
                _content.AppendLine("ET")
            End Sub

            Public Sub DrawImage(resourceName As String,
                                 left As Double,
                                 top As Double,
                                 width As Double,
                                 height As Double)
                Dim normalizedResourceName As String =
                    If(resourceName, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalizedResourceName) OrElse
                   width <= 0D OrElse
                   height <= 0D Then
                    Return
                End If

                _content.AppendLine("q")
                _content.AppendLine(String.Format(Invariant,
                                                  "{0} 0 0 {1} {2} {3} cm",
                                                  FormatNumber(width),
                                                  FormatNumber(height),
                                                  FormatNumber(left),
                                                  FormatNumber(ToPdfY(top + height))))
                _content.AppendLine("/" & normalizedResourceName & " Do")
                _content.AppendLine("Q")
            End Sub

            Public Function BuildContent() As String
                Return _content.ToString()
            End Function

            Private Shared Function BuildColorCommand(colorHex As String,
                                                      isStroke As Boolean) As String
                Dim red As Double
                Dim green As Double
                Dim blue As Double
                ParseColor(colorHex, red, green, blue)

                Return String.Format(Invariant,
                                     "{0} {1} {2} {3}",
                                     FormatNumber(red),
                                     FormatNumber(green),
                                     FormatNumber(blue),
                                     If(isStroke, "RG", "rg"))
            End Function

            Private Shared Sub ParseColor(colorHex As String,
                                          ByRef red As Double,
                                          ByRef green As Double,
                                          ByRef blue As Double)
                red = 0D
                green = 0D
                blue = 0D

                Dim normalizedColor As String =
                    If(colorHex, String.Empty).Trim().TrimStart("#"c)
                If normalizedColor.Length <> 6 Then
                    Return
                End If

                Dim parsedRed As Integer
                Dim parsedGreen As Integer
                Dim parsedBlue As Integer

                If Integer.TryParse(normalizedColor.Substring(0, 2),
                                    NumberStyles.HexNumber,
                                    Invariant,
                                    parsedRed) AndAlso
                   Integer.TryParse(normalizedColor.Substring(2, 2),
                                    NumberStyles.HexNumber,
                                    Invariant,
                                    parsedGreen) AndAlso
                   Integer.TryParse(normalizedColor.Substring(4, 2),
                                    NumberStyles.HexNumber,
                                    Invariant,
                                    parsedBlue) Then
                    red = parsedRed / 255D
                    green = parsedGreen / 255D
                    blue = parsedBlue / 255D
                End If
            End Sub

            Private Shared Function EncodeText(value As String) As String
                Dim normalizedText As New StringBuilder()

                For Each characterValue As Char In If(value, String.Empty)
                    Dim codePoint As Integer = AscW(characterValue)
                    If codePoint < 32 Then
                        normalizedText.Append(" "c)
                    ElseIf codePoint <= 255 Then
                        normalizedText.Append(characterValue)
                    Else
                        normalizedText.Append("?"c)
                    End If
                Next

                Dim bytes As Byte() = Encoding.Latin1.GetBytes(normalizedText.ToString())
                Dim builder As New StringBuilder()
                builder.Append("<"c)

                For Each byteValue As Byte In bytes
                    builder.Append(byteValue.ToString("X2", Invariant))
                Next

                builder.Append(">"c)
                Return builder.ToString()
            End Function

            Private Shared Function FormatNumber(value As Double) As String
                Return value.ToString("0.##", Invariant)
            End Function

            Private Shared Function ToPdfY(topPosition As Double) As Double
                Return PageHeight - topPosition
            End Function
        End Class

        Private Shared ReadOnly Invariant As CultureInfo =
            CultureInfo.InvariantCulture
        Private Shared ReadOnly TimetableDayHeaders As String() = New String() {
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday"
        }
        Private Const TimetableTitleHeaderTop As Double = 232D
        Private Const TimetableContinuationHeaderTop As Double = 150D
        Private Const TimetableHeaderHeight As Double = 22D

        Private ReadOnly _teacherScheduleManagementService As TeacherScheduleManagementService

        Public Sub New()
            Me.New(New TeacherScheduleManagementService())
        End Sub

        Public Sub New(teacherScheduleManagementService As TeacherScheduleManagementService)
            _teacherScheduleManagementService = teacherScheduleManagementService
        End Sub

        Public Function ExportCertificate(snapshot As StudentEnrollmentSnapshot,
                                          targetFilePath As String) As ServiceResult(Of String)
            If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
                Return ServiceResult(Of String).Failure(
                    "No student enrollment data is available for certificate export.")
            End If

            If Not snapshot.IsFullyEnrolled Then
                Return ServiceResult(Of String).Failure(
                    "Finish enrollment before exporting the certificate of enrollment.")
            End If

            Dim normalizedFilePath As String = If(targetFilePath, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedFilePath) Then
                Return ServiceResult(Of String).Failure("A PDF file path is required.")
            End If

            Dim scheduleResult As ServiceResult(Of List(Of TeacherScheduleRecord)) =
                _teacherScheduleManagementService.GetSchedules()
            If scheduleResult Is Nothing OrElse Not scheduleResult.IsSuccess Then
                Dim message As String =
                    If(scheduleResult Is Nothing,
                       "Unable to load teacher schedules for the certificate.",
                       scheduleResult.Message)
                Return ServiceResult(Of String).Failure(message)
            End If

            Try
                Dim scheduleEntries As List(Of CertificateScheduleEntry) =
                    BuildScheduleEntries(scheduleResult.Data,
                                         snapshot.Student,
                                         snapshot.SelectedSubjects)
                Dim timetableRows As List(Of CertificateTimetableRow) =
                    BuildTimetableRows(scheduleEntries)
                Dim schoolLogo As PdfImageResource = LoadSchoolLogoImage()

                Dim targetDirectory As String =
                    Path.GetDirectoryName(normalizedFilePath)
                If Not String.IsNullOrWhiteSpace(targetDirectory) Then
                    Directory.CreateDirectory(targetDirectory)
                End If

                Dim pageContents As List(Of String) =
                    BuildCertificatePages(snapshot, timetableRows, schoolLogo)
                WritePdfDocument(normalizedFilePath, pageContents, schoolLogo)

                Return ServiceResult(Of String).Success(
                    normalizedFilePath,
                    "Certificate of enrollment exported successfully.")
            Catch ex As Exception
                Return ServiceResult(Of String).Failure(
                    "Unable to export the certificate of enrollment." &
                    Environment.NewLine &
                    ex.Message)
            End Try
        End Function

        Private Function BuildScheduleEntries(records As IEnumerable(Of TeacherScheduleRecord),
                                              student As StudentRecord,
                                              selectedSubjects As IEnumerable(Of SubjectRecord)) As List(Of CertificateScheduleEntry)
            Dim entries As New List(Of CertificateScheduleEntry)()

            If records Is Nothing OrElse student Is Nothing Then
                Return entries
            End If

            Dim timetableSectionYearLabel As String =
                BuildTimetableSectionYearLabel(student)

            For Each record As TeacherScheduleRecord In records
                If Not StudentScheduleHelper.MatchesStudentSchedule(record,
                                                                   student,
                                                                   selectedSubjects) Then
                    Continue For
                End If

                Dim normalizedDay As String =
                    StudentScheduleHelper.NormalizeDayLabel(record.Day)
                Dim normalizedSession As String =
                    StudentScheduleHelper.NormalizeSession(record.Session)
                Dim sessionStart As Integer = Integer.MaxValue
                Dim parsedStart As Integer

                If StudentScheduleHelper.TryGetSessionStartMinutes(record.Session,
                                                                   parsedStart) Then
                    sessionStart = parsedStart
                End If

                entries.Add(New CertificateScheduleEntry() With {
                    .DayLabel = NormalizeValue(normalizedDay,
                                               NormalizeValue(record.Day, "TBA")),
                    .DaySortOrder = StudentScheduleHelper.GetDaySortOrder(record.Day),
                    .SessionLabel = NormalizeValue(normalizedSession, "Time TBA"),
                    .SessionSortOrder = sessionStart,
                    .SubjectLabel = BuildSubjectLabel(record),
                    .RoomLabel = NormalizeValue(record.Room, "Room TBA"),
                    .TeacherLabel = BuildTeacherLabel(record),
                    .SectionYearLabel = timetableSectionYearLabel
                })
            Next

            entries.Sort(AddressOf CompareScheduleEntries)
            Return entries
        End Function

        Private Function CompareScheduleEntries(left As CertificateScheduleEntry,
                                                right As CertificateScheduleEntry) As Integer
            Dim dayComparison As Integer =
                If(left Is Nothing, Integer.MaxValue, left.DaySortOrder).CompareTo(
                    If(right Is Nothing, Integer.MaxValue, right.DaySortOrder))
            If dayComparison <> 0 Then
                Return dayComparison
            End If

            Dim sessionComparison As Integer =
                If(left Is Nothing, Integer.MaxValue, left.SessionSortOrder).CompareTo(
                    If(right Is Nothing, Integer.MaxValue, right.SessionSortOrder))
            If sessionComparison <> 0 Then
                Return sessionComparison
            End If

            Return StringComparer.OrdinalIgnoreCase.Compare(
                If(left Is Nothing, String.Empty, left.SubjectLabel),
                If(right Is Nothing, String.Empty, right.SubjectLabel))
        End Function

        Private Function BuildTimetableRows(entries As IEnumerable(Of CertificateScheduleEntry)) As List(Of CertificateTimetableRow)
            Dim rowsBySession As New Dictionary(Of String, CertificateTimetableRow)(
                StringComparer.OrdinalIgnoreCase)

            If entries IsNot Nothing Then
                For Each entry As CertificateScheduleEntry In entries
                    If entry Is Nothing Then
                        Continue For
                    End If

                    Dim sessionLabel As String =
                        NormalizeValue(entry.SessionLabel, "Time TBA")
                    Dim row As CertificateTimetableRow = Nothing

                    If Not rowsBySession.TryGetValue(sessionLabel, row) Then
                        row = New CertificateTimetableRow() With {
                            .SessionLabel = sessionLabel,
                            .SessionSortOrder = entry.SessionSortOrder
                        }

                        For Each dayHeader As String In TimetableDayHeaders
                            row.DayCells(dayHeader) = "--"
                        Next

                        rowsBySession(sessionLabel) = row
                    ElseIf entry.SessionSortOrder < row.SessionSortOrder Then
                        row.SessionSortOrder = entry.SessionSortOrder
                    End If

                    Dim targetDay As String =
                        StudentScheduleHelper.NormalizeDayLabel(entry.DayLabel)
                    If String.IsNullOrWhiteSpace(targetDay) OrElse
                       Not row.DayCells.ContainsKey(targetDay) Then
                        Continue For
                    End If

                    Dim cellText As String = BuildTimetableCellText(entry)
                    Dim currentCellText As String = row.DayCells(targetDay)
                    If currentCellText = "--" Then
                        currentCellText = String.Empty
                    End If

                    row.DayCells(targetDay) =
                        AppendCellText(currentCellText, cellText)
                Next
            End If

            Dim rows As New List(Of CertificateTimetableRow)(rowsBySession.Values)
            rows.Sort(Function(left, right)
                          Dim sessionComparison As Integer =
                              If(left Is Nothing,
                                 Integer.MaxValue,
                                 left.SessionSortOrder).CompareTo(
                                     If(right Is Nothing,
                                        Integer.MaxValue,
                                        right.SessionSortOrder))
                          If sessionComparison <> 0 Then
                              Return sessionComparison
                          End If

                          Return StringComparer.OrdinalIgnoreCase.Compare(
                              If(left Is Nothing, String.Empty, left.SessionLabel),
                              If(right Is Nothing, String.Empty, right.SessionLabel))
                      End Function)
            Return rows
        End Function

        Private Function BuildCertificatePages(snapshot As StudentEnrollmentSnapshot,
                                               timetableRows As IList(Of CertificateTimetableRow),
                                               schoolLogo As PdfImageResource) As List(Of String)
            Dim pages As New List(Of String)()
            Dim student As StudentRecord = snapshot.Student
            Dim issueDateLabel As String =
                DateTime.Now.ToString("MMMM dd, yyyy", Invariant)
            Dim studentName As String =
                NormalizeValue(student.FullName,
                               NormalizeValue(student.StudentNumber, "Student"))
            Dim yearLabel As String =
                StudentScheduleHelper.BuildStudentYearLevelValue(student,
                                                                 "Year not assigned")
            Dim sectionLabel As String =
                StudentScheduleHelper.BuildStudentSectionValue(student, "--")
            Dim courseLabel As String =
                NormalizeValue(student.CourseDisplayName, "--")
            Dim currentPage As PdfPageBuilder =
                BuildPrintFriendlyTitlePage(snapshot,
                                            issueDateLabel,
                                            studentName,
                                            yearLabel,
                                            sectionLabel,
                                            courseLabel,
                                            schoolLogo)
            Dim currentTop As Double = TimetableTitleHeaderTop + TimetableHeaderHeight
            Dim nextRowIndex As Integer = 0

            If timetableRows Is Nothing OrElse timetableRows.Count = 0 Then
                DrawNoScheduleMessage(currentPage, currentTop)
                AddFooter(currentPage, issueDateLabel)
                pages.Add(currentPage.BuildContent())
                Return pages
            End If

            Do
                nextRowIndex = DrawTimetableRows(currentPage,
                                                currentTop,
                                                timetableRows,
                                                nextRowIndex)
                AddFooter(currentPage, issueDateLabel)
                pages.Add(currentPage.BuildContent())

                If nextRowIndex >= timetableRows.Count Then
                    Exit Do
                End If

                currentPage = BuildPrintFriendlyContinuationPage(studentName,
                                                                 sectionLabel,
                                                                 issueDateLabel,
                                                                 schoolLogo)
                currentTop = TimetableContinuationHeaderTop + TimetableHeaderHeight
            Loop

            Return pages
        End Function

        Private Function BuildPrintFriendlyTitlePage(snapshot As StudentEnrollmentSnapshot,
                                                     issueDateLabel As String,
                                                     studentName As String,
                                                     yearLabel As String,
                                                     sectionLabel As String,
                                                     courseLabel As String,
                                                     schoolLogo As PdfImageResource) As PdfPageBuilder
            Dim student As StudentRecord = snapshot.Student
            Dim page As New PdfPageBuilder()

            DrawLogoIfAvailable(page, schoolLogo, 34D, 28D, 44D)
            page.DrawText("PRESIDENT RAMON MAGSAYSAY STATE UNIVERSITY",
                          94D,
                          28D,
                          "F2",
                          12D,
                          "000000")
            page.DrawText("Certificate of Enrollment",
                          94D,
                          50D,
                          "F2",
                          18D,
                          "000000")
            page.DrawText("Issued " & issueDateLabel,
                          432D,
                          54D,
                          "F1",
                          10D,
                          "000000")
            page.DrawLine(34D, 84D, 561D, 84D, "000000", 0.8D)

            Dim bodyLines As List(Of String) =
                WrapText("This certifies that " &
                         studentName &
                         " is officially enrolled in the current term. " &
                         "The timetable below reflects the student's approved class schedule.",
                         527D,
                         11D)
            Dim bodyTop As Double = 98D

            For Each bodyLine As String In bodyLines
                page.DrawText(bodyLine,
                              34D,
                              bodyTop,
                              "F1",
                              11D,
                              "000000")
                bodyTop += 14D
            Next

            WriteMetadataLine(page, 34D, 138D, "Student Name", studentName, 250D)
            WriteMetadataLine(page,
                              300D,
                              138D,
                              "Student ID",
                              NormalizeValue(student.StudentNumber, "--"),
                              261D)
            WriteMetadataLine(page, 34D, 160D, "Year Level", yearLabel, 250D)
            WriteMetadataLine(page, 300D, 160D, "Section", sectionLabel, 261D)
            WriteMetadataLine(page, 34D, 182D, "Course", courseLabel, 527D)

            page.DrawText("Class Timetable",
                          34D,
                          212D,
                          "F2",
                          12D,
                          "000000")
            DrawTimetableHeaderRow(page, TimetableTitleHeaderTop)
            Return page
        End Function

        Private Function BuildPrintFriendlyContinuationPage(studentName As String,
                                                            sectionLabel As String,
                                                            issueDateLabel As String,
                                                            schoolLogo As PdfImageResource) As PdfPageBuilder
            Dim page As New PdfPageBuilder()

            DrawLogoIfAvailable(page, schoolLogo, 34D, 28D, 36D)
            page.DrawText("Certificate of Enrollment",
                          84D,
                          32D,
                          "F2",
                          15D,
                          "000000")
            page.DrawText(studentName & " | " & sectionLabel,
                          84D,
                          52D,
                          "F1",
                          10D,
                          "000000")
            page.DrawText("Issued " & issueDateLabel,
                          432D,
                          52D,
                          "F1",
                          10D,
                          "000000")
            page.DrawLine(34D, 84D, 561D, 84D, "000000", 0.8D)
            page.DrawText("Class Timetable",
                          34D,
                          124D,
                          "F2",
                          12D,
                          "000000")
            DrawTimetableHeaderRow(page, TimetableContinuationHeaderTop)
            Return page
        End Function

        Private Sub DrawLogoIfAvailable(page As PdfPageBuilder,
                                        schoolLogo As PdfImageResource,
                                        left As Double,
                                        top As Double,
                                        targetHeight As Double)
            If page Is Nothing OrElse schoolLogo Is Nothing OrElse
               schoolLogo.Height <= 0 Then
                Return
            End If

            Dim targetWidth As Double =
                (schoolLogo.Width / CDbl(schoolLogo.Height)) * targetHeight
            page.DrawImage(schoolLogo.ResourceName,
                           left,
                           top,
                           targetWidth,
                           targetHeight)
        End Sub

        Private Sub WriteMetadataLine(page As PdfPageBuilder,
                                      left As Double,
                                      top As Double,
                                      label As String,
                                      value As String,
                                      availableWidth As Double)
            page.DrawText(label & ":",
                          left,
                          top,
                          "F2",
                          10D,
                          "000000")

            Dim labelWidth As Double = EstimateTextWidth(label & ":", 10D) + 8D
            Dim valueLines As List(Of String) =
                WrapText(value,
                         Math.Max(40D, availableWidth - labelWidth),
                         10D)
            Dim valueTop As Double = top

            For Each valueLine As String In valueLines
                page.DrawText(valueLine,
                              left + labelWidth,
                              valueTop,
                              "F1",
                              10D,
                              "000000")
                valueTop += 12D
            Next
        End Sub

        Private Sub DrawTimetableHeaderRow(page As PdfPageBuilder,
                                           top As Double)
            Const leftMargin As Double = 34D
            Const totalWidth As Double = 527D
            Const timeColumnWidth As Double = 82D
            Dim dayColumnWidth As Double = (totalWidth - timeColumnWidth) / 5D

            page.StrokeRectangle(leftMargin,
                                 top,
                                 totalWidth,
                                 TimetableHeaderHeight,
                                 "000000",
                                 0.7D)
            page.DrawLine(leftMargin + timeColumnWidth,
                          top,
                          leftMargin + timeColumnWidth,
                          top + TimetableHeaderHeight,
                          "000000",
                          0.6D)

            DrawCenteredText(page,
                             "Time",
                             leftMargin,
                             top + 6D,
                             timeColumnWidth,
                             "F2",
                             9.5D)

            For index As Integer = 0 To TimetableDayHeaders.Length - 1
                Dim columnLeft As Double =
                    leftMargin + timeColumnWidth + (index * dayColumnWidth)

                If index > 0 Then
                    page.DrawLine(columnLeft,
                                  top,
                                  columnLeft,
                                  top + TimetableHeaderHeight,
                                  "000000",
                                  0.6D)
                End If

                DrawCenteredText(page,
                                 TimetableDayHeaders(index),
                                 columnLeft,
                                 top + 6D,
                                 dayColumnWidth,
                                 "F2",
                                 9.5D)
            Next
        End Sub

        Private Function DrawTimetableRows(page As PdfPageBuilder,
                                           top As Double,
                                           rows As IList(Of CertificateTimetableRow),
                                           startIndex As Integer) As Integer
            Const totalWidth As Double = 527D
            Const timeColumnWidth As Double = 82D
            Const pageBottom As Double = 774D
            Dim dayColumnWidth As Double = (totalWidth - timeColumnWidth) / 5D
            Dim currentTop As Double = top
            Dim nextIndex As Integer = startIndex

            While rows IsNot Nothing AndAlso nextIndex < rows.Count
                Dim row As CertificateTimetableRow = rows(nextIndex)
                Dim rowHeight As Double =
                    MeasureTimetableRowHeight(row, timeColumnWidth, dayColumnWidth)
                If currentTop + rowHeight > pageBottom Then
                    Exit While
                End If

                DrawTimetableRow(page,
                                 currentTop,
                                 row,
                                 timeColumnWidth,
                                 dayColumnWidth)
                currentTop += rowHeight
                nextIndex += 1
            End While

            Return nextIndex
        End Function

        Private Sub DrawTimetableRow(page As PdfPageBuilder,
                                     top As Double,
                                     row As CertificateTimetableRow,
                                     timeColumnWidth As Double,
                                     dayColumnWidth As Double)
            Const leftMargin As Double = 34D
            Const totalWidth As Double = 527D
            Dim rowHeight As Double =
                MeasureTimetableRowHeight(row, timeColumnWidth, dayColumnWidth)

            page.StrokeRectangle(leftMargin,
                                 top,
                                 totalWidth,
                                 rowHeight,
                                 "000000",
                                 0.55D)
            page.DrawLine(leftMargin + timeColumnWidth,
                          top,
                          leftMargin + timeColumnWidth,
                          top + rowHeight,
                          "000000",
                          0.5D)

            Dim sessionLines As List(Of String) =
                WrapText(If(row Is Nothing, String.Empty, row.SessionLabel),
                         timeColumnWidth - 8D,
                         9D)
            DrawCenteredTextLines(page,
                                  sessionLines,
                                  leftMargin,
                                  top,
                                  timeColumnWidth,
                                  rowHeight,
                                  "F1",
                                  9D,
                                  10D)

            For index As Integer = 0 To TimetableDayHeaders.Length - 1
                Dim columnLeft As Double =
                    leftMargin + timeColumnWidth + (index * dayColumnWidth)
                If index > 0 Then
                    page.DrawLine(columnLeft,
                                  top,
                                  columnLeft,
                                  top + rowHeight,
                                  "000000",
                                  0.5D)
                End If

                Dim dayHeader As String = TimetableDayHeaders(index)
                Dim cellText As String = "--"
                If row IsNot Nothing AndAlso row.DayCells.ContainsKey(dayHeader) Then
                    cellText = row.DayCells(dayHeader)
                End If

                Dim cellLines As List(Of String) =
                    WrapText(cellText, dayColumnWidth - 8D, 8.5D)
                DrawCenteredTextLines(page,
                                      cellLines,
                                      columnLeft,
                                      top,
                                      dayColumnWidth,
                                      rowHeight,
                                      "F1",
                                      8.5D,
                                      10D)
            Next
        End Sub

        Private Sub DrawNoScheduleMessage(page As PdfPageBuilder,
                                          top As Double)
            page.StrokeRectangle(34D, top, 527D, 46D, "000000", 0.7D)
            page.DrawText("No teacher schedule is available for the current section yet.",
                          42D,
                          top + 10D,
                          "F2",
                          10.5D,
                          "000000")
            page.DrawText("The certificate still reflects the student's approved enrollment record.",
                          42D,
                          top + 24D,
                          "F1",
                          10D,
                          "000000")
        End Sub

        Private Shared Function MeasureTimetableRowHeight(row As CertificateTimetableRow,
                                                          timeColumnWidth As Double,
                                                          dayColumnWidth As Double) As Double
            Dim maxLineCount As Integer =
                Math.Max(1,
                         WrapText(If(row Is Nothing, String.Empty, row.SessionLabel),
                                  timeColumnWidth - 8D,
                                  9D).Count)

            For Each dayHeader As String In TimetableDayHeaders
                Dim cellText As String = "--"
                If row IsNot Nothing AndAlso row.DayCells.ContainsKey(dayHeader) Then
                    cellText = row.DayCells(dayHeader)
                End If

                maxLineCount = Math.Max(maxLineCount,
                                        WrapText(cellText,
                                                 dayColumnWidth - 8D,
                                                 8.5D).Count)
            Next

            Return Math.Max(24D, 10D + (maxLineCount * 10D))
        End Function

        Private Shared Sub DrawCenteredText(page As PdfPageBuilder,
                                            text As String,
                                            left As Double,
                                            top As Double,
                                            width As Double,
                                            fontResource As String,
                                            fontSize As Double)
            Dim textWidth As Double = EstimateTextWidth(text, fontSize)
            Dim drawLeft As Double = left + Math.Max(4D, (width - textWidth) / 2D)

            page.DrawText(text,
                          drawLeft,
                          top,
                          fontResource,
                          fontSize,
                          "000000")
        End Sub

        Private Shared Sub DrawCenteredTextLines(page As PdfPageBuilder,
                                                 lines As IList(Of String),
                                                 left As Double,
                                                 top As Double,
                                                 width As Double,
                                                 height As Double,
                                                 fontResource As String,
                                                 fontSize As Double,
                                                 lineHeight As Double)
            Dim safeLines As IList(Of String) = lines
            If safeLines Is Nothing OrElse safeLines.Count = 0 Then
                safeLines = New List(Of String) From {String.Empty}
            End If

            Dim blockHeight As Double = safeLines.Count * lineHeight
            Dim drawTop As Double =
                top + Math.Max(4D, (height - blockHeight) / 2D)

            For Each line As String In safeLines
                DrawCenteredText(page,
                                 line,
                                 left,
                                 drawTop,
                                 width,
                                 fontResource,
                                 fontSize)
                drawTop += lineHeight
            Next
        End Sub

        Private Shared Function BuildTimetableCellText(entry As CertificateScheduleEntry) As String
            If entry Is Nothing Then
                Return "--"
            End If

            Dim lines As New List(Of String)()
            Dim subjectText As String = NormalizeValue(entry.SubjectLabel, String.Empty)
            Dim teacherText As String = NormalizeValue(entry.TeacherLabel, String.Empty)
            Dim sectionYearText As String =
                NormalizeValue(entry.SectionYearLabel, String.Empty)
            Dim roomText As String = NormalizeValue(entry.RoomLabel, String.Empty)

            If Not String.IsNullOrWhiteSpace(subjectText) Then
                lines.Add(subjectText)
            End If

            If Not String.IsNullOrWhiteSpace(teacherText) Then
                lines.Add(teacherText)
            End If

            If Not String.IsNullOrWhiteSpace(sectionYearText) Then
                lines.Add(sectionYearText)
            End If

            If Not String.IsNullOrWhiteSpace(roomText) AndAlso
               Not String.Equals(roomText, "Room TBA", StringComparison.OrdinalIgnoreCase) Then
                lines.Add(roomText)
            End If

            If lines.Count = 0 Then
                Return "--"
            End If

            Return String.Join(Environment.NewLine, lines)
        End Function

        Private Shared Function AppendCellText(currentValue As String,
                                               appendedValue As String) As String
            Dim leftValue As String = If(currentValue, String.Empty).Trim()
            Dim rightValue As String = If(appendedValue, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(leftValue) Then
                Return NormalizeValue(rightValue)
            End If

            If String.IsNullOrWhiteSpace(rightValue) Then
                Return NormalizeValue(leftValue)
            End If

            If String.Equals(leftValue,
                             rightValue,
                             StringComparison.OrdinalIgnoreCase) Then
                Return leftValue
            End If

            Return leftValue & Environment.NewLine & rightValue
        End Function

        Private Function BuildCertificateIntroPage(snapshot As StudentEnrollmentSnapshot,
                                                   issueDateLabel As String,
                                                   studentName As String,
                                                   yearLabel As String,
                                                   sectionLabel As String) As PdfPageBuilder
            Dim student As StudentRecord = snapshot.Student
            Dim page As New PdfPageBuilder()

            page.FillRectangle(0D, 0D, PdfPageBuilder.PageWidth, 112D, "1E3A8A")
            page.DrawText("President Ramon Magsaysay State University",
                          44D,
                          28D,
                          "F2",
                          14D,
                          "FFFFFF")
            page.DrawText("Certificate of Enrollment",
                          44D,
                          54D,
                          "F2",
                          25D,
                          "FFFFFF")
            page.DrawText("Generated from the School Management System",
                          44D,
                          84D,
                          "F1",
                          10D,
                          "DBEAFE")
            page.DrawText("Issued on " & issueDateLabel,
                          390D,
                          84D,
                          "F1",
                          10D,
                          "DBEAFE")

            Dim bodyLines As List(Of String) =
                WrapText("This certifies that " &
                         studentName &
                         " is currently enrolled as a " &
                         yearLabel &
                         " student under " &
                         sectionLabel &
                         ". The schedule below reflects the teacher schedules " &
                         "assigned to the student's active load.",
                         500D,
                         12D)
            Dim bodyY As Double = 142D

            For Each bodyLine As String In bodyLines
                page.DrawText(bodyLine,
                              46D,
                              bodyY,
                              "F1",
                              12D,
                              "1F2937")
                bodyY += 16D
            Next

            page.FillRectangle(44D, 204D, 507D, 126D, "F8FBFF")
            page.StrokeRectangle(44D, 204D, 507D, 126D, "D7E3F4", 1D)

            WriteDetailBlock(page,
                             60D,
                             224D,
                             "Student Name",
                             studentName)
            WriteDetailBlock(page,
                             60D,
                             264D,
                             "Student ID",
                             NormalizeValue(student.StudentNumber, "--"))
            WriteDetailBlock(page,
                             60D,
                             304D,
                             "Course",
                             NormalizeValue(student.CourseDisplayName, "--"))

            WriteDetailBlock(page,
                             312D,
                             224D,
                             "Year Level",
                             yearLabel)
            WriteDetailBlock(page,
                             312D,
                             264D,
                             "Section",
                             sectionLabel)
            WriteDetailBlock(page,
                             312D,
                             304D,
                             "Current Load",
                             snapshot.SelectedSubjectCount.ToString(Invariant) &
                             " subjects | " &
                             snapshot.SelectedTotalUnitsLabel &
                             " units")

            page.DrawText("Current Class Schedule",
                          44D,
                          350D,
                          "F2",
                          16D,
                          "0F172A")
            page.DrawLine(44D, 370D, 551D, 370D, "CBD5E1", 1D)

            Return page
        End Function

        Private Function BuildContinuationPage(studentName As String,
                                               sectionLabel As String,
                                               issueDateLabel As String) As PdfPageBuilder
            Dim page As New PdfPageBuilder()

            page.DrawText("Certificate of Enrollment",
                          44D,
                          34D,
                          "F2",
                          18D,
                          "0F172A")
            page.DrawText(studentName & " | " & sectionLabel,
                          44D,
                          60D,
                          "F1",
                          10D,
                          "475569")
            page.DrawText("Issued on " & issueDateLabel,
                          420D,
                          60D,
                          "F1",
                          10D,
                          "475569")
            page.DrawText("Current Class Schedule",
                          44D,
                          94D,
                          "F2",
                          15D,
                          "0F172A")
            page.DrawLine(44D, 114D, 551D, 114D, "CBD5E1", 1D)

            Return page
        End Function

        Private Sub WriteEmptyScheduleState(page As PdfPageBuilder,
                                            top As Double)
            page.FillRectangle(44D, top, 507D, 66D, "FFF7ED")
            page.StrokeRectangle(44D, top, 507D, 66D, "FED7AA", 1D)
            page.DrawText("No teacher schedule is available for the current section yet.",
                          60D,
                          top + 22D,
                          "F2",
                          12D,
                          "9A3412")
            page.DrawText("The certificate was generated from the active enrollment record.",
                          60D,
                          top + 42D,
                          "F1",
                          10D,
                          "9A3412")
        End Sub

        Private Sub WriteScheduleEntry(page As PdfPageBuilder,
                                       top As Double,
                                       entry As CertificateScheduleEntry,
                                       index As Integer)
            Dim totalHeight As Double = MeasureScheduleEntryHeight(entry)
            Dim backgroundColor As String =
                If(index Mod 2 = 0, "FFFFFF", "F8FBFF")

            page.FillRectangle(44D, top, 507D, totalHeight, backgroundColor)
            page.StrokeRectangle(44D, top, 507D, totalHeight, "D7E3F4", 1D)
            page.FillRectangle(44D, top, 8D, totalHeight, "2563EB")

            page.DrawText(entry.DayLabel.ToUpperInvariant(),
                          66D,
                          top + 16D,
                          "F2",
                          11D,
                          "2563EB")
            page.DrawText(entry.SessionLabel,
                          170D,
                          top + 16D,
                          "F2",
                          11D,
                          "0F172A")

            Dim subjectLines As List(Of String) =
                WrapText(entry.SubjectLabel, 445D, 12D)
            Dim lineTop As Double = top + 38D

            For Each subjectLine As String In subjectLines
                page.DrawText(subjectLine,
                              66D,
                              lineTop,
                              "F2",
                              12D,
                              "111827")
                lineTop += 16D
            Next

            Dim detailLines As New List(Of String)()
            detailLines.AddRange(WrapText("Room: " & entry.RoomLabel, 445D, 10D))
            detailLines.AddRange(WrapText("Instructor: " & entry.TeacherLabel,
                                          445D,
                                          10D))

            For Each detailLine As String In detailLines
                page.DrawText(detailLine,
                              66D,
                              lineTop,
                              "F1",
                              10D,
                              "475569")
                lineTop += 14D
            Next
        End Sub

        Private Shared Function MeasureScheduleEntryHeight(entry As CertificateScheduleEntry) As Double
            Dim subjectLines As Integer =
                Math.Max(1, WrapText(If(entry Is Nothing, String.Empty, entry.SubjectLabel),
                                     445D,
                                     12D).Count)
            Dim detailLines As Integer =
                WrapText("Room: " & If(entry Is Nothing, String.Empty, entry.RoomLabel),
                         445D,
                         10D).Count +
                WrapText("Instructor: " & If(entry Is Nothing,
                                             String.Empty,
                                             entry.TeacherLabel),
                         445D,
                         10D).Count

            Return 38D + (subjectLines * 16D) + (detailLines * 14D)
        End Function

        Private Sub WriteDetailBlock(page As PdfPageBuilder,
                                     left As Double,
                                     top As Double,
                                     label As String,
                                     value As String)
            page.DrawText(label,
                          left,
                          top,
                          "F1",
                          9D,
                          "64748B")

            Dim valueLines As List(Of String) = WrapText(value, 215D, 11.5D)
            Dim currentTop As Double = top + 14D

            For Each valueLine As String In valueLines
                page.DrawText(valueLine,
                              left,
                              currentTop,
                              "F2",
                              11.5D,
                              "0F172A")
                currentTop += 14D
            Next
        End Sub

        Private Sub AddFooter(page As PdfPageBuilder,
                              issueDateLabel As String)
            page.DrawLine(34D, 790D, 561D, 790D, "000000", 0.7D)
            page.DrawText("Generated on " & issueDateLabel,
                          34D,
                          804D,
                          "F1",
                          9D,
                          "000000")
            page.DrawText("School Management System",
                          435D,
                          804D,
                          "F1",
                          9D,
                          "000000")
        End Sub

        Private Shared Function WrapText(text As String,
                                         maxWidth As Double,
                                         fontSize As Double) As List(Of String)
            Dim lines As New List(Of String)()
            Dim normalizedText As String = If(text, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedText) Then
                lines.Add(String.Empty)
                Return lines
            End If

            Dim paragraphs As String() =
                normalizedText.Replace(vbCrLf, vbLf).
                    Replace(vbCr, vbLf).
                    Split(New String() {vbLf},
                          StringSplitOptions.None)

            For Each paragraph As String In paragraphs
                Dim normalizedParagraph As String = If(paragraph, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalizedParagraph) Then
                    lines.Add(String.Empty)
                    Continue For
                End If

                Dim currentLine As String = String.Empty
                Dim words As String() =
                    normalizedParagraph.Split(New Char() {" "c},
                                              StringSplitOptions.RemoveEmptyEntries)

                For Each word As String In words
                    Dim candidateLine As String =
                        If(String.IsNullOrWhiteSpace(currentLine),
                           word,
                           currentLine & " " & word)

                    If EstimateTextWidth(candidateLine, fontSize) <= maxWidth Then
                        currentLine = candidateLine
                        Continue For
                    End If

                    If Not String.IsNullOrWhiteSpace(currentLine) Then
                        lines.Add(currentLine)
                        currentLine = String.Empty
                    End If

                    If EstimateTextWidth(word, fontSize) <= maxWidth Then
                        currentLine = word
                        Continue For
                    End If

                    For Each fragment As String In BreakLongToken(word,
                                                                  maxWidth,
                                                                  fontSize)
                        If Not String.IsNullOrWhiteSpace(fragment) Then
                            lines.Add(fragment)
                        End If
                    Next
                Next

                If Not String.IsNullOrWhiteSpace(currentLine) Then
                    lines.Add(currentLine)
                End If
            Next

            If lines.Count = 0 Then
                lines.Add(String.Empty)
            End If

            Return lines
        End Function

        Private Shared Function BreakLongToken(token As String,
                                               maxWidth As Double,
                                               fontSize As Double) As List(Of String)
            Dim parts As New List(Of String)()
            Dim currentPart As New StringBuilder()

            For Each characterValue As Char In If(token, String.Empty)
                Dim candidate As String = currentPart.ToString() & characterValue
                If currentPart.Length > 0 AndAlso
                   EstimateTextWidth(candidate, fontSize) > maxWidth Then
                    parts.Add(currentPart.ToString())
                    currentPart.Clear()
                End If

                currentPart.Append(characterValue)
            Next

            If currentPart.Length > 0 Then
                parts.Add(currentPart.ToString())
            End If

            If parts.Count = 0 Then
                parts.Add(String.Empty)
            End If

            Return parts
        End Function

        Private Shared Function EstimateTextWidth(text As String,
                                                  fontSize As Double) As Double
            Dim width As Double = 0D

            For Each characterValue As Char In If(text, String.Empty)
                If Char.IsWhiteSpace(characterValue) Then
                    width += fontSize * 0.28D
                ElseIf Char.IsUpper(characterValue) Then
                    width += fontSize * 0.63D
                ElseIf Char.IsLower(characterValue) Then
                    width += fontSize * 0.55D
                ElseIf Char.IsDigit(characterValue) Then
                    width += fontSize * 0.56D
                Else
                    width += fontSize * 0.38D
                End If
            Next

            Return width
        End Function

        Private Shared Function LoadSchoolLogoImage() As PdfImageResource
            Try
                Dim resourceUri As New Uri("pack://application:,,,/Resources/Images/prmsu-logo.png")
                Dim resourceInfo = Application.GetResourceStream(resourceUri)
                If resourceInfo Is Nothing OrElse resourceInfo.Stream Is Nothing Then
                    Return Nothing
                End If

                Using resourceStream As Stream = resourceInfo.Stream
                    Dim decoder As BitmapDecoder =
                        BitmapDecoder.Create(resourceStream,
                                             BitmapCreateOptions.PreservePixelFormat,
                                             BitmapCacheOption.OnLoad)
                    If decoder Is Nothing OrElse decoder.Frames.Count = 0 Then
                        Return Nothing
                    End If

                    Dim source As BitmapSource = decoder.Frames(0)
                    Dim pixelWidth As Integer = Math.Max(1, source.PixelWidth)
                    Dim pixelHeight As Integer = Math.Max(1, source.PixelHeight)
                    Dim drawingVisual As New DrawingVisual()

                    Using drawingContext = drawingVisual.RenderOpen()
                        drawingContext.DrawRectangle(Brushes.White,
                                                     Nothing,
                                                     New Rect(0,
                                                              0,
                                                              pixelWidth,
                                                              pixelHeight))
                        drawingContext.DrawImage(source,
                                                 New Rect(0,
                                                          0,
                                                          pixelWidth,
                                                          pixelHeight))
                    End Using

                    Dim renderedBitmap As New RenderTargetBitmap(pixelWidth,
                                                                 pixelHeight,
                                                                 96D,
                                                                 96D,
                                                                 PixelFormats.Pbgra32)
                    renderedBitmap.Render(drawingVisual)

                    Dim encoder As New JpegBitmapEncoder() With {
                        .QualityLevel = 92
                    }
                    encoder.Frames.Add(BitmapFrame.Create(renderedBitmap))

                    Using imageStream As New MemoryStream()
                        encoder.Save(imageStream)
                        Return New PdfImageResource() With {
                            .ResourceName = "Im1",
                            .Width = pixelWidth,
                            .Height = pixelHeight,
                            .ImageBytes = imageStream.ToArray()
                        }
                    End Using
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub WritePdfDocument(filePath As String,
                                            pageContents As IList(Of String),
                                            Optional logoImage As PdfImageResource = Nothing)
            Dim resolvedPageContents As IList(Of String) =
                If(pageContents, New List(Of String)())
            If resolvedPageContents.Count = 0 Then
                Throw New IOException("No certificate content was generated.")
            End If

            Dim pageCount As Integer = resolvedPageContents.Count
            Dim firstFontObjectNumber As Integer = 3 + (pageCount * 2)
            Dim fontRegularObjectNumber As Integer = firstFontObjectNumber
            Dim fontBoldObjectNumber As Integer = firstFontObjectNumber + 1
            Dim imageObjectNumber As Integer = 0
            If logoImage IsNot Nothing AndAlso
               logoImage.ImageBytes IsNot Nothing AndAlso
               logoImage.ImageBytes.Length > 0 Then
                imageObjectNumber = firstFontObjectNumber + 2
            End If

            Dim objects As New List(Of Byte())()

            objects.Add(EncodeAscii("<< /Type /Catalog /Pages 2 0 R >>"))
            objects.Add(EncodeAscii(BuildPagesObject(pageCount)))

            For pageIndex As Integer = 0 To pageCount - 1
                Dim contentObjectNumber As Integer = 4 + (pageIndex * 2)

                objects.Add(EncodeAscii(BuildPageObject(contentObjectNumber,
                                                        fontRegularObjectNumber,
                                                        fontBoldObjectNumber,
                                                        imageObjectNumber,
                                                        If(logoImage Is Nothing,
                                                           String.Empty,
                                                           logoImage.ResourceName))))
                objects.Add(BuildContentStreamObject(resolvedPageContents(pageIndex)))
            Next

            objects.Add(EncodeAscii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"))
            objects.Add(EncodeAscii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>"))
            If imageObjectNumber > 0 Then
                objects.Add(BuildImageObject(logoImage))
            End If

            Using stream As New FileStream(filePath,
                                           FileMode.Create,
                                           FileAccess.Write,
                                           FileShare.None)
                WriteAscii(stream, "%PDF-1.4" & vbLf)
                WriteAscii(stream, "%1234" & vbLf)

                Dim offsets As New List(Of Long)()
                offsets.Add(0L)

                For index As Integer = 0 To objects.Count - 1
                    offsets.Add(stream.Position)
                    WriteAscii(stream,
                               (index + 1).ToString(Invariant) &
                               " 0 obj" &
                               vbLf)
                    stream.Write(objects(index), 0, objects(index).Length)
                    WriteAscii(stream, vbLf & "endobj" & vbLf)
                Next

                Dim xrefPosition As Long = stream.Position
                WriteAscii(stream, "xref" & vbLf)
                WriteAscii(stream,
                           "0 " & (objects.Count + 1).ToString(Invariant) & vbLf)
                WriteAscii(stream, "0000000000 65535 f " & vbLf)

                For index As Integer = 1 To offsets.Count - 1
                    WriteAscii(stream,
                               offsets(index).ToString("0000000000", Invariant) &
                               " 00000 n " &
                               vbLf)
                Next

                WriteAscii(stream, "trailer" & vbLf)
                WriteAscii(stream,
                           "<< /Size " &
                           (objects.Count + 1).ToString(Invariant) &
                           " /Root 1 0 R >>" &
                           vbLf)
                WriteAscii(stream, "startxref" & vbLf)
                WriteAscii(stream, xrefPosition.ToString(Invariant) & vbLf)
                WriteAscii(stream, "%%EOF")
            End Using
        End Sub

        Private Shared Function BuildPagesObject(pageCount As Integer) As String
            Dim kidReferences As New StringBuilder()

            For pageIndex As Integer = 0 To pageCount - 1
                If kidReferences.Length > 0 Then
                    kidReferences.Append(" "c)
                End If

                kidReferences.Append((3 + (pageIndex * 2)).ToString(Invariant))
                kidReferences.Append(" 0 R")
            Next

            Return "<< /Type /Pages /Count " &
                   pageCount.ToString(Invariant) &
                   " /Kids [ " &
                   kidReferences.ToString() &
                   " ] >>"
        End Function

        Private Shared Function BuildPageObject(contentObjectNumber As Integer,
                                                fontRegularObjectNumber As Integer,
                                                fontBoldObjectNumber As Integer,
                                                imageObjectNumber As Integer,
                                                imageResourceName As String) As String
            Dim xObjectSegment As String = String.Empty
            If imageObjectNumber > 0 AndAlso
               Not String.IsNullOrWhiteSpace(imageResourceName) Then
                xObjectSegment =
                    " /XObject << /" &
                    imageResourceName &
                    " " &
                    imageObjectNumber.ToString(Invariant) &
                    " 0 R >>"
            End If

            Return "<< /Type /Page /Parent 2 0 R " &
                   "/MediaBox [0 0 595 842] " &
                   "/Resources << /Font << /F1 " &
                   fontRegularObjectNumber.ToString(Invariant) &
                   " 0 R /F2 " &
                   fontBoldObjectNumber.ToString(Invariant) &
                   " 0 R >>" &
                   xObjectSegment &
                   " >> " &
                   "/Contents " &
                   contentObjectNumber.ToString(Invariant) &
                   " 0 R >>"
        End Function

        Private Shared Function BuildContentStreamObject(content As String) As Byte()
            Dim contentBytes As Byte() = EncodeAscii(If(content, String.Empty))
            Dim prefix As Byte() =
                EncodeAscii("<< /Length " &
                            contentBytes.Length.ToString(Invariant) &
                            " >>" &
                            vbLf &
                            "stream" &
                            vbLf)
            Dim suffix As Byte() = EncodeAscii(vbLf & "endstream")
            Dim objectBytes As Byte() =
                New Byte(prefix.Length + contentBytes.Length + suffix.Length - 1) {}

            Buffer.BlockCopy(prefix, 0, objectBytes, 0, prefix.Length)
            Buffer.BlockCopy(contentBytes,
                             0,
                             objectBytes,
                             prefix.Length,
                             contentBytes.Length)
            Buffer.BlockCopy(suffix,
                             0,
                             objectBytes,
                             prefix.Length + contentBytes.Length,
                             suffix.Length)

            Return objectBytes
        End Function

        Private Shared Function BuildImageObject(imageResource As PdfImageResource) As Byte()
            Dim imageBytes As Byte() = If(imageResource Is Nothing,
                                          Array.Empty(Of Byte)(),
                                          If(imageResource.ImageBytes,
                                             Array.Empty(Of Byte)()))
            Dim prefix As Byte() =
                EncodeAscii("<< /Type /XObject " &
                            "/Subtype /Image " &
                            "/Width " &
                            Math.Max(1, If(imageResource Is Nothing, 0, imageResource.Width)).
                                ToString(Invariant) &
                            " /Height " &
                            Math.Max(1, If(imageResource Is Nothing, 0, imageResource.Height)).
                                ToString(Invariant) &
                            " /ColorSpace /DeviceRGB " &
                            "/BitsPerComponent 8 " &
                            "/Filter /DCTDecode " &
                            "/Length " &
                            imageBytes.Length.ToString(Invariant) &
                            " >>" &
                            vbLf &
                            "stream" &
                            vbLf)
            Dim suffix As Byte() = EncodeAscii(vbLf & "endstream")
            Dim objectBytes As Byte() =
                New Byte(prefix.Length + imageBytes.Length + suffix.Length - 1) {}

            Buffer.BlockCopy(prefix, 0, objectBytes, 0, prefix.Length)
            Buffer.BlockCopy(imageBytes,
                             0,
                             objectBytes,
                             prefix.Length,
                             imageBytes.Length)
            Buffer.BlockCopy(suffix,
                             0,
                             objectBytes,
                             prefix.Length + imageBytes.Length,
                             suffix.Length)

            Return objectBytes
        End Function

        Private Shared Sub WriteAscii(stream As Stream,
                                      value As String)
            Dim bytes As Byte() = EncodeAscii(value)
            stream.Write(bytes, 0, bytes.Length)
        End Sub

        Private Shared Function EncodeAscii(value As String) As Byte()
            Return Encoding.ASCII.GetBytes(If(value, String.Empty))
        End Function

        Private Shared Function BuildSubjectLabel(record As TeacherScheduleRecord) As String
            Dim subjectCode As String = NormalizeValue(record.SubjectCode)
            Dim subjectName As String = NormalizeValue(record.SubjectName)

            If subjectCode <> "--" AndAlso
               subjectName <> "--" AndAlso
               Not String.Equals(subjectCode,
                                 subjectName,
                                 StringComparison.OrdinalIgnoreCase) Then
                Return subjectCode & " - " & subjectName
            End If

            If subjectCode <> "--" Then
                Return subjectCode
            End If

            Return subjectName
        End Function

        Private Shared Function BuildTeacherLabel(record As TeacherScheduleRecord) As String
            Dim teacherName As String = NormalizeValue(record.TeacherName, String.Empty)
            If teacherName <> String.Empty Then
                Return teacherName
            End If

            Return NormalizeValue(record.TeacherId, "Teacher TBA")
        End Function

        Private Shared Function BuildTimetableSectionYearLabel(student As StudentRecord) As String
            Return StudentScheduleHelper.BuildStudentSectionLabel(student, String.Empty)
        End Function

        Private Shared Function NormalizeValue(value As String,
                                               Optional fallbackValue As String = "--") As String
            Dim normalizedValue As String = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Return fallbackValue
            End If

            Return normalizedValue
        End Function
    End Class
End Namespace
