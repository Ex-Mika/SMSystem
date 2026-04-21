Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models

Namespace Backend.Services
    Public Class TeacherStudentListExportService
        Private Class SubjectStudentListSection
            Public Property SubjectCode As String = String.Empty
            Public Property SubjectName As String = String.Empty
            Public Property SectionValue As String = String.Empty
            Public Property YearLevel As String = String.Empty
            Public Property Students As New List(Of StudentSubjectGradeRecord)()

            Public ReadOnly Property SubjectLabel As String
                Get
                    Dim normalizedCode As String = NormalizeText(SubjectCode)
                    Dim normalizedName As String = NormalizeText(SubjectName)

                    If normalizedCode <> String.Empty AndAlso
                       normalizedName <> String.Empty AndAlso
                       Not String.Equals(normalizedCode,
                                         normalizedName,
                                         StringComparison.OrdinalIgnoreCase) Then
                        Return normalizedCode & " - " & normalizedName
                    End If

                    If normalizedCode <> String.Empty Then
                        Return normalizedCode
                    End If

                    If normalizedName <> String.Empty Then
                        Return normalizedName
                    End If

                    Return "Untitled Subject"
                End Get
            End Property

            Public ReadOnly Property SectionLabel As String
                Get
                    Return StudentScheduleHelper.BuildCompactSectionValue(
                        NormalizeText(SectionValue),
                        NormalizeText(YearLevel),
                        "Section TBA")
                End Get
            End Property

            Public Function ToExportTarget() As TeacherStudentListExportTarget
                Return New TeacherStudentListExportTarget() With {
                    .SubjectCode = SubjectCode,
                    .SubjectName = SubjectName,
                    .SectionValue = SectionValue,
                    .YearLevel = YearLevel
                }
            End Function
        End Class

        Private Shared ReadOnly Invariant As CultureInfo =
            CultureInfo.InvariantCulture

        Private Const PageLeftMargin As Double = 42D
        Private Const PageRightMargin As Double = 42D
        Private Const PageTopMargin As Double = 42D
        Private Const PageBottomMargin As Double = 42D
        Private Const FooterReservedHeight As Double = 28D
        Private Const TableHeaderHeight As Double = 24D
        Private Const TableRowHeight As Double = 22D
        Private Const NumberColumnWidth As Double = 42D
        Private Const StudentNameColumnWidth As Double = 338D
        Private Const YearSectionColumnWidth As Double = 131D
        Private Const HeaderFillColor As String = "EFEFEF"
        Private Const LineColor As String = "000000"
        Private Const TextColor As String = "000000"

        Private ReadOnly _studentGradeManagementService As StudentGradeManagementService
        Private ReadOnly _teacherScheduleManagementService As TeacherScheduleManagementService

        Public Sub New()
            Me.New(New StudentGradeManagementService(),
                   New TeacherScheduleManagementService())
        End Sub

        Public Sub New(studentGradeManagementService As StudentGradeManagementService,
                       teacherScheduleManagementService As TeacherScheduleManagementService)
            _studentGradeManagementService =
                If(studentGradeManagementService, New StudentGradeManagementService())
            _teacherScheduleManagementService =
                If(teacherScheduleManagementService, New TeacherScheduleManagementService())
        End Sub

        Public Function GetExportTargets(teacherId As String) As ServiceResult(Of List(Of TeacherStudentListExportTarget))
            Dim subjectSectionsResult As ServiceResult(Of List(Of SubjectStudentListSection)) =
                LoadSubjectSections(teacherId)
            If subjectSectionsResult Is Nothing OrElse
               Not subjectSectionsResult.IsSuccess Then
                Return ServiceResult(Of List(Of TeacherStudentListExportTarget)).Failure(
                    If(subjectSectionsResult Is Nothing,
                       "Unable to load the teacher student list targets.",
                       subjectSectionsResult.Message))
            End If

            Dim exportTargets As New List(Of TeacherStudentListExportTarget)()
            For Each subjectSection As SubjectStudentListSection In subjectSectionsResult.Data
                If subjectSection Is Nothing Then
                    Continue For
                End If

                exportTargets.Add(subjectSection.ToExportTarget())
            Next

            Return ServiceResult(Of List(Of TeacherStudentListExportTarget)).Success(
                exportTargets,
                "Teacher student list targets loaded successfully.")
        End Function

        Public Function ExportTeacherStudentList(teacherId As String,
                                                 teacherName As String,
                                                 targetFilePath As String,
                                                 Optional exportTarget As TeacherStudentListExportTarget = Nothing) As ServiceResult(Of String)
            Dim normalizedTeacherId As String = NormalizeText(teacherId)
            If normalizedTeacherId = String.Empty Then
                Return ServiceResult(Of String).Failure("Teacher ID is required.")
            End If

            Dim normalizedFilePath As String = NormalizeText(targetFilePath)
            If normalizedFilePath = String.Empty Then
                Return ServiceResult(Of String).Failure("A PDF file path is required.")
            End If

            Dim subjectSectionsResult As ServiceResult(Of List(Of SubjectStudentListSection)) =
                LoadSubjectSections(normalizedTeacherId)
            If subjectSectionsResult Is Nothing OrElse
               Not subjectSectionsResult.IsSuccess Then
                Return ServiceResult(Of String).Failure(
                    If(subjectSectionsResult Is Nothing,
                       "Unable to load the teacher student list.",
                       subjectSectionsResult.Message))
            End If

            Dim subjectSections As List(Of SubjectStudentListSection) =
                FilterSubjectSections(subjectSectionsResult.Data,
                                      exportTarget)
            If subjectSections.Count = 0 Then
                Return ServiceResult(Of String).Failure(
                    "No assigned subject and section matched the selected export list.")
            End If

            Try
                Dim targetDirectory As String = Path.GetDirectoryName(normalizedFilePath)
                If NormalizeText(targetDirectory) <> String.Empty Then
                    Directory.CreateDirectory(targetDirectory)
                End If

                Dim pageContents As List(Of String) =
                    BuildStudentListPages(subjectSections,
                                          BuildTeacherLabel(teacherName,
                                                            normalizedTeacherId),
                                          normalizedTeacherId)
                WritePdfDocument(normalizedFilePath, pageContents)

                Return ServiceResult(Of String).Success(
                    normalizedFilePath,
                    "Teacher student list exported successfully.")
            Catch ex As Exception
                Return ServiceResult(Of String).Failure(
                    "Unable to export the teacher student list." &
                    Environment.NewLine &
                     ex.Message)
            End Try
        End Function

        Private Function LoadSubjectSections(teacherId As String) As ServiceResult(Of List(Of SubjectStudentListSection))
            Dim normalizedTeacherId As String = NormalizeText(teacherId)
            If normalizedTeacherId = String.Empty Then
                Return ServiceResult(Of List(Of SubjectStudentListSection)).Failure(
                    "Teacher ID is required.")
            End If

            Dim schedulesResult As ServiceResult(Of List(Of TeacherScheduleRecord)) =
                _teacherScheduleManagementService.GetSchedules()
            If schedulesResult Is Nothing OrElse Not schedulesResult.IsSuccess Then
                Return ServiceResult(Of List(Of SubjectStudentListSection)).Failure(
                    If(schedulesResult Is Nothing,
                       "Unable to load the teacher schedule list.",
                       schedulesResult.Message))
            End If

            Dim rosterResult As ServiceResult(Of List(Of StudentSubjectGradeRecord)) =
                _studentGradeManagementService.GetTeacherGradeRecords(normalizedTeacherId)
            If rosterResult Is Nothing OrElse Not rosterResult.IsSuccess Then
                Return ServiceResult(Of List(Of SubjectStudentListSection)).Failure(
                    If(rosterResult Is Nothing,
                       "Unable to load the teacher grading roster.",
                       rosterResult.Message))
            End If

            Dim subjectSections As List(Of SubjectStudentListSection) =
                BuildSubjectSections(schedulesResult.Data,
                                     rosterResult.Data,
                                     normalizedTeacherId)
            If subjectSections.Count = 0 Then
                Return ServiceResult(Of List(Of SubjectStudentListSection)).Failure(
                    "No assigned subjects were found for this teacher.")
            End If

            Return ServiceResult(Of List(Of SubjectStudentListSection)).Success(
                subjectSections,
                "Teacher student list data loaded successfully.")
        End Function

        Private Function BuildSubjectSections(scheduleRecords As IEnumerable(Of TeacherScheduleRecord),
                                              rosterRecords As IEnumerable(Of StudentSubjectGradeRecord),
                                              teacherId As String) As List(Of SubjectStudentListSection)
            Dim subjectSections As New List(Of SubjectStudentListSection)()

            If rosterRecords IsNot Nothing Then
                For Each rosterRecord As StudentSubjectGradeRecord In rosterRecords
                    If rosterRecord Is Nothing Then
                        Continue For
                    End If

                    Dim subjectSection As SubjectStudentListSection =
                        FindMatchingRosterSection(subjectSections,
                                                  rosterRecord)
                    If subjectSection Is Nothing Then
                        subjectSection = New SubjectStudentListSection() With {
                            .SubjectCode = NormalizeText(rosterRecord.SubjectCode),
                            .SubjectName = NormalizeText(rosterRecord.SubjectName),
                            .SectionValue = BuildSectionValue(rosterRecord),
                            .YearLevel = NormalizeText(rosterRecord.StudentYearLevel)
                        }
                        subjectSections.Add(subjectSection)
                    End If

                    subjectSection.Students.Add(rosterRecord)
                Next
            End If

            If scheduleRecords IsNot Nothing Then
                For Each scheduleRecord As TeacherScheduleRecord In scheduleRecords
                    If scheduleRecord Is Nothing OrElse
                       Not String.Equals(NormalizeText(scheduleRecord.TeacherId),
                                         teacherId,
                                         StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    If FindMatchingScheduleSection(subjectSections,
                                                  scheduleRecord) IsNot Nothing Then
                        Continue For
                    End If

                    subjectSections.Add(New SubjectStudentListSection() With {
                        .SubjectCode = NormalizeText(scheduleRecord.SubjectCode),
                        .SubjectName = NormalizeText(scheduleRecord.SubjectName),
                        .SectionValue = NormalizeText(scheduleRecord.Section)
                    })
                Next
            End If

            For Each subjectSection As SubjectStudentListSection In subjectSections
                subjectSection.Students.Sort(AddressOf CompareStudentRecords)
            Next

            subjectSections.Sort(Function(left, right)
                                     Dim subjectComparison As Integer =
                                         StringComparer.OrdinalIgnoreCase.Compare(
                                             If(left Is Nothing, String.Empty, left.SubjectLabel),
                                             If(right Is Nothing, String.Empty, right.SubjectLabel))
                                     If subjectComparison <> 0 Then
                                         Return subjectComparison
                                     End If

                                     Return StringComparer.OrdinalIgnoreCase.Compare(
                                         If(left Is Nothing, String.Empty, left.SectionLabel),
                                         If(right Is Nothing, String.Empty, right.SectionLabel))
                                 End Function)
            Return subjectSections
        End Function

        Private Function FindMatchingRosterSection(subjectSections As IEnumerable(Of SubjectStudentListSection),
                                                   rosterRecord As StudentSubjectGradeRecord) As SubjectStudentListSection
            If subjectSections Is Nothing OrElse rosterRecord Is Nothing Then
                Return Nothing
            End If

            For Each subjectSection As SubjectStudentListSection In subjectSections
                If subjectSection Is Nothing OrElse
                   Not IsSubjectMatch(subjectSection,
                                      rosterRecord.SubjectCode,
                                      rosterRecord.SubjectName) Then
                    Continue For
                End If

                If IsSectionMatch(subjectSection,
                                  BuildSectionValue(rosterRecord),
                                  NormalizeText(rosterRecord.StudentYearLevel)) Then
                    Return subjectSection
                End If
            Next

            Return Nothing
        End Function

        Private Function FindMatchingScheduleSection(subjectSections As IEnumerable(Of SubjectStudentListSection),
                                                     scheduleRecord As TeacherScheduleRecord) As SubjectStudentListSection
            If subjectSections Is Nothing OrElse scheduleRecord Is Nothing Then
                Return Nothing
            End If

            Dim normalizedScheduleSection As String = NormalizeText(scheduleRecord.Section)

            For Each subjectSection As SubjectStudentListSection In subjectSections
                If subjectSection Is Nothing OrElse
                   Not IsSubjectMatch(subjectSection,
                                      scheduleRecord.SubjectCode,
                                      scheduleRecord.SubjectName) Then
                    Continue For
                End If

                If normalizedScheduleSection = String.Empty Then
                    Return subjectSection
                End If

                If IsSectionMatch(subjectSection,
                                  normalizedScheduleSection,
                                  subjectSection.YearLevel) Then
                    Return subjectSection
                End If
            Next

            Return Nothing
        End Function

        Private Function IsSubjectMatch(subjectSection As SubjectStudentListSection,
                                        subjectCode As String,
                                        subjectName As String) As Boolean
            If subjectSection Is Nothing Then
                Return False
            End If

            Return String.Equals(BuildSubjectKey(subjectSection.SubjectCode,
                                                 subjectSection.SubjectName),
                                 BuildSubjectKey(subjectCode, subjectName),
                                 StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function IsSectionMatch(subjectSection As SubjectStudentListSection,
                                        sectionValue As String,
                                        yearLevel As String) As Boolean
            If subjectSection Is Nothing Then
                Return False
            End If

            Dim normalizedCandidateSection As String = NormalizeText(sectionValue)
            If normalizedCandidateSection = String.Empty Then
                Return NormalizeText(subjectSection.SectionValue) = String.Empty
            End If

            Return StudentScheduleHelper.SectionMatches(subjectSection.SectionValue,
                                                        normalizedCandidateSection,
                                                        NormalizeText(yearLevel),
                                                        NormalizeText(yearLevel))
        End Function

        Private Function FilterSubjectSections(subjectSections As IEnumerable(Of SubjectStudentListSection),
                                               exportTarget As TeacherStudentListExportTarget) As List(Of SubjectStudentListSection)
            Dim filteredSections As New List(Of SubjectStudentListSection)()
            If subjectSections Is Nothing Then
                Return filteredSections
            End If

            Dim normalizedSubjectCode As String =
                NormalizeText(If(exportTarget Is Nothing,
                                 String.Empty,
                                 exportTarget.SubjectCode))
            Dim normalizedSubjectName As String =
                NormalizeText(If(exportTarget Is Nothing,
                                 String.Empty,
                                 exportTarget.SubjectName))
            Dim normalizedSectionValue As String =
                NormalizeText(If(exportTarget Is Nothing,
                                 String.Empty,
                                 exportTarget.SectionValue))
            Dim normalizedYearLevel As String =
                NormalizeText(If(exportTarget Is Nothing,
                                 String.Empty,
                                 exportTarget.YearLevel))
            Dim hasSubjectFilter As Boolean =
                normalizedSubjectCode <> String.Empty OrElse
                normalizedSubjectName <> String.Empty
            Dim hasSectionFilter As Boolean = normalizedSectionValue <> String.Empty

            For Each subjectSection As SubjectStudentListSection In subjectSections
                If subjectSection Is Nothing Then
                    Continue For
                End If

                If hasSubjectFilter AndAlso
                   Not IsSubjectMatch(subjectSection,
                                      normalizedSubjectCode,
                                      normalizedSubjectName) Then
                    Continue For
                End If

                If hasSectionFilter AndAlso
                   Not IsSectionMatch(subjectSection,
                                      normalizedSectionValue,
                                      normalizedYearLevel) Then
                    Continue For
                End If

                filteredSections.Add(subjectSection)
            Next

            Return filteredSections
        End Function

        Private Function BuildStudentListPages(subjectSections As IEnumerable(Of SubjectStudentListSection),
                                               teacherLabel As String,
                                               teacherId As String) As List(Of String)
            Dim pageContents As New List(Of String)()
            Dim generatedLabel As String =
                DateTime.Now.ToString("MMMM d, yyyy h:mm tt", Invariant)

            For Each subjectSection As SubjectStudentListSection In subjectSections
                If subjectSection Is Nothing Then
                    Continue For
                End If

                Dim currentStudentIndex As Integer = 0
                Dim hasStudents As Boolean = subjectSection.Students.Count > 0

                Do
                    Dim page As New PdfPageBuilder()
                    currentStudentIndex =
                        DrawSubjectPage(page,
                                        subjectSection,
                                        teacherLabel,
                                        teacherId,
                                        generatedLabel,
                                        currentStudentIndex)
                    pageContents.Add(page.BuildContent())

                    If Not hasStudents OrElse
                       currentStudentIndex >= subjectSection.Students.Count Then
                        Exit Do
                    End If
                Loop
            Next

            Return pageContents
        End Function

        Private Function DrawSubjectPage(page As PdfPageBuilder,
                                         subjectSection As SubjectStudentListSection,
                                         teacherLabel As String,
                                         teacherId As String,
                                         generatedLabel As String,
                                         startStudentIndex As Integer) As Integer
            Dim contentWidth As Double =
                PdfPageBuilder.PageWidth - PageLeftMargin - PageRightMargin
            Dim currentTop As Double = PageTopMargin
            Dim continuationSuffix As String =
                If(startStudentIndex > 0, " (continued)", String.Empty)

            page.DrawText("Teacher Student List",
                          PageLeftMargin,
                          currentTop,
                          "F2",
                          22D,
                          TextColor)
            currentTop += 26D

            page.DrawText("Black-and-white print roster grouped by handled subject and section.",
                          PageLeftMargin,
                          currentTop,
                          "F1",
                          10D,
                          TextColor)
            currentTop += 16D

            page.DrawLine(PageLeftMargin,
                          currentTop,
                          PageLeftMargin + contentWidth,
                          currentTop,
                          LineColor,
                          1D)
            currentTop += 12D

            currentTop = DrawWrappedLabelLine(page,
                                              "Teacher: " & NormalizeText(teacherLabel,
                                                                          teacherId),
                                              PageLeftMargin,
                                              currentTop,
                                              contentWidth,
                                              "F1",
                                              10D) + 2D

            currentTop = DrawWrappedLabelLine(page,
                                              "Generated: " & NormalizeText(generatedLabel,
                                                                            "--"),
                                              PageLeftMargin,
                                              currentTop,
                                              contentWidth,
                                              "F1",
                                              10D) + 8D

            currentTop = DrawWrappedLabelLine(page,
                                              "Subject: " &
                                              subjectSection.SubjectLabel &
                                              continuationSuffix,
                                              PageLeftMargin,
                                              currentTop,
                                              contentWidth,
                                              "F2",
                                              14D) + 2D

            currentTop = DrawWrappedLabelLine(page,
                                              "Section: " &
                                              subjectSection.SectionLabel,
                                              PageLeftMargin,
                                              currentTop,
                                              contentWidth,
                                              "F1",
                                              10D) + 4D

            page.DrawText("Total Students: " &
                          subjectSection.Students.Count.ToString(Invariant),
                          PageLeftMargin,
                          currentTop,
                          "F1",
                          10D,
                          TextColor)
            currentTop += 22D

            DrawTableHeader(page, currentTop)
            currentTop += TableHeaderHeight

            If subjectSection.Students.Count = 0 Then
                DrawEmptyState(page,
                               currentTop + 8D,
                               contentWidth)
                AddFooter(page,
                          teacherId,
                          generatedLabel)
                Return 0
            End If

            Dim maxTableBottom As Double =
                PdfPageBuilder.PageHeight - PageBottomMargin - FooterReservedHeight
            Dim currentStudentIndex As Integer = startStudentIndex

            While currentStudentIndex < subjectSection.Students.Count AndAlso
                  currentTop + TableRowHeight <= maxTableBottom
                DrawStudentRow(page,
                               currentTop,
                               currentStudentIndex + 1,
                               subjectSection.Students(currentStudentIndex))
                currentTop += TableRowHeight
                currentStudentIndex += 1
            End While

            AddFooter(page,
                      teacherId,
                      generatedLabel)
            Return currentStudentIndex
        End Function

        Private Sub DrawTableHeader(page As PdfPageBuilder,
                                    top As Double)
            Dim totalWidth As Double =
                NumberColumnWidth + StudentNameColumnWidth + YearSectionColumnWidth
            Dim secondColumnLeft As Double = PageLeftMargin + NumberColumnWidth
            Dim thirdColumnLeft As Double = secondColumnLeft + StudentNameColumnWidth

            page.FillRectangle(PageLeftMargin,
                               top,
                               totalWidth,
                               TableHeaderHeight,
                               HeaderFillColor)
            page.StrokeRectangle(PageLeftMargin,
                                 top,
                                 totalWidth,
                                 TableHeaderHeight,
                                 LineColor,
                                 0.75D)
            page.DrawLine(secondColumnLeft,
                          top,
                          secondColumnLeft,
                          top + TableHeaderHeight,
                          LineColor,
                          0.75D)
            page.DrawLine(thirdColumnLeft,
                          top,
                          thirdColumnLeft,
                          top + TableHeaderHeight,
                          LineColor,
                          0.75D)

            page.DrawText("No.",
                          PageLeftMargin + 6D,
                          top + 7D,
                          "F2",
                          9D,
                          TextColor)
            page.DrawText("Student Name",
                          secondColumnLeft + 6D,
                          top + 7D,
                          "F2",
                          9D,
                          TextColor)
            page.DrawText("Year + Section",
                          thirdColumnLeft + 6D,
                          top + 7D,
                          "F2",
                          9D,
                          TextColor)
        End Sub

        Private Sub DrawStudentRow(page As PdfPageBuilder,
                                   top As Double,
                                   rowNumber As Integer,
                                   record As StudentSubjectGradeRecord)
            Dim totalWidth As Double =
                NumberColumnWidth + StudentNameColumnWidth + YearSectionColumnWidth
            Dim secondColumnLeft As Double = PageLeftMargin + NumberColumnWidth
            Dim thirdColumnLeft As Double = secondColumnLeft + StudentNameColumnWidth

            page.StrokeRectangle(PageLeftMargin,
                                 top,
                                 totalWidth,
                                 TableRowHeight,
                                 LineColor,
                                 0.5D)
            page.DrawLine(secondColumnLeft,
                          top,
                          secondColumnLeft,
                          top + TableRowHeight,
                          LineColor,
                          0.5D)
            page.DrawLine(thirdColumnLeft,
                          top,
                          thirdColumnLeft,
                          top + TableRowHeight,
                          LineColor,
                          0.5D)

            page.DrawText(rowNumber.ToString(Invariant),
                          PageLeftMargin + 6D,
                          top + 7D,
                          "F1",
                          9D,
                          TextColor)
            page.DrawText(FitTextToWidth(NormalizeText(If(record Is Nothing,
                                                          String.Empty,
                                                          record.StudentName),
                                                       "--"),
                                         StudentNameColumnWidth - 12D,
                                         9D),
                          secondColumnLeft + 6D,
                          top + 7D,
                          "F1",
                          9D,
                          TextColor)
            page.DrawText(FitTextToWidth(BuildYearSectionLabel(record),
                                         YearSectionColumnWidth - 12D,
                                         9D),
                          thirdColumnLeft + 6D,
                          top + 7D,
                          "F1",
                          9D,
                          TextColor)
        End Sub

        Private Sub DrawEmptyState(page As PdfPageBuilder,
                                   top As Double,
                                   contentWidth As Double)
            Dim noteHeight As Double = 54D

            page.StrokeRectangle(PageLeftMargin,
                                 top,
                                 contentWidth,
                                 noteHeight,
                                 LineColor,
                                 0.75D)
            page.DrawText("No enrolled students were found for this assigned subject.",
                          PageLeftMargin + 10D,
                          top + 16D,
                          "F2",
                          11D,
                          TextColor)
            page.DrawText("The subject is still included so the teacher can keep a printable record of handled loads.",
                          PageLeftMargin + 10D,
                          top + 32D,
                          "F1",
                          9D,
                          TextColor)
        End Sub

        Private Sub AddFooter(page As PdfPageBuilder,
                              teacherId As String,
                              generatedLabel As String)
            Dim footerTop As Double = PdfPageBuilder.PageHeight - PageBottomMargin - 16D
            Dim contentWidth As Double =
                PdfPageBuilder.PageWidth - PageLeftMargin - PageRightMargin
            Dim footerText As String =
                "Teacher ID: " & NormalizeText(teacherId, "--") &
                " | Generated: " & NormalizeText(generatedLabel, "--")

            page.DrawLine(PageLeftMargin,
                          footerTop - 6D,
                          PageLeftMargin + contentWidth,
                          footerTop - 6D,
                          LineColor,
                          0.75D)
            page.DrawText(FitTextToWidth(footerText,
                                         contentWidth,
                                         8.5D),
                          PageLeftMargin,
                          footerTop,
                          "F1",
                          8.5D,
                          TextColor)
        End Sub

        Private Function DrawWrappedLabelLine(page As PdfPageBuilder,
                                              value As String,
                                              left As Double,
                                              top As Double,
                                              maxWidth As Double,
                                              fontResource As String,
                                              fontSize As Double) As Double
            Dim lines As List(Of String) = WrapText(value, maxWidth, fontSize)
            Dim currentTop As Double = top

            For Each lineValue As String In lines
                page.DrawText(lineValue,
                              left,
                              currentTop,
                              fontResource,
                              fontSize,
                              TextColor)
                currentTop += fontSize + 3D
            Next

            Return currentTop
        End Function

        Private Shared Function CompareStudentRecords(left As StudentSubjectGradeRecord,
                                                      right As StudentSubjectGradeRecord) As Integer
            Dim nameComparison As Integer =
                StringComparer.OrdinalIgnoreCase.Compare(
                    NormalizeText(If(left Is Nothing, String.Empty, left.StudentName)),
                    NormalizeText(If(right Is Nothing, String.Empty, right.StudentName)))
            If nameComparison <> 0 Then
                Return nameComparison
            End If

            Return StringComparer.OrdinalIgnoreCase.Compare(
                NormalizeText(If(left Is Nothing, String.Empty, left.StudentNumber)),
                NormalizeText(If(right Is Nothing, String.Empty, right.StudentNumber)))
        End Function

        Private Shared Function BuildYearSectionLabel(record As StudentSubjectGradeRecord) As String
            If record Is Nothing Then
                Return "--"
            End If

            Return StudentScheduleHelper.BuildCompactSectionValue(
                BuildSectionValue(record),
                NormalizeText(record.StudentYearLevel),
                "--")
        End Function

        Private Shared Function BuildSectionValue(record As StudentSubjectGradeRecord) As String
            If record Is Nothing Then
                Return String.Empty
            End If

            Dim sectionValue As String = NormalizeText(record.SectionName)
            If sectionValue <> String.Empty Then
                Return sectionValue
            End If

            Return NormalizeText(record.StudentSection)
        End Function

        Private Shared Function BuildTeacherLabel(teacherName As String,
                                                  teacherId As String) As String
            Dim normalizedTeacherName As String = NormalizeText(teacherName)
            Dim normalizedTeacherId As String = NormalizeText(teacherId)

            If normalizedTeacherName <> String.Empty AndAlso
               normalizedTeacherId <> String.Empty Then
                Return normalizedTeacherName & " (" & normalizedTeacherId & ")"
            End If

            If normalizedTeacherName <> String.Empty Then
                Return normalizedTeacherName
            End If

            Return NormalizeText(normalizedTeacherId, "Teacher")
        End Function

        Private Shared Function BuildSubjectKey(subjectCode As String,
                                                subjectName As String) As String
            Dim normalizedCode As String = NormalizeText(subjectCode)
            If normalizedCode <> String.Empty Then
                Return "CODE:" & normalizedCode.ToUpperInvariant()
            End If

            Dim normalizedName As String = NormalizeText(subjectName)
            If normalizedName <> String.Empty Then
                Return "NAME:" & normalizedName.ToUpperInvariant()
            End If

            Return String.Empty
        End Function

        Private Shared Function FitTextToWidth(value As String,
                                               maxWidth As Double,
                                               fontSize As Double) As String
            Dim normalizedValue As String = NormalizeText(value, "--")
            If EstimateTextWidth(normalizedValue, fontSize) <= maxWidth Then
                Return normalizedValue
            End If

            Dim candidate As String = normalizedValue
            Do While candidate.Length > 1
                candidate = candidate.Substring(0, candidate.Length - 1).TrimEnd()
                Dim ellipsisCandidate As String = candidate & "..."
                If EstimateTextWidth(ellipsisCandidate, fontSize) <= maxWidth Then
                    Return ellipsisCandidate
                End If
            Loop

            Return "..."
        End Function

        Private Shared Function WrapText(value As String,
                                         maxWidth As Double,
                                         fontSize As Double) As List(Of String)
            Dim lines As New List(Of String)()
            Dim normalizedValue As String = NormalizeText(value)

            If normalizedValue = String.Empty Then
                lines.Add(String.Empty)
                Return lines
            End If

            Dim paragraphs As String() =
                normalizedValue.Split({vbCrLf, vbLf},
                                      StringSplitOptions.None)

            For paragraphIndex As Integer = 0 To paragraphs.Length - 1
                Dim paragraph As String = paragraphs(paragraphIndex)
                Dim currentLine As String = String.Empty

                For Each token As String In paragraph.Split({" "c},
                                                            StringSplitOptions.RemoveEmptyEntries)
                    Dim fragments As List(Of String) =
                        BreakLongToken(token, maxWidth, fontSize)

                    For Each fragment As String In fragments
                        Dim candidate As String =
                            If(currentLine = String.Empty,
                               fragment,
                               currentLine & " " & fragment)

                        If currentLine <> String.Empty AndAlso
                           EstimateTextWidth(candidate, fontSize) > maxWidth Then
                            lines.Add(currentLine)
                            currentLine = fragment
                        Else
                            currentLine = candidate
                        End If
                    Next
                Next

                If currentLine <> String.Empty Then
                    lines.Add(currentLine)
                End If

                If paragraphIndex < paragraphs.Length - 1 AndAlso
                   currentLine = String.Empty Then
                    lines.Add(String.Empty)
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

                Dim normalizedFont As String = NormalizeText(fontResource)
                If normalizedFont = String.Empty Then
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
                    NormalizeText(colorHex).TrimStart("#"c)
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

        Private Shared Sub WritePdfDocument(filePath As String,
                                            pageContents As IList(Of String))
            Dim resolvedPageContents As IList(Of String) =
                If(pageContents, New List(Of String)())
            If resolvedPageContents.Count = 0 Then
                Throw New IOException("No PDF content was generated.")
            End If

            Dim pageCount As Integer = resolvedPageContents.Count
            Dim firstFontObjectNumber As Integer = 3 + (pageCount * 2)
            Dim fontRegularObjectNumber As Integer = firstFontObjectNumber
            Dim fontBoldObjectNumber As Integer = firstFontObjectNumber + 1
            Dim objects As New List(Of Byte())()

            objects.Add(EncodeAscii("<< /Type /Catalog /Pages 2 0 R >>"))
            objects.Add(EncodeAscii(BuildPagesObject(pageCount)))

            For pageIndex As Integer = 0 To pageCount - 1
                Dim contentObjectNumber As Integer = 4 + (pageIndex * 2)

                objects.Add(EncodeAscii(BuildPageObject(contentObjectNumber,
                                                        fontRegularObjectNumber,
                                                        fontBoldObjectNumber)))
                objects.Add(BuildContentStreamObject(resolvedPageContents(pageIndex)))
            Next

            objects.Add(EncodeAscii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"))
            objects.Add(EncodeAscii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>"))

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
                                                fontBoldObjectNumber As Integer) As String
            Return "<< /Type /Page /Parent 2 0 R " &
                   "/MediaBox [0 0 595 842] " &
                   "/Resources << /Font << /F1 " &
                   fontRegularObjectNumber.ToString(Invariant) &
                   " 0 R /F2 " &
                   fontBoldObjectNumber.ToString(Invariant) &
                   " 0 R >> >> " &
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

        Private Shared Sub WriteAscii(stream As Stream,
                                      value As String)
            Dim bytes As Byte() = EncodeAscii(value)
            stream.Write(bytes, 0, bytes.Length)
        End Sub

        Private Shared Function EncodeAscii(value As String) As Byte()
            Return Encoding.ASCII.GetBytes(If(value, String.Empty))
        End Function

        Private Shared Function NormalizeText(value As String,
                                              Optional fallbackValue As String = "") As String
            Dim normalizedValue As String = If(value, String.Empty).Trim()
            If normalizedValue = String.Empty Then
                Return fallbackValue
            End If

            Return normalizedValue
        End Function
    End Class
End Namespace
