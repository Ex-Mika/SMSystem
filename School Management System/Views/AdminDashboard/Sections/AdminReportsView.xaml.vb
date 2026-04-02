Imports System.Collections.Generic
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.Json
Imports System.Windows.Documents
Imports System.Windows.Media
Imports Microsoft.Win32
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class AdminReportsView
    Private Enum StatusTone
        Neutral
        Success
        [Error]
    End Enum

    Private Class StudentStorageRecord
        Public Property StudentId As String
        Public Property FullName As String
        Public Property YearLevel As String
        Public Property Course As String
        Public Property Section As String
    End Class

    Private Class CourseStorageRecord
        Public Property CourseCode As String
        Public Property CourseTitle As String
        Public Property Department As String
        Public Property Units As String
    End Class

    Private Class DepartmentStorageRecord
        Public Property DepartmentId As String
        Public Property DepartmentName As String
        Public Property Head As String
    End Class

    Private Class SubjectStorageRecord
        Public Property SubjectCode As String
        Public Property SubjectName As String
        Public Property Units As String
        Public Property Course As String
        Public Property YearLevel As String
    End Class

    Private Class ScheduleStorageRecord
        Public Property TeacherId As String
        Public Property TeacherName As String
        Public Property Day As String
        Public Property Session As String
        Public Property SubjectCode As String
        Public Property SubjectName As String
        Public Property Room As String
    End Class

    Private Const TotalReportOptions As Integer = 6

    Private _searchTerm As String = String.Empty
    Private ReadOnly _teacherManagementService As New TeacherManagementService()

    Private ReadOnly _studentsStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "students.json")
    Private ReadOnly _coursesStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "courses.json")
    Private ReadOnly _departmentsStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "departments.json")
    Private ReadOnly _subjectsStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "subjects.json")
    Private ReadOnly _schedulesStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "professor-schedules.json")
    Private ReadOnly _jsonOptions As New JsonSerializerOptions() With {
        .WriteIndented = True
    }

    Public Sub New()
        InitializeComponent()
        ApplyReportsFilter()
        UpdateStatusMessage("Select a report option to print.", StatusTone.Neutral)
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyReportsFilter()
    End Sub

    Public Sub RefreshData()
        ApplyReportsFilter()
    End Sub

    Private Sub ApplyReportsFilter()
        Dim visibleCount As Integer = 0

        visibleCount += ApplyOptionVisibility(StudentsReportCard,
                                              "students report student id full name year level course section")
        visibleCount += ApplyOptionVisibility(TeachersReportCard,
                                              "teachers report teacher id full name department advisory")
        visibleCount += ApplyOptionVisibility(CoursesReportCard,
                                              "courses report course code course title department units")
        visibleCount += ApplyOptionVisibility(DepartmentsReportCard,
                                              "departments report department id department name head")
        visibleCount += ApplyOptionVisibility(SubjectsReportCard,
                                              "subjects report subject code subject name units course year level")
        visibleCount += ApplyOptionVisibility(SchedulingReportCard,
                                              "scheduling report schedule teacher day session subject room")

        If NoReportsMatchTextBlock IsNot Nothing Then
            NoReportsMatchTextBlock.Visibility = If(visibleCount = 0, Visibility.Visible, Visibility.Collapsed)
        End If

        UpdateReportsCount(visibleCount)
    End Sub

    Private Function ApplyOptionVisibility(optionCard As Border, optionSearchText As String) As Integer
        If optionCard Is Nothing Then
            Return 0
        End If

        Dim isVisible As Boolean
        If String.IsNullOrWhiteSpace(_searchTerm) Then
            isVisible = True
        Else
            isVisible = optionSearchText.IndexOf(_searchTerm, StringComparison.OrdinalIgnoreCase) >= 0
        End If

        optionCard.Visibility = If(isVisible, Visibility.Visible, Visibility.Collapsed)
        Return If(isVisible, 1, 0)
    End Function

    Private Sub UpdateReportsCount(visibleCount As Integer)
        If ReportsCountTextBlock Is Nothing Then
            Return
        End If

        Dim hasSearch As Boolean = Not String.IsNullOrWhiteSpace(_searchTerm)
        If hasSearch AndAlso visibleCount <> TotalReportOptions Then
            ReportsCountTextBlock.Text = visibleCount.ToString() & " of " & TotalReportOptions.ToString() & " report options"
        Else
            ReportsCountTextBlock.Text = TotalReportOptions.ToString() & " report options"
        End If
    End Sub

    Private Sub PrintStudentsReportButton_Click(sender As Object, e As RoutedEventArgs)
        PrintTabularReport("Students Report",
                           "Current students list",
                           New String() {"Student ID", "Full Name", "Year Level", "Course", "Section"},
                           GetStudentsReportRows())
    End Sub

    Private Sub ExportStudentsReportButton_Click(sender As Object, e As RoutedEventArgs)
        ExportTabularReport("Students Report",
                            "Current students list",
                            New String() {"Student ID", "Full Name", "Year Level", "Course", "Section"},
                            GetStudentsReportRows())
    End Sub

    Private Sub PrintTeachersReportButton_Click(sender As Object, e As RoutedEventArgs)
        PrintTabularReport("Teachers Report",
                           "Current teachers list",
                           New String() {"Teacher ID", "Full Name", "Department", "Advisory"},
                           GetTeachersReportRows())
    End Sub

    Private Sub ExportTeachersReportButton_Click(sender As Object, e As RoutedEventArgs)
        ExportTabularReport("Teachers Report",
                            "Current teachers list",
                            New String() {"Teacher ID", "Full Name", "Department", "Advisory"},
                            GetTeachersReportRows())
    End Sub

    Private Sub PrintCoursesReportButton_Click(sender As Object, e As RoutedEventArgs)
        PrintTabularReport("Courses Report",
                           "Current courses list",
                           New String() {"Course Code", "Course Title", "Department", "Units"},
                           GetCoursesReportRows())
    End Sub

    Private Sub ExportCoursesReportButton_Click(sender As Object, e As RoutedEventArgs)
        ExportTabularReport("Courses Report",
                            "Current courses list",
                            New String() {"Course Code", "Course Title", "Department", "Units"},
                            GetCoursesReportRows())
    End Sub

    Private Sub PrintDepartmentsReportButton_Click(sender As Object, e As RoutedEventArgs)
        PrintTabularReport("Departments Report",
                           "Current departments list",
                           New String() {"Department ID", "Department Name", "Head"},
                           GetDepartmentsReportRows())
    End Sub

    Private Sub ExportDepartmentsReportButton_Click(sender As Object, e As RoutedEventArgs)
        ExportTabularReport("Departments Report",
                            "Current departments list",
                            New String() {"Department ID", "Department Name", "Head"},
                            GetDepartmentsReportRows())
    End Sub

    Private Sub PrintSubjectsReportButton_Click(sender As Object, e As RoutedEventArgs)
        PrintTabularReport("Subjects Report",
                           "Current subjects list",
                           New String() {"Subject Code", "Subject Name", "Units", "Course", "Year Level"},
                           GetSubjectsReportRows())
    End Sub

    Private Sub ExportSubjectsReportButton_Click(sender As Object, e As RoutedEventArgs)
        ExportTabularReport("Subjects Report",
                            "Current subjects list",
                            New String() {"Subject Code", "Subject Name", "Units", "Course", "Year Level"},
                            GetSubjectsReportRows())
    End Sub

    Private Sub PrintSchedulingReportButton_Click(sender As Object, e As RoutedEventArgs)
        PrintTabularReport("Scheduling Report",
                           "Current professor schedules",
                           New String() {"Teacher ID", "Teacher Name", "Day", "Session", "Subject", "Room"},
                           GetSchedulingReportRows())
    End Sub

    Private Sub ExportSchedulingReportButton_Click(sender As Object, e As RoutedEventArgs)
        ExportTabularReport("Scheduling Report",
                            "Current professor schedules",
                            New String() {"Teacher ID", "Teacher Name", "Day", "Session", "Subject", "Room"},
                            GetSchedulingReportRows())
    End Sub

    Private Function GetStudentsReportRows() As List(Of String())
        Dim records As List(Of StudentStorageRecord) = ReadRecordsFromStorage(Of StudentStorageRecord)(_studentsStoragePath, "students")

        Dim rows As New List(Of String())()
        For Each record As StudentStorageRecord In records
            rows.Add(New String() {
                NormalizeValue(record.StudentId),
                NormalizeValue(record.FullName),
                NormalizeValue(record.YearLevel),
                NormalizeValue(record.Course),
                NormalizeValue(record.Section)
            })
        Next

        Return rows
    End Function

    Private Function GetTeachersReportRows() As List(Of String())
        Dim rows As New List(Of String())()
        Dim result = _teacherManagementService.GetTeachers()

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Reports",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            Return rows
        End If

        For Each record As TeacherRecord In result.Data
            rows.Add(New String() {
                NormalizeValue(record.EmployeeNumber),
                NormalizeValue(record.FullName),
                NormalizeValue(record.DepartmentDisplayName),
                NormalizeValue(record.AdvisorySection)
            })
        Next

        Return rows
    End Function

    Private Function GetCoursesReportRows() As List(Of String())
        Dim records As List(Of CourseStorageRecord) = ReadRecordsFromStorage(Of CourseStorageRecord)(_coursesStoragePath, "courses")

        Dim rows As New List(Of String())()
        For Each record As CourseStorageRecord In records
            rows.Add(New String() {
                NormalizeValue(record.CourseCode),
                NormalizeValue(record.CourseTitle),
                NormalizeValue(record.Department),
                NormalizeValue(record.Units)
            })
        Next

        Return rows
    End Function

    Private Function GetDepartmentsReportRows() As List(Of String())
        Dim records As List(Of DepartmentStorageRecord) = ReadRecordsFromStorage(Of DepartmentStorageRecord)(_departmentsStoragePath, "departments")

        Dim rows As New List(Of String())()
        For Each record As DepartmentStorageRecord In records
            rows.Add(New String() {
                NormalizeValue(record.DepartmentId),
                NormalizeValue(record.DepartmentName),
                NormalizeValue(record.Head)
            })
        Next

        Return rows
    End Function

    Private Function GetSubjectsReportRows() As List(Of String())
        Dim records As List(Of SubjectStorageRecord) = ReadRecordsFromStorage(Of SubjectStorageRecord)(_subjectsStoragePath, "subjects")

        Dim rows As New List(Of String())()
        For Each record As SubjectStorageRecord In records
            rows.Add(New String() {
                NormalizeValue(record.SubjectCode),
                NormalizeValue(record.SubjectName),
                NormalizeValue(record.Units),
                NormalizeValue(record.Course),
                NormalizeValue(record.YearLevel)
            })
        Next

        Return rows
    End Function

    Private Function GetSchedulingReportRows() As List(Of String())
        Dim records As List(Of ScheduleStorageRecord) = ReadRecordsFromStorage(Of ScheduleStorageRecord)(_schedulesStoragePath, "schedules")

        Dim rows As New List(Of String())()
        For Each record As ScheduleStorageRecord In records
            rows.Add(New String() {
                NormalizeValue(record.TeacherId),
                NormalizeValue(record.TeacherName),
                NormalizeValue(record.Day),
                NormalizeValue(record.Session),
                BuildSchedulingSubjectText(record),
                NormalizeValue(record.Room)
            })
        Next

        Return rows
    End Function

    Private Function ReadRecordsFromStorage(Of T)(storagePath As String, entityLabel As String) As List(Of T)
        If Not File.Exists(storagePath) Then
            Return New List(Of T)()
        End If

        Try
            Dim json As String = File.ReadAllText(storagePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return New List(Of T)()
            End If

            Dim records As List(Of T) = JsonSerializer.Deserialize(Of List(Of T))(json, _jsonOptions)
            If records Is Nothing Then
                Return New List(Of T)()
            End If

            Return records
        Catch ex As Exception
            MessageBox.Show("Unable to load " & entityLabel & " data." & Environment.NewLine & ex.Message,
                            "Reports",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            UpdateStatusMessage("Unable to load " & entityLabel & " data for printing.", StatusTone.Error)
            Return New List(Of T)()
        End Try
    End Function

    Private Sub PrintTabularReport(reportTitle As String,
                                   reportSubtitle As String,
                                   headers As IList(Of String),
                                   rows As IList(Of String()))
        Dim printDialog As New PrintDialog()
        Dim printAccepted As Boolean? = printDialog.ShowDialog()
        If printAccepted <> True Then
            UpdateStatusMessage("Print canceled.", StatusTone.Neutral)
            Return
        End If

        Try
            Dim document As FlowDocument = BuildReportDocument(reportTitle, reportSubtitle, headers, rows)
            document.PageHeight = printDialog.PrintableAreaHeight
            document.PageWidth = printDialog.PrintableAreaWidth
            document.ColumnWidth = Math.Max(100.0, printDialog.PrintableAreaWidth - document.PagePadding.Left - document.PagePadding.Right)

            Dim paginatorSource As IDocumentPaginatorSource = document
            printDialog.PrintDocument(paginatorSource.DocumentPaginator, reportTitle)
            UpdateStatusMessage(reportTitle & " sent to printer.", StatusTone.Success)
        Catch ex As Exception
            MessageBox.Show("Unable to print " & reportTitle & "." & Environment.NewLine & ex.Message,
                            "Reports",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            UpdateStatusMessage("Unable to print " & reportTitle & ".", StatusTone.Error)
        End Try
    End Sub

    Private Sub ExportTabularReport(reportTitle As String,
                                    reportSubtitle As String,
                                    headers As IList(Of String),
                                    rows As IList(Of String()))
        Dim saveDialog As New SaveFileDialog() With {
            .Title = "Export " & reportTitle,
            .Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            .DefaultExt = ".xlsx",
            .AddExtension = True,
            .OverwritePrompt = True,
            .FileName = BuildDefaultExcelFileName(reportTitle)
        }

        Dim accepted As Boolean? = saveDialog.ShowDialog()
        If accepted <> True Then
            UpdateStatusMessage("Export canceled.", StatusTone.Neutral)
            Return
        End If

        Try
            WriteTabularDataAsXlsx(saveDialog.FileName, reportTitle, reportSubtitle, headers, rows)
            UpdateStatusMessage(reportTitle & " exported to Excel.", StatusTone.Success)
        Catch ex As Exception
            MessageBox.Show("Unable to export " & reportTitle & " to Excel." & Environment.NewLine & ex.Message,
                            "Reports",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            UpdateStatusMessage("Unable to export " & reportTitle & " to Excel.", StatusTone.Error)
        End Try
    End Sub

    Private Function BuildDefaultExcelFileName(reportTitle As String) As String
        Dim safeTitle As String = BuildSafeFileNameComponent(reportTitle)
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd-HHmm")
        Return safeTitle & "-" & timestamp & ".xlsx"
    End Function

    Private Function BuildSafeFileNameComponent(value As String) As String
        Dim candidate As String = NormalizeValue(value)
        If candidate = "--" Then
            candidate = "Report"
        End If

        For Each invalid As Char In Path.GetInvalidFileNameChars()
            candidate = candidate.Replace(invalid, "_"c)
        Next

        Return candidate
    End Function

    Private Sub WriteTabularDataAsXlsx(filePath As String,
                                       reportTitle As String,
                                       reportSubtitle As String,
                                       headers As IList(Of String),
                                       rows As IList(Of String()))
        Dim targetDirectory As String = Path.GetDirectoryName(filePath)
        If Not String.IsNullOrWhiteSpace(targetDirectory) Then
            Directory.CreateDirectory(targetDirectory)
        End If

        If File.Exists(filePath) Then
            File.Delete(filePath)
        End If

        Dim sheetName As String = BuildSafeSheetName(reportTitle)

        Using archive As ZipArchive = ZipFile.Open(filePath, ZipArchiveMode.Create)
            WriteZipEntry(archive,
                          "[Content_Types].xml",
                          "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" &
                          "<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">" &
                          "<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>" &
                          "<Default Extension=""xml"" ContentType=""application/xml""/>" &
                          "<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>" &
                          "<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>" &
                          "<Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>" &
                          "</Types>")

            WriteZipEntry(archive,
                          "_rels/.rels",
                          "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" &
                          "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">" &
                          "<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>" &
                          "</Relationships>")

            WriteZipEntry(archive,
                          "xl/workbook.xml",
                          "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" &
                          "<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" " &
                          "xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">" &
                          "<sheets><sheet name=""" & EscapeXml(sheetName) & """ sheetId=""1"" r:id=""rId1""/></sheets>" &
                          "</workbook>")

            WriteZipEntry(archive,
                          "xl/_rels/workbook.xml.rels",
                          "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" &
                          "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">" &
                          "<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>" &
                          "<Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>" &
                          "</Relationships>")

            WriteZipEntry(archive, "xl/styles.xml", BuildStylesXml())
            WriteZipEntry(archive, "xl/worksheets/sheet1.xml", BuildWorksheetXml(reportTitle, reportSubtitle, headers, rows))
        End Using
    End Sub

    Private Sub WriteZipEntry(archive As ZipArchive, entryPath As String, content As String)
        Dim entry As ZipArchiveEntry = archive.CreateEntry(entryPath, CompressionLevel.Optimal)
        Using stream As Stream = entry.Open()
            Using writer As New StreamWriter(stream, New UTF8Encoding(False))
                writer.Write(content)
            End Using
        End Using
    End Sub

    Private Function BuildStylesXml() As String
        Return "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" &
               "<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">" &
               "<fonts count=""2"">" &
               "<font><sz val=""11""/><color rgb=""FF1F2937""/><name val=""Segoe UI""/><family val=""2""/></font>" &
               "<font><b/><sz val=""11""/><color rgb=""FF1F2937""/><name val=""Segoe UI""/><family val=""2""/></font>" &
               "</fonts>" &
               "<fills count=""2""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill></fills>" &
               "<borders count=""1""><border><left/><right/><top/><bottom/><diagonal/></border></borders>" &
               "<cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>" &
               "<cellXfs count=""2"">" &
               "<xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyAlignment=""1""><alignment vertical=""top"" wrapText=""1""/></xf>" &
               "<xf numFmtId=""0"" fontId=""1"" fillId=""0"" borderId=""0"" xfId=""0"" applyAlignment=""1""><alignment vertical=""top"" wrapText=""1""/></xf>" &
               "</cellXfs>" &
               "<cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>" &
               "</styleSheet>"
    End Function

    Private Function BuildWorksheetXml(reportTitle As String,
                                       reportSubtitle As String,
                                       headers As IList(Of String),
                                       rows As IList(Of String())) As String
        Dim body As New StringBuilder()
        Dim rowIndex As Integer = 1

        AppendTextRow(body, rowIndex, New String() {reportTitle}, True)
        rowIndex += 1
        AppendTextRow(body, rowIndex, New String() {reportSubtitle}, False)
        rowIndex += 1
        AppendTextRow(body, rowIndex, New String() {"Generated on " & DateTime.Now.ToString("MMMM d, yyyy h:mm tt")}, False)
        rowIndex += 2

        AppendTextRow(body, rowIndex, headers, True)
        rowIndex += 1

        If rows Is Nothing OrElse rows.Count = 0 Then
            AppendTextRow(body, rowIndex, New String() {"No records available."}, False)
        Else
            For Each dataRow As String() In rows
                AppendTextRow(body, rowIndex, dataRow, False)
                rowIndex += 1
            Next
        End If

        Return "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" &
               "<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">" &
               "<sheetData>" & body.ToString() & "</sheetData>" &
               "</worksheet>"
    End Function

    Private Sub AppendTextRow(builder As StringBuilder, rowIndex As Integer, values As IEnumerable(Of String), isHeader As Boolean)
        builder.Append("<row r=""").Append(rowIndex.ToString()).Append(""">")

        Dim columnIndex As Integer = 1
        If values IsNot Nothing Then
            For Each value As String In values
                Dim styleIndex As Integer = If(isHeader, 1, 0)
                Dim cellReference As String = GetCellReference(columnIndex, rowIndex)
                builder.Append("<c r=""").Append(cellReference).Append(""" t=""inlineStr"" s=""").Append(styleIndex.ToString()).Append("""><is><t>")
                builder.Append(EscapeXml(NormalizeValue(value)))
                builder.Append("</t></is></c>")
                columnIndex += 1
            Next
        End If

        builder.Append("</row>")
    End Sub

    Private Function GetCellReference(columnIndex As Integer, rowIndex As Integer) As String
        Return GetColumnName(columnIndex) & rowIndex.ToString()
    End Function

    Private Function GetColumnName(columnIndex As Integer) As String
        If columnIndex <= 0 Then
            Return "A"
        End If

        Dim index As Integer = columnIndex
        Dim columnName As String = String.Empty

        While index > 0
            Dim remainder As Integer = (index - 1) Mod 26
            columnName = ChrW(65 + remainder) & columnName
            index = (index - 1) \ 26
        End While

        Return columnName
    End Function

    Private Function BuildSafeSheetName(reportTitle As String) As String
        Dim candidate As String = NormalizeValue(reportTitle)
        If candidate = "--" Then
            candidate = "Report"
        End If

        Dim invalidChars As String = "[]:*?/\"
        For Each invalidChar As Char In invalidChars
            candidate = candidate.Replace(invalidChar, "_"c)
        Next

        If candidate.Length > 31 Then
            candidate = candidate.Substring(0, 31)
        End If

        If String.IsNullOrWhiteSpace(candidate) Then
            Return "Report"
        End If

        Return candidate
    End Function

    Private Function EscapeXml(value As String) As String
        Dim safe As String = If(value, String.Empty)
        Return safe.
            Replace("&", "&amp;").
            Replace("<", "&lt;").
            Replace(">", "&gt;").
            Replace("""", "&quot;").
            Replace("'", "&apos;")
    End Function

    Private Function BuildReportDocument(reportTitle As String,
                                         reportSubtitle As String,
                                         headers As IList(Of String),
                                         rows As IList(Of String())) As FlowDocument
        Dim document As New FlowDocument() With {
            .PagePadding = New Thickness(40),
            .FontFamily = New FontFamily("Segoe UI"),
            .FontSize = 12
        }

        document.Blocks.Add(New Paragraph(New Run(reportTitle)) With {
            .FontSize = 24,
            .FontWeight = FontWeights.Bold,
            .Margin = New Thickness(0, 0, 0, 4)
        })
        document.Blocks.Add(New Paragraph(New Run(reportSubtitle)) With {
            .FontSize = 13,
            .Foreground = New SolidColorBrush(Color.FromRgb(&H5D, &H70, &H86)),
            .Margin = New Thickness(0, 0, 0, 2)
        })
        document.Blocks.Add(New Paragraph(New Run("Generated on " & DateTime.Now.ToString("MMMM d, yyyy h:mm tt"))) With {
            .FontSize = 11,
            .Foreground = New SolidColorBrush(Color.FromRgb(&H7A, &H8C, &HA0)),
            .Margin = New Thickness(0, 0, 0, 14)
        })

        If rows Is Nothing OrElse rows.Count = 0 Then
            document.Blocks.Add(New Paragraph(New Run("No records available.")) With {
                .FontSize = 13,
                .Foreground = New SolidColorBrush(Color.FromRgb(&H5D, &H70, &H86))
            })
            Return document
        End If

        Dim table As New Table() With {
            .CellSpacing = 0
        }

        For index As Integer = 0 To headers.Count - 1
            table.Columns.Add(New TableColumn())
        Next

        Dim body As New TableRowGroup()
        Dim headerRow As New TableRow()
        For Each header As String In headers
            Dim headerParagraph As New Paragraph(New Bold(New Run(header))) With {
                .Margin = New Thickness(6, 3, 6, 3),
                .TextAlignment = TextAlignment.Left
            }

            Dim headerCell As New TableCell(headerParagraph) With {
                .Background = New SolidColorBrush(Color.FromRgb(&HEF, &HF4, &HFA)),
                .BorderBrush = New SolidColorBrush(Color.FromRgb(&HD8, &HE1, &HEC)),
                .BorderThickness = New Thickness(0.5)
            }

            headerRow.Cells.Add(headerCell)
        Next
        body.Rows.Add(headerRow)

        For rowIndex As Integer = 0 To rows.Count - 1
            Dim sourceRow As String() = rows(rowIndex)
            Dim dataRow As New TableRow()

            For columnIndex As Integer = 0 To headers.Count - 1
                Dim cellValue As String = "--"
                If sourceRow IsNot Nothing AndAlso columnIndex < sourceRow.Length Then
                    cellValue = NormalizeValue(sourceRow(columnIndex))
                End If

                Dim valueParagraph As New Paragraph(New Run(cellValue)) With {
                    .Margin = New Thickness(6, 3, 6, 3)
                }

                Dim valueCell As New TableCell(valueParagraph) With {
                    .BorderBrush = New SolidColorBrush(Color.FromRgb(&HD8, &HE1, &HEC)),
                    .BorderThickness = New Thickness(0.5)
                }

                If rowIndex Mod 2 = 1 Then
                    valueCell.Background = New SolidColorBrush(Color.FromRgb(&HFA, &HFC, &HFF))
                End If

                dataRow.Cells.Add(valueCell)
            Next

            body.Rows.Add(dataRow)
        Next

        table.RowGroups.Add(body)
        document.Blocks.Add(table)

        Return document
    End Function

    Private Function BuildSchedulingSubjectText(record As ScheduleStorageRecord) As String
        If record Is Nothing Then
            Return "--"
        End If

        Dim subjectCode As String = NormalizeValue(record.SubjectCode)
        Dim subjectName As String = NormalizeValue(record.SubjectName)

        If subjectCode = "--" AndAlso subjectName = "--" Then
            Return "--"
        End If

        If subjectCode = "--" Then
            Return subjectName
        End If

        If subjectName = "--" Then
            Return subjectCode
        End If

        Return subjectCode & " - " & subjectName
    End Function

    Private Function NormalizeValue(value As String) As String
        Dim normalized As String = If(value, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalized) Then
            Return "--"
        End If

        Return normalized
    End Function

    Private Sub UpdateStatusMessage(message As String, tone As StatusTone)
        If ReportsActionStatusTextBlock Is Nothing Then
            Return
        End If

        ReportsActionStatusTextBlock.Text = message

        Dim fallbackBrush As Brush = New SolidColorBrush(Color.FromRgb(&H7A, &H8C, &HA0))
        Dim accentBrush As Brush = fallbackBrush

        Select Case tone
            Case StatusTone.Success
                accentBrush = TryCast(TryFindResource("DashboardSuccessBrush"), Brush)
                If accentBrush Is Nothing Then
                    accentBrush = New SolidColorBrush(Color.FromRgb(&H2A, &H7E, &H46))
                End If

            Case StatusTone.Error
                accentBrush = TryCast(TryFindResource("DashboardDangerBrush"), Brush)
                If accentBrush Is Nothing Then
                    accentBrush = New SolidColorBrush(Color.FromRgb(&HC2, &H3D, &H3D))
                End If

            Case Else
                accentBrush = TryCast(TryFindResource("DashboardTextMutedBrush"), Brush)
                If accentBrush Is Nothing Then
                    accentBrush = fallbackBrush
                End If
        End Select

        ReportsActionStatusTextBlock.Foreground = accentBrush
    End Sub
End Class
