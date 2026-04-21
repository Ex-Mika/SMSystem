Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.Json
Imports System.Windows.Documents
Imports System.Windows.Media
Imports School_Management_System.Backend.Common
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

    Private Const TotalReportOptions As Integer = 6

    Private _searchTerm As String = String.Empty
    Private ReadOnly _courseManagementService As New CourseManagementService()
    Private ReadOnly _departmentManagementService As New DepartmentManagementService()
    Private ReadOnly _subjectManagementService As New SubjectManagementService()
    Private ReadOnly _teacherScheduleManagementService As New TeacherScheduleManagementService()
    Private ReadOnly _teacherManagementService As New TeacherManagementService()

    Private ReadOnly _studentsStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "students.json")
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
                StudentScheduleHelper.BuildYearLevelValue(record.YearLevel),
                NormalizeValue(record.Course),
                StudentScheduleHelper.BuildSectionValue(record.Section,
                                                       record.YearLevel)
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
        Dim rows As New List(Of String())()
        Dim result = _courseManagementService.GetCourses()

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Reports",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            UpdateStatusMessage("Unable to load courses data for printing.", StatusTone.Error)
            Return rows
        End If

        For Each record As CourseRecord In result.Data
            rows.Add(New String() {
                NormalizeValue(record.CourseCode),
                NormalizeValue(record.CourseName),
                NormalizeValue(record.DepartmentDisplayName),
                NormalizeValue(record.Units)
            })
        Next

        Return rows
    End Function

    Private Function GetDepartmentsReportRows() As List(Of String())
        Dim rows As New List(Of String())()
        Dim result = _departmentManagementService.GetDepartments()

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Reports",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            UpdateStatusMessage("Unable to load departments data for printing.", StatusTone.Error)
            Return rows
        End If

        For Each record As DepartmentRecord In result.Data
            rows.Add(New String() {
                NormalizeValue(record.DepartmentCode),
                NormalizeValue(record.DepartmentName),
                NormalizeValue(record.HeadName)
            })
        Next

        Return rows
    End Function

    Private Function GetSubjectsReportRows() As List(Of String())
        Dim rows As New List(Of String())()
        Dim result = _subjectManagementService.GetSubjects()

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Reports",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            UpdateStatusMessage("Unable to load subjects data for printing.", StatusTone.Error)
            Return rows
        End If

        For Each record As SubjectRecord In result.Data
            rows.Add(New String() {
                NormalizeValue(record.SubjectCode),
                NormalizeValue(record.SubjectName),
                NormalizeValue(record.Units),
                NormalizeValue(record.CourseDisplayName),
                NormalizeValue(record.YearLevel)
            })
        Next

        Return rows
    End Function

    Private Function GetSchedulingReportRows() As List(Of String())
        Dim rows As New List(Of String())()
        Dim result = _teacherScheduleManagementService.GetSchedules()

        If Not result.IsSuccess Then
            MessageBox.Show(result.Message,
                            "Reports",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
            UpdateStatusMessage("Unable to load schedules data for printing.", StatusTone.Error)
            Return rows
        End If

        For Each record As TeacherScheduleRecord In result.Data
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

        Dim exportHeaders As List(Of String) = BuildExportHeaders(headers)
        Dim exportRows As List(Of String()) = BuildExportRows(rows, exportHeaders.Count)
        Dim sheetName As String = BuildSafeSheetName(reportTitle)

        Using archive As ZipArchive = ZipFile.Open(filePath, ZipArchiveMode.Create)
            WriteZipEntry(archive, "[Content_Types].xml", BuildContentTypesXml())

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
            WriteZipEntry(archive,
                          "xl/worksheets/sheet1.xml",
                          BuildWorksheetXml(reportTitle, reportSubtitle, exportHeaders, exportRows))
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

    Private Function BuildContentTypesXml() As String
        Dim xml As New StringBuilder()
        xml.Append("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
        xml.Append("<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">")
        xml.Append("<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>")
        xml.Append("<Default Extension=""xml"" ContentType=""application/xml""/>")
        xml.Append("<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>")
        xml.Append("<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>")
        xml.Append("<Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>")
        xml.Append("</Types>")
        Return xml.ToString()
    End Function

    Private Function BuildStylesXml() As String
        Dim xml As New StringBuilder()
        xml.Append("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
        xml.Append("<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">")
        xml.Append("<fonts count=""3"">")
        xml.Append("<font><sz val=""11""/><color rgb=""FF000000""/><name val=""Segoe UI""/><family val=""2""/></font>")
        xml.Append("<font><b/><sz val=""14""/><color rgb=""FF000000""/><name val=""Segoe UI Semibold""/><family val=""2""/></font>")
        xml.Append("<font><b/><sz val=""11""/><color rgb=""FF000000""/><name val=""Segoe UI Semibold""/><family val=""2""/></font>")
        xml.Append("</fonts>")
        xml.Append("<fills count=""2"">")
        xml.Append("<fill><patternFill patternType=""none""/></fill>")
        xml.Append("<fill><patternFill patternType=""gray125""/></fill>")
        xml.Append("</fills>")
        xml.Append("<borders count=""2"">")
        xml.Append("<border><left/><right/><top/><bottom/><diagonal/></border>")
        xml.Append("<border><left style=""thin""><color rgb=""FF000000""/></left><right style=""thin""><color rgb=""FF000000""/></right><top style=""thin""><color rgb=""FF000000""/></top><bottom style=""thin""><color rgb=""FF000000""/></bottom><diagonal/></border>")
        xml.Append("</borders>")
        xml.Append("<cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>")
        xml.Append("<cellXfs count=""4"">")
        xml.Append("<xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyFont=""1"" applyAlignment=""1""><alignment vertical=""center"" wrapText=""1""/></xf>")
        xml.Append("<xf numFmtId=""0"" fontId=""1"" fillId=""0"" borderId=""0"" xfId=""0"" applyFont=""1"" applyAlignment=""1""><alignment horizontal=""left"" vertical=""center"" wrapText=""1""/></xf>")
        xml.Append("<xf numFmtId=""0"" fontId=""2"" fillId=""0"" borderId=""1"" xfId=""0"" applyFont=""1"" applyBorder=""1"" applyAlignment=""1""><alignment horizontal=""left"" vertical=""center"" wrapText=""1""/></xf>")
        xml.Append("<xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyFont=""1"" applyBorder=""1"" applyAlignment=""1""><alignment vertical=""top"" wrapText=""1""/></xf>")
        xml.Append("</cellXfs>")
        xml.Append("<cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>")
        xml.Append("</styleSheet>")
        Return xml.ToString()
    End Function

    Private Function BuildWorksheetXml(reportTitle As String,
                                       reportSubtitle As String,
                                       headers As IList(Of String),
                                       rows As IList(Of String())) As String
        Dim totalColumns As Integer = Math.Max(1, headers.Count)
        Dim lastColumnName As String = GetColumnName(totalColumns)
        Dim hasData As Boolean = rows.Count > 0
        Dim lastRowIndex As Integer = If(hasData, WorksheetTableHeaderRowIndex + rows.Count, WorksheetFirstDataRowIndex)
        Dim body As New StringBuilder()

        AppendStyledTextRow(body,
                            WorksheetReportTitleRowIndex,
                            New String() {reportTitle},
                            1,
                            24)
        AppendStyledTextRow(body,
                            WorksheetReportInfoRowIndex,
                            New String() {BuildWorksheetInfoText(reportSubtitle, rows.Count)},
                            0,
                            18)
        AppendStyledTextRow(body,
                            WorksheetTableHeaderRowIndex,
                            headers,
                            2,
                            22,
                            totalColumns)

        If hasData Then
            Dim rowIndex As Integer = WorksheetFirstDataRowIndex
            For Each dataRow As String() In rows
                AppendStyledTextRow(body, rowIndex, dataRow, 3, 20, totalColumns)
                rowIndex += 1
            Next
        Else
            AppendStyledTextRow(body,
                                WorksheetFirstDataRowIndex,
                                New String() {"No records available."},
                                3,
                                20)
        End If

        Dim xml As New StringBuilder()
        xml.Append("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
        xml.Append("<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" ")
        xml.Append("xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">")
        xml.Append("<dimension ref=""").Append(BuildSheetDimension(lastColumnName, lastRowIndex)).Append("""/>")
        xml.Append("<sheetViews><sheetView tabSelected=""1"" workbookViewId=""0"">")
        xml.Append("<pane ySplit=""").Append(WorksheetFrozenRowCount.ToString()).Append(""" topLeftCell=""A")
        xml.Append((WorksheetFrozenRowCount + 1).ToString()).Append(""" activePane=""bottomLeft"" state=""frozen""/>")
        xml.Append("</sheetView></sheetViews>")
        xml.Append("<sheetFormatPr defaultRowHeight=""20""/>")
        xml.Append(BuildColumnsXml(headers, rows))
        xml.Append("<sheetData>").Append(body.ToString()).Append("</sheetData>")
        xml.Append(BuildMergeCellsXml(lastColumnName, Not hasData))
        xml.Append("<pageMargins left=""0.35"" right=""0.35"" top=""0.6"" bottom=""0.6"" header=""0.3"" footer=""0.3""/>")
        xml.Append("<pageSetup orientation=""landscape"" fitToWidth=""1"" fitToHeight=""0""/>")
        xml.Append("</worksheet>")
        Return xml.ToString()
    End Function

    Private Sub AppendStyledTextRow(builder As StringBuilder,
                                    rowIndex As Integer,
                                    values As IEnumerable(Of String),
                                    styleIndex As Integer,
                                    rowHeight As Double,
                                    Optional expectedColumnCount As Integer = 0)
        builder.Append("<row r=""").Append(rowIndex.ToString()).Append(""" ht=""")
        builder.Append(rowHeight.ToString("0.##", CultureInfo.InvariantCulture))
        builder.Append(""" customHeight=""1"">")

        Dim columnIndex As Integer = 1
        If values IsNot Nothing Then
            For Each value As String In values
                AppendInlineStringCell(builder, rowIndex, columnIndex, NormalizeValue(value), styleIndex)
                columnIndex += 1
            Next
        End If

        While expectedColumnCount > 0 AndAlso columnIndex <= expectedColumnCount
            AppendInlineStringCell(builder, rowIndex, columnIndex, "--", styleIndex)
            columnIndex += 1
        End While

        builder.Append("</row>")
    End Sub

    Private Sub AppendInlineStringCell(builder As StringBuilder,
                                       rowIndex As Integer,
                                       columnIndex As Integer,
                                       value As String,
                                       styleIndex As Integer)
        builder.Append("<c r=""").Append(GetCellReference(columnIndex, rowIndex)).Append(""" t=""inlineStr"" s=""")
        builder.Append(styleIndex.ToString()).Append("""><is><t>")
        builder.Append(EscapeXml(value))
        builder.Append("</t></is></c>")
    End Sub

    Private Function BuildColumnsXml(headers As IList(Of String), rows As IList(Of String())) As String
        Dim xml As New StringBuilder()
        xml.Append("<cols>")

        For columnIndex As Integer = 0 To headers.Count - 1
            xml.Append("<col min=""").Append((columnIndex + 1).ToString()).Append(""" max=""")
            xml.Append((columnIndex + 1).ToString()).Append(""" width=""")
            xml.Append(CalculateColumnWidth(headers, rows, columnIndex).ToString("0.##", CultureInfo.InvariantCulture))
            xml.Append(""" customWidth=""1""/>")
        Next

        xml.Append("</cols>")
        Return xml.ToString()
    End Function

    Private Function CalculateColumnWidth(headers As IList(Of String),
                                          rows As IList(Of String()),
                                          columnIndex As Integer) As Double
        Dim maxLength As Integer = 14

        If headers IsNot Nothing AndAlso columnIndex < headers.Count Then
            maxLength = Math.Max(maxLength, NormalizeValue(headers(columnIndex)).Length)
        End If

        If rows IsNot Nothing Then
            For Each row As String() In rows
                If row IsNot Nothing AndAlso columnIndex < row.Length Then
                    maxLength = Math.Max(maxLength, NormalizeValue(row(columnIndex)).Length)
                End If
            Next
        End If

        Dim calculatedWidth As Double = Math.Ceiling((maxLength * 1.12) + 2)
        Return Math.Max(14.0, Math.Min(34.0, calculatedWidth))
    End Function

    Private Function BuildMergeCellsXml(lastColumnName As String, includeEmptyStateRow As Boolean) As String
        Dim xml As New StringBuilder()
        Dim mergeCount As Integer = If(includeEmptyStateRow, 3, 2)
        xml.Append("<mergeCells count=""").Append(mergeCount.ToString()).Append(""">")
        xml.Append("<mergeCell ref=""A").Append(WorksheetReportTitleRowIndex.ToString()).Append(":")
        xml.Append(lastColumnName).Append(WorksheetReportTitleRowIndex.ToString()).Append("""/>")
        xml.Append("<mergeCell ref=""A").Append(WorksheetReportInfoRowIndex.ToString()).Append(":")
        xml.Append(lastColumnName).Append(WorksheetReportInfoRowIndex.ToString()).Append("""/>")

        If includeEmptyStateRow Then
            xml.Append("<mergeCell ref=""A").Append(WorksheetFirstDataRowIndex.ToString()).Append(":")
            xml.Append(lastColumnName).Append(WorksheetFirstDataRowIndex.ToString()).Append("""/>")
        End If

        xml.Append("</mergeCells>")
        Return xml.ToString()
    End Function

    Private Function BuildSheetDimension(lastColumnName As String, lastRowIndex As Integer) As String
        Return "A1:" & lastColumnName & lastRowIndex.ToString()
    End Function

    Private Function BuildWorksheetInfoText(reportSubtitle As String, recordCount As Integer) As String
        Return NormalizeValue(reportSubtitle) &
               " | Generated on " &
               DateTime.Now.ToString("MMMM d, yyyy h:mm tt") &
               " | Total records: " &
               recordCount.ToString()
    End Function

    Private Function BuildExportHeaders(headers As IList(Of String)) As List(Of String)
        Dim exportHeaders As New List(Of String)()

        If headers IsNot Nothing Then
            For Each header As String In headers
                exportHeaders.Add(NormalizeValue(header))
            Next
        End If

        If exportHeaders.Count = 0 Then
            exportHeaders.Add("Data")
        End If

        Return exportHeaders
    End Function

    Private Function BuildExportRows(rows As IList(Of String()), columnCount As Integer) As List(Of String())
        Dim exportRows As New List(Of String())()
        If rows Is Nothing Then
            Return exportRows
        End If

        For Each sourceRow As String() In rows
            Dim normalizedRow(columnCount - 1) As String

            For columnIndex As Integer = 0 To columnCount - 1
                If sourceRow IsNot Nothing AndAlso columnIndex < sourceRow.Length Then
                    normalizedRow(columnIndex) = NormalizeValue(sourceRow(columnIndex))
                Else
                    normalizedRow(columnIndex) = "--"
                End If
            Next

            exportRows.Add(normalizedRow)
        Next

        Return exportRows
    End Function

    Private Const WorksheetReportTitleRowIndex As Integer = 1
    Private Const WorksheetReportInfoRowIndex As Integer = 2
    Private Const WorksheetTableHeaderRowIndex As Integer = 4
    Private Const WorksheetFirstDataRowIndex As Integer = 5
    Private Const WorksheetFrozenRowCount As Integer = 4

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

    Private Function BuildSchedulingSubjectText(record As TeacherScheduleRecord) As String
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
