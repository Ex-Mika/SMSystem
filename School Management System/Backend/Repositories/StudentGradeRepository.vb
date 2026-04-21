Imports System.Collections.Generic
Imports System.Globalization
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models

Namespace Backend.Repositories
    Public Class StudentGradeRepository
        Private Const TeacherRosterSelectSql As String =
            "SELECT " &
            "COALESCE(g.grade_id, 0) AS grade_id, " &
            "t.teacher_id AS teacher_record_id, " &
            "t.employee_number AS teacher_id, " &
            "TRIM(CONCAT_WS(' ', t.first_name, NULLIF(t.middle_name, ''), t.last_name)) AS teacher_name, " &
            "s.student_id AS student_record_id, " &
            "s.student_number, " &
            "TRIM(CONCAT_WS(' ', s.first_name, NULLIF(s.middle_name, ''), s.last_name)) AS student_name, " &
            "COALESCE(s.section_name, '') AS student_section, " &
            "COALESCE(CAST(s.year_level AS CHAR), '') AS student_year_level, " &
            "subj.subject_id, " &
            "COALESCE(subj.subject_code, '') AS subject_code, " &
            "COALESCE(subj.subject_title, '') AS subject_name, " &
            "CAST(subj.units AS CHAR) AS subject_units, " &
            "COALESCE(subj.year_level, '') AS subject_year_level, " &
            "COALESCE(a.section_name, '') AS scheduled_section_name, " &
            "COALESCE(NULLIF(g.section_name, ''), a.section_name, COALESCE(s.section_name, '')) AS section_name, " &
            "g.quiz_score, " &
            "g.project_score, " &
            "g.midterm_score, " &
            "g.final_exam_score, " &
            "g.final_grade, " &
            "COALESCE(g.remarks, '') AS remarks, " &
            "g.updated_at " &
            "FROM (" &
            " SELECT DISTINCT " &
            "   ts.teacher_id, " &
            "   TRIM(COALESCE(ts.section_name, '')) AS section_name, " &
            "   TRIM(COALESCE(ts.subject_code, '')) AS subject_code, " &
            "   TRIM(COALESCE(ts.subject_name, '')) AS subject_name " &
            " FROM teacher_schedules ts " &
            " WHERE COALESCE(TRIM(ts.subject_code), '') <> '' " &
            "    OR COALESCE(TRIM(ts.subject_name), '') <> ''" &
            ") a " &
            "INNER JOIN teachers t ON t.teacher_id = a.teacher_id " &
            "INNER JOIN subjects subj ON (" &
            "    (a.subject_code <> '' AND LOWER(subj.subject_code) = LOWER(a.subject_code)) " &
            " OR (a.subject_code = '' AND a.subject_name <> '' AND LOWER(subj.subject_title) = LOWER(a.subject_name))" &
            ") " &
            "INNER JOIN student_subject_enrollments e ON e.subject_id = subj.subject_id " &
            "INNER JOIN students s ON s.student_id = e.student_id " &
            "LEFT JOIN student_subject_grades g ON g.teacher_id = a.teacher_id " &
            " AND g.student_id = s.student_id " &
            " AND g.subject_id = subj.subject_id " &
            "WHERE t.employee_number = @teacherId "

        Private Const StudentGradesSelectSql As String =
            "SELECT " &
            "g.grade_id, " &
            "t.teacher_id AS teacher_record_id, " &
            "t.employee_number AS teacher_id, " &
            "TRIM(CONCAT_WS(' ', t.first_name, NULLIF(t.middle_name, ''), t.last_name)) AS teacher_name, " &
            "s.student_id AS student_record_id, " &
            "s.student_number, " &
            "TRIM(CONCAT_WS(' ', s.first_name, NULLIF(s.middle_name, ''), s.last_name)) AS student_name, " &
            "COALESCE(s.section_name, '') AS student_section, " &
            "COALESCE(CAST(s.year_level AS CHAR), '') AS student_year_level, " &
            "subj.subject_id, " &
            "COALESCE(subj.subject_code, '') AS subject_code, " &
            "COALESCE(subj.subject_title, '') AS subject_name, " &
            "CAST(subj.units AS CHAR) AS subject_units, " &
            "COALESCE(subj.year_level, '') AS subject_year_level, " &
            "COALESCE(g.section_name, COALESCE(s.section_name, '')) AS section_name, " &
            "g.quiz_score, " &
            "g.project_score, " &
            "g.midterm_score, " &
            "g.final_exam_score, " &
            "g.final_grade, " &
            "COALESCE(g.remarks, '') AS remarks, " &
            "g.updated_at " &
            "FROM student_subject_grades g " &
            "INNER JOIN teachers t ON t.teacher_id = g.teacher_id " &
            "INNER JOIN students s ON s.student_id = g.student_id " &
            "INNER JOIN subjects subj ON subj.subject_id = g.subject_id " &
            "WHERE s.student_number = @studentNumber "

        Public Function GetTeacherGradeRoster(teacherId As String) As List(Of StudentSubjectGradeRecord)
            Dim gradeRecords As New List(Of StudentSubjectGradeRecord)()
            If String.IsNullOrWhiteSpace(teacherId) Then
                Return gradeRecords
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    TeacherRosterSelectSql &
                    "ORDER BY subj.subject_code, section_name, s.last_name, s.first_name;",
                    connection)
                    command.Parameters.AddWithValue("@teacherId", teacherId.Trim())

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            If Not ShouldIncludeTeacherRosterRecord(reader) Then
                                Continue While
                            End If

                            UpsertTeacherRosterRecord(gradeRecords,
                                                      MapGradeRecord(reader))
                        End While
                    End Using
                End Using
            End Using

            Return gradeRecords
        End Function

        Public Function GetStudentGradesByStudentNumber(studentNumber As String) As List(Of StudentSubjectGradeRecord)
            Dim gradeRecords As New List(Of StudentSubjectGradeRecord)()
            If String.IsNullOrWhiteSpace(studentNumber) Then
                Return gradeRecords
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    StudentGradesSelectSql &
                    "ORDER BY subj.subject_code, subj.subject_title;",
                    connection)
                    command.Parameters.AddWithValue("@studentNumber", studentNumber.Trim())

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            gradeRecords.Add(MapGradeRecord(reader))
                        End While
                    End Using
                End Using
            End Using

            Return gradeRecords
        End Function

        Public Function SaveGrade(teacherRecordId As Integer,
                                  studentRecordId As Integer,
                                  subjectId As Integer,
                                  sectionName As String,
                                  quizScore As Decimal,
                                  projectScore As Decimal,
                                  midtermScore As Decimal,
                                  finalExamScore As Decimal,
                                  finalGrade As Decimal,
                                  remarks As String) As Boolean
            If teacherRecordId <= 0 OrElse
               studentRecordId <= 0 OrElse
               subjectId <= 0 Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "INSERT INTO student_subject_grades (" &
                    "teacher_id, student_id, subject_id, section_name, " &
                    "quiz_score, project_score, midterm_score, final_exam_score, final_grade, remarks" &
                    ") VALUES (" &
                    "@teacherRecordId, @studentRecordId, @subjectId, @sectionName, " &
                    "@quizScore, @projectScore, @midtermScore, @finalExamScore, @finalGrade, @remarks" &
                    ") ON DUPLICATE KEY UPDATE " &
                    "section_name = VALUES(section_name), " &
                    "quiz_score = VALUES(quiz_score), " &
                    "project_score = VALUES(project_score), " &
                    "midterm_score = VALUES(midterm_score), " &
                    "final_exam_score = VALUES(final_exam_score), " &
                    "final_grade = VALUES(final_grade), " &
                    "remarks = VALUES(remarks), " &
                    "updated_at = CURRENT_TIMESTAMP;",
                    connection)
                    command.Parameters.AddWithValue("@teacherRecordId", teacherRecordId)
                    command.Parameters.AddWithValue("@studentRecordId", studentRecordId)
                    command.Parameters.AddWithValue("@subjectId", subjectId)
                    command.Parameters.AddWithValue("@sectionName", NormalizeNullableValue(sectionName))
                    command.Parameters.AddWithValue("@quizScore", quizScore)
                    command.Parameters.AddWithValue("@projectScore", projectScore)
                    command.Parameters.AddWithValue("@midtermScore", midtermScore)
                    command.Parameters.AddWithValue("@finalExamScore", finalExamScore)
                    command.Parameters.AddWithValue("@finalGrade", finalGrade)
                    command.Parameters.AddWithValue("@remarks", NormalizeNullableValue(remarks))
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Private Function MapGradeRecord(reader As MySqlDataReader) As StudentSubjectGradeRecord
            Return New StudentSubjectGradeRecord() With {
                .GradeRecordId = Convert.ToInt32(reader("grade_id")),
                .TeacherRecordId = Convert.ToInt32(reader("teacher_record_id")),
                .TeacherId = Convert.ToString(reader("teacher_id")),
                .TeacherName = Convert.ToString(reader("teacher_name")),
                .StudentRecordId = Convert.ToInt32(reader("student_record_id")),
                .StudentNumber = Convert.ToString(reader("student_number")),
                .StudentName = Convert.ToString(reader("student_name")),
                .StudentSection = Convert.ToString(reader("student_section")),
                .StudentYearLevel = Convert.ToString(reader("student_year_level")),
                .SubjectId = Convert.ToInt32(reader("subject_id")),
                .SubjectCode = Convert.ToString(reader("subject_code")),
                .SubjectName = Convert.ToString(reader("subject_name")),
                .SubjectUnits = NormalizeUnitsValue(reader("subject_units")),
                .SectionName = Convert.ToString(reader("section_name")),
                .QuizScore = ReadNullableDecimal(reader, "quiz_score"),
                .ProjectScore = ReadNullableDecimal(reader, "project_score"),
                .MidtermScore = ReadNullableDecimal(reader, "midterm_score"),
                .FinalExamScore = ReadNullableDecimal(reader, "final_exam_score"),
                .FinalGrade = ReadNullableDecimal(reader, "final_grade"),
                .Remarks = Convert.ToString(reader("remarks")),
                .UpdatedAt = ReadNullableDateTime(reader, "updated_at")
            }
        End Function

        Private Function ShouldIncludeTeacherRosterRecord(reader As MySqlDataReader) As Boolean
            If reader Is Nothing Then
                Return False
            End If

            Dim scheduledSection As String =
                Convert.ToString(reader("scheduled_section_name"))
            If String.IsNullOrWhiteSpace(scheduledSection) Then
                Return True
            End If

            Return StudentScheduleHelper.SectionMatches(
                scheduledSection,
                Convert.ToString(reader("student_section")),
                Convert.ToString(reader("student_year_level")),
                Convert.ToString(reader("subject_year_level")))
        End Function

        Private Sub UpsertTeacherRosterRecord(records As List(Of StudentSubjectGradeRecord),
                                              candidate As StudentSubjectGradeRecord)
            If records Is Nothing OrElse candidate Is Nothing Then
                Return
            End If

            Dim existingIndex As Integer =
                FindTeacherRosterRecordIndex(records,
                                             candidate)
            If existingIndex < 0 Then
                records.Add(candidate)
                Return
            End If

            If ShouldReplaceTeacherRosterRecord(records(existingIndex),
                                                candidate) Then
                records(existingIndex) = candidate
            End If
        End Sub

        Private Function FindTeacherRosterRecordIndex(records As List(Of StudentSubjectGradeRecord),
                                                      candidate As StudentSubjectGradeRecord) As Integer
            If records Is Nothing OrElse candidate Is Nothing Then
                Return -1
            End If

            For index As Integer = 0 To records.Count - 1
                Dim currentRecord As StudentSubjectGradeRecord = records(index)
                If currentRecord Is Nothing Then
                    Continue For
                End If

                If currentRecord.TeacherRecordId = candidate.TeacherRecordId AndAlso
                   currentRecord.StudentRecordId = candidate.StudentRecordId AndAlso
                   currentRecord.SubjectId = candidate.SubjectId Then
                    Return index
                End If
            Next

            Return -1
        End Function

        Private Function ShouldReplaceTeacherRosterRecord(existingRecord As StudentSubjectGradeRecord,
                                                          candidateRecord As StudentSubjectGradeRecord) As Boolean
            If existingRecord Is Nothing Then
                Return True
            End If

            If candidateRecord Is Nothing Then
                Return False
            End If

            If existingRecord.GradeRecordId <= 0 AndAlso
               candidateRecord.GradeRecordId > 0 Then
                Return True
            End If

            If String.IsNullOrWhiteSpace(existingRecord.SectionName) AndAlso
               Not String.IsNullOrWhiteSpace(candidateRecord.SectionName) Then
                Return True
            End If

            If String.IsNullOrWhiteSpace(existingRecord.StudentSection) AndAlso
               Not String.IsNullOrWhiteSpace(candidateRecord.StudentSection) Then
                Return True
            End If

            Return False
        End Function

        Private Function ReadNullableDecimal(reader As MySqlDataReader,
                                             columnName As String) As Decimal?
            Dim ordinal As Integer = reader.GetOrdinal(columnName)
            If reader.IsDBNull(ordinal) Then
                Return Nothing
            End If

            Return Convert.ToDecimal(reader(columnName), CultureInfo.InvariantCulture)
        End Function

        Private Function ReadNullableDateTime(reader As MySqlDataReader,
                                              columnName As String) As DateTime?
            Dim ordinal As Integer = reader.GetOrdinal(columnName)
            If reader.IsDBNull(ordinal) Then
                Return Nothing
            End If

            Return Convert.ToDateTime(reader(columnName), CultureInfo.InvariantCulture)
        End Function

        Private Function NormalizeUnitsValue(value As Object) As String
            If value Is Nothing OrElse value Is DBNull.Value Then
                Return String.Empty
            End If

            Dim unitsValue As Decimal
            If Decimal.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture),
                                NumberStyles.Number,
                                CultureInfo.InvariantCulture,
                                unitsValue) Then
                Return unitsValue.ToString("0.#", CultureInfo.InvariantCulture)
            End If

            Return Convert.ToString(value, CultureInfo.InvariantCulture).Trim()
        End Function

        Private Function NormalizeNullableValue(value As String) As Object
            Dim normalizedValue As String = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Return DBNull.Value
            End If

            Return normalizedValue
        End Function
    End Class
End Namespace
