Imports System.Globalization
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models

Namespace Backend.Repositories
    Public Class SubjectRepository
        Private Const SelectColumnsSql As String =
            "SELECT " &
            "s.subject_id, " &
            "s.subject_code, " &
            "s.subject_title, " &
            "CAST(s.units AS CHAR) AS units, " &
            "s.department_id, " &
            "COALESCE(d.department_code, '') AS department_code, " &
            "COALESCE(d.department_name, '') AS department_name, " &
            "s.course_id, " &
            "COALESCE(s.course_label, '') AS course_label, " &
            "COALESCE(c.course_code, '') AS course_code, " &
            "COALESCE(c.course_name, '') AS course_name, " &
            "COALESCE(s.year_level, '') AS year_level, " &
            "COALESCE(s.description, '') AS description " &
            "FROM subjects s " &
            "LEFT JOIN departments d ON d.department_id = s.department_id " &
            "LEFT JOIN courses c ON c.course_id = s.course_id "

        Public Function GetAll() As List(Of SubjectRecord)
            Dim subjects As New List(Of SubjectRecord)()

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "ORDER BY s.subject_code, s.subject_title;",
                    connection)
                    Using reader As MySqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            subjects.Add(MapSubject(reader))
                        End While
                    End Using
                End Using
            End Using

            Return subjects
        End Function

        Public Function GetBySubjectCode(subjectCode As String) As SubjectRecord
            If String.IsNullOrWhiteSpace(subjectCode) Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "WHERE s.subject_code = @subjectCode LIMIT 1;",
                    connection)
                    command.Parameters.AddWithValue("@subjectCode", subjectCode.Trim())

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        If Not reader.Read() Then
                            Return Nothing
                        End If

                        Return MapSubject(reader)
                    End Using
                End Using
            End Using
        End Function

        Public Function Create(request As SubjectSaveRequest) As SubjectRecord
            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using transaction As MySqlTransaction = connection.BeginTransaction()
                    Dim courseInfo As Tuple(Of Object, Object, Object) =
                        ResolveCourseInfo(connection, transaction, request.CourseText)

                    Using command As New MySqlCommand(
                        "INSERT INTO subjects (" &
                        "subject_code, subject_title, units, department_id, course_id, course_label, year_level, description" &
                        ") VALUES (" &
                        "@subjectCode, @subjectTitle, @units, @departmentId, @courseId, @courseLabel, @yearLevel, @description" &
                        ");",
                        connection,
                        transaction)
                        command.Parameters.AddWithValue("@subjectCode", request.SubjectCode.Trim())
                        command.Parameters.AddWithValue("@subjectTitle", request.SubjectName.Trim())
                        command.Parameters.AddWithValue("@units", NormalizeUnitsValue(request.Units))
                        command.Parameters.AddWithValue("@departmentId", courseInfo.Item3)
                        command.Parameters.AddWithValue("@courseId", courseInfo.Item1)
                        command.Parameters.AddWithValue("@courseLabel", courseInfo.Item2)
                        command.Parameters.AddWithValue("@yearLevel", NormalizeNullableValue(request.YearLevel))
                        command.Parameters.AddWithValue("@description", DBNull.Value)
                        command.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetBySubjectCode(request.SubjectCode)
        End Function

        Public Function Update(existingRecord As SubjectRecord,
                               request As SubjectSaveRequest) As SubjectRecord
            If existingRecord Is Nothing Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using transaction As MySqlTransaction = connection.BeginTransaction()
                    Dim courseInfo As Tuple(Of Object, Object, Object) =
                        ResolveCourseInfo(connection, transaction, request.CourseText)

                    Using command As New MySqlCommand(
                        "UPDATE subjects " &
                        "SET subject_code = @subjectCode, " &
                        "subject_title = @subjectTitle, " &
                        "units = @units, " &
                        "department_id = @departmentId, " &
                        "course_id = @courseId, " &
                        "course_label = @courseLabel, " &
                        "year_level = @yearLevel " &
                        "WHERE subject_id = @subjectId;",
                        connection,
                        transaction)
                        command.Parameters.AddWithValue("@subjectCode", request.SubjectCode.Trim())
                        command.Parameters.AddWithValue("@subjectTitle", request.SubjectName.Trim())
                        command.Parameters.AddWithValue("@units", NormalizeUnitsValue(request.Units))
                        command.Parameters.AddWithValue("@departmentId", courseInfo.Item3)
                        command.Parameters.AddWithValue("@courseId", courseInfo.Item1)
                        command.Parameters.AddWithValue("@courseLabel", courseInfo.Item2)
                        command.Parameters.AddWithValue("@yearLevel", NormalizeNullableValue(request.YearLevel))
                        command.Parameters.AddWithValue("@subjectId", existingRecord.SubjectId)
                        command.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetBySubjectCode(request.SubjectCode)
        End Function

        Public Function DeleteBySubjectCode(subjectCode As String) As Boolean
            If String.IsNullOrWhiteSpace(subjectCode) Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "DELETE FROM subjects WHERE subject_code = @subjectCode;",
                    connection)
                    command.Parameters.AddWithValue("@subjectCode", subjectCode.Trim())
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Private Function MapSubject(reader As MySqlDataReader) As SubjectRecord
            Dim departmentIdOrdinal As Integer = reader.GetOrdinal("department_id")
            Dim courseIdOrdinal As Integer = reader.GetOrdinal("course_id")

            Return New SubjectRecord() With {
                .SubjectId = Convert.ToInt32(reader("subject_id")),
                .SubjectCode = Convert.ToString(reader("subject_code")),
                .SubjectName = Convert.ToString(reader("subject_title")),
                .Units = NormalizeUnitsText(Convert.ToString(reader("units"))),
                .DepartmentId = If(reader.IsDBNull(departmentIdOrdinal),
                                   CType(Nothing, Integer?),
                                   Convert.ToInt32(reader("department_id"))),
                .DepartmentCode = Convert.ToString(reader("department_code")),
                .DepartmentName = Convert.ToString(reader("department_name")),
                .CourseId = If(reader.IsDBNull(courseIdOrdinal),
                               CType(Nothing, Integer?),
                               Convert.ToInt32(reader("course_id"))),
                .CourseLabel = Convert.ToString(reader("course_label")),
                .CourseCode = Convert.ToString(reader("course_code")),
                .CourseName = Convert.ToString(reader("course_name")),
                .YearLevel = Convert.ToString(reader("year_level")),
                .Description = Convert.ToString(reader("description"))
            }
        End Function

        Private Function ResolveCourseInfo(connection As MySqlConnection,
                                           transaction As MySqlTransaction,
                                           courseText As String) As Tuple(Of Object, Object, Object)
            Dim normalizedCourseText As String = If(courseText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedCourseText) Then
                Return Tuple.Create(Of Object, Object, Object)(DBNull.Value, DBNull.Value, DBNull.Value)
            End If

            Using command As New MySqlCommand(
                "SELECT course_id, department_id " &
                "FROM courses " &
                "WHERE course_code = @courseText OR course_name = @courseText " &
                "ORDER BY CASE WHEN course_code = @courseText THEN 0 ELSE 1 END " &
                "LIMIT 1;",
                connection,
                transaction)
                command.Parameters.AddWithValue("@courseText", normalizedCourseText)

                Using reader As MySqlDataReader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Return Tuple.Create(Of Object, Object, Object)(DBNull.Value,
                                                                       normalizedCourseText,
                                                                       DBNull.Value)
                    End If

                    Dim departmentIdOrdinal As Integer = reader.GetOrdinal("department_id")
                    Dim departmentId As Object =
                        If(reader.IsDBNull(departmentIdOrdinal),
                           DBNull.Value,
                           CType(Convert.ToInt32(reader("department_id")), Object))

                    Return Tuple.Create(Of Object, Object, Object)(Convert.ToInt32(reader("course_id")),
                                                                   DBNull.Value,
                                                                   departmentId)
                End Using
            End Using
        End Function

        Private Function NormalizeUnitsValue(value As String) As Decimal
            Dim normalizedValue As String = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Return 3D
            End If

            Dim parsedUnits As Decimal
            If Decimal.TryParse(normalizedValue,
                                NumberStyles.Number,
                                CultureInfo.InvariantCulture,
                                parsedUnits) OrElse
               Decimal.TryParse(normalizedValue,
                                NumberStyles.Number,
                                CultureInfo.CurrentCulture,
                                parsedUnits) Then
                Return Decimal.Round(parsedUnits, 1, MidpointRounding.AwayFromZero)
            End If

            Return 3D
        End Function

        Private Function NormalizeUnitsText(value As String) As String
            Dim normalizedValue As String = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Return String.Empty
            End If

            Dim parsedUnits As Decimal
            If Decimal.TryParse(normalizedValue,
                                NumberStyles.Number,
                                CultureInfo.InvariantCulture,
                                parsedUnits) OrElse
               Decimal.TryParse(normalizedValue,
                                NumberStyles.Number,
                                CultureInfo.CurrentCulture,
                                parsedUnits) Then
                Return parsedUnits.ToString("0.#", CultureInfo.InvariantCulture)
            End If

            Return normalizedValue
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
