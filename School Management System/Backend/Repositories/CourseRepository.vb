Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models

Namespace Backend.Repositories
    Public Class CourseRepository
        Private Const SelectColumnsSql As String =
            "SELECT " &
            "c.course_id, " &
            "c.course_code, " &
            "c.course_name, " &
            "c.department_id, " &
            "COALESCE(c.department_label, '') AS department_label, " &
            "COALESCE(d.department_code, '') AS department_code, " &
            "COALESCE(d.department_name, '') AS department_name, " &
            "COALESCE(c.units, '') AS units " &
            "FROM courses c " &
            "LEFT JOIN departments d ON d.department_id = c.department_id "

        Public Function GetAll() As List(Of CourseRecord)
            Dim courses As New List(Of CourseRecord)()

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "ORDER BY c.course_code, c.course_name;",
                    connection)
                    Using reader As MySqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            courses.Add(MapCourse(reader))
                        End While
                    End Using
                End Using
            End Using

            Return courses
        End Function

        Public Function GetByCourseCode(courseCode As String) As CourseRecord
            If String.IsNullOrWhiteSpace(courseCode) Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "WHERE c.course_code = @courseCode LIMIT 1;",
                    connection)
                    command.Parameters.AddWithValue("@courseCode", courseCode.Trim())

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        If Not reader.Read() Then
                            Return Nothing
                        End If

                        Return MapCourse(reader)
                    End Using
                End Using
            End Using
        End Function

        Public Function Create(request As CourseSaveRequest) As CourseRecord
            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using transaction As MySqlTransaction = connection.BeginTransaction()
                    Dim departmentInfo As Tuple(Of Object, Object) =
                        ResolveDepartmentInfo(connection, transaction, request.DepartmentText)

                    Using command As New MySqlCommand(
                        "INSERT INTO courses (" &
                        "course_code, course_name, department_id, department_label, units" &
                        ") VALUES (" &
                        "@courseCode, @courseName, @departmentId, @departmentLabel, @units" &
                        ");",
                        connection,
                        transaction)
                        command.Parameters.AddWithValue("@courseCode", request.CourseCode.Trim())
                        command.Parameters.AddWithValue("@courseName", request.CourseName.Trim())
                        command.Parameters.AddWithValue("@departmentId", departmentInfo.Item1)
                        command.Parameters.AddWithValue("@departmentLabel", departmentInfo.Item2)
                        command.Parameters.AddWithValue("@units", NormalizeNullableValue(request.Units))
                        command.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetByCourseCode(request.CourseCode)
        End Function

        Public Function Update(existingRecord As CourseRecord,
                               request As CourseSaveRequest) As CourseRecord
            If existingRecord Is Nothing Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using transaction As MySqlTransaction = connection.BeginTransaction()
                    Dim departmentInfo As Tuple(Of Object, Object) =
                        ResolveDepartmentInfo(connection, transaction, request.DepartmentText)

                    Using command As New MySqlCommand(
                        "UPDATE courses " &
                        "SET course_code = @courseCode, " &
                        "course_name = @courseName, " &
                        "department_id = @departmentId, " &
                        "department_label = @departmentLabel, " &
                        "units = @units " &
                        "WHERE course_id = @courseId;",
                        connection,
                        transaction)
                        command.Parameters.AddWithValue("@courseCode", request.CourseCode.Trim())
                        command.Parameters.AddWithValue("@courseName", request.CourseName.Trim())
                        command.Parameters.AddWithValue("@departmentId", departmentInfo.Item1)
                        command.Parameters.AddWithValue("@departmentLabel", departmentInfo.Item2)
                        command.Parameters.AddWithValue("@units", NormalizeNullableValue(request.Units))
                        command.Parameters.AddWithValue("@courseId", existingRecord.CourseId)
                        command.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetByCourseCode(request.CourseCode)
        End Function

        Public Function DeleteByCourseCode(courseCode As String) As Boolean
            If String.IsNullOrWhiteSpace(courseCode) Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "DELETE FROM courses WHERE course_code = @courseCode;",
                    connection)
                    command.Parameters.AddWithValue("@courseCode", courseCode.Trim())
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Private Function MapCourse(reader As MySqlDataReader) As CourseRecord
            Dim departmentIdOrdinal As Integer = reader.GetOrdinal("department_id")

            Return New CourseRecord() With {
                .CourseId = Convert.ToInt32(reader("course_id")),
                .CourseCode = Convert.ToString(reader("course_code")),
                .CourseName = Convert.ToString(reader("course_name")),
                .DepartmentId = If(reader.IsDBNull(departmentIdOrdinal),
                                   CType(Nothing, Integer?),
                                   Convert.ToInt32(reader("department_id"))),
                .DepartmentLabel = Convert.ToString(reader("department_label")),
                .DepartmentCode = Convert.ToString(reader("department_code")),
                .DepartmentName = Convert.ToString(reader("department_name")),
                .Units = Convert.ToString(reader("units"))
            }
        End Function

        Private Function ResolveDepartmentInfo(connection As MySqlConnection,
                                               transaction As MySqlTransaction,
                                               departmentText As String) As Tuple(Of Object, Object)
            Dim normalizedDepartmentText As String = If(departmentText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedDepartmentText) Then
                Return Tuple.Create(Of Object, Object)(DBNull.Value, DBNull.Value)
            End If

            Using command As New MySqlCommand(
                "SELECT department_id " &
                "FROM departments " &
                "WHERE department_code = @departmentText OR department_name = @departmentText " &
                "ORDER BY CASE " &
                " WHEN department_code = @departmentText THEN 0 " &
                " WHEN department_name = @departmentText THEN 1 " &
                " ELSE 2 END " &
                "LIMIT 1;",
                connection,
                transaction)
                command.Parameters.AddWithValue("@departmentText", normalizedDepartmentText)

                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return Tuple.Create(Of Object, Object)(DBNull.Value, normalizedDepartmentText)
                End If

                Return Tuple.Create(Of Object, Object)(Convert.ToInt32(result), DBNull.Value)
            End Using
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
