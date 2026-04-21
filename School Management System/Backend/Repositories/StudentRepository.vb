Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models

Namespace Backend.Repositories
    Public Class StudentRepository
        Private Const SelectColumnsSql As String =
            "SELECT " &
            "s.student_id, " &
            "s.user_id, " &
            "s.student_number, " &
            "s.first_name, " &
            "COALESCE(s.middle_name, '') AS middle_name, " &
            "s.last_name, " &
            "s.year_level, " &
            "s.course_id, " &
            "COALESCE(c.course_code, '') AS course_code, " &
            "COALESCE(c.course_name, '') AS course_name, " &
            "COALESCE(s.section_name, '') AS section_name, " &
            "COALESCE(s.photo_path, '') AS photo_path, " &
            "u.email " &
            "FROM students s " &
            "INNER JOIN users u ON u.user_id = s.user_id " &
            "LEFT JOIN courses c ON c.course_id = s.course_id " &
            "WHERE u.role_key = 'student' "

        Public Function GetAll() As List(Of StudentRecord)
            Dim students As New List(Of StudentRecord)()

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "ORDER BY s.student_number;",
                    connection)
                    Using reader As MySqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            students.Add(MapStudent(reader))
                        End While
                    End Using
                End Using
            End Using

            Return students
        End Function

        Public Function GetByStudentNumber(studentNumber As String) As StudentRecord
            If String.IsNullOrWhiteSpace(studentNumber) Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "AND s.student_number = @studentNumber LIMIT 1;",
                    connection)
                    command.Parameters.AddWithValue("@studentNumber", studentNumber.Trim())

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        If Not reader.Read() Then
                            Return Nothing
                        End If

                        Return MapStudent(reader)
                    End Using
                End Using
            End Using
        End Function

        Public Function Create(request As StudentSaveRequest,
                               passwordHash As String,
                               emailAddress As String) As StudentRecord
            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using transaction As MySqlTransaction = connection.BeginTransaction()
                    Dim userId As Integer

                    Using userCommand As New MySqlCommand(
                        "INSERT INTO users (" &
                        "role_key, username, email, password_hash, is_active" &
                        ") VALUES (" &
                        "'student', @username, @email, @passwordHash, 1" &
                        ");",
                        connection,
                        transaction)
                        userCommand.Parameters.AddWithValue("@username", BuildDisplayName(request))
                        userCommand.Parameters.AddWithValue("@email", emailAddress)
                        userCommand.Parameters.AddWithValue("@passwordHash", passwordHash)
                        userCommand.ExecuteNonQuery()
                        userId = CInt(userCommand.LastInsertedId)
                    End Using

                    Using studentCommand As New MySqlCommand(
                        "INSERT INTO students (" &
                        "user_id, student_number, first_name, middle_name, last_name, " &
                        "course_id, year_level, section_name, photo_path" &
                        ") VALUES (" &
                        "@userId, @studentNumber, @firstName, @middleName, @lastName, " &
                        "@courseId, @yearLevel, @sectionName, @photoPath" &
                        ");",
                        connection,
                        transaction)
                        studentCommand.Parameters.AddWithValue("@userId", userId)
                        studentCommand.Parameters.AddWithValue("@studentNumber", request.StudentNumber.Trim())
                        studentCommand.Parameters.AddWithValue("@firstName", request.FirstName.Trim())
                        studentCommand.Parameters.AddWithValue("@middleName", NormalizeNullableValue(request.MiddleName))
                        studentCommand.Parameters.AddWithValue("@lastName", request.LastName.Trim())
                        studentCommand.Parameters.AddWithValue("@courseId", ResolveCourseId(connection,
                                                                                          transaction,
                                                                                          request.CourseText))
                        studentCommand.Parameters.AddWithValue("@yearLevel", NormalizeNullableInteger(request.YearLevel))
                        studentCommand.Parameters.AddWithValue("@sectionName", NormalizeNullableValue(request.SectionName))
                        studentCommand.Parameters.AddWithValue("@photoPath", NormalizeNullableValue(request.PhotoPath))
                        studentCommand.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetByStudentNumber(request.StudentNumber)
        End Function

        Public Function Update(existingRecord As StudentRecord,
                               request As StudentSaveRequest,
                               emailAddress As String,
                               shouldUpdatePassword As Boolean,
                               passwordHash As String) As StudentRecord
            If existingRecord Is Nothing Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using transaction As MySqlTransaction = connection.BeginTransaction()
                    Dim updateUsersSql As String =
                        "UPDATE users " &
                        "SET username = @username, " &
                        "email = @email"

                    If shouldUpdatePassword Then
                        updateUsersSql &= ", password_hash = @passwordHash"
                    End If

                    updateUsersSql &= " WHERE user_id = @userId AND role_key = 'student';"

                    Using userCommand As New MySqlCommand(updateUsersSql, connection, transaction)
                        userCommand.Parameters.AddWithValue("@username", BuildDisplayName(request))
                        userCommand.Parameters.AddWithValue("@email", emailAddress)
                        userCommand.Parameters.AddWithValue("@userId", existingRecord.UserId)

                        If shouldUpdatePassword Then
                            userCommand.Parameters.AddWithValue("@passwordHash", passwordHash)
                        End If

                        userCommand.ExecuteNonQuery()
                    End Using

                    Using studentCommand As New MySqlCommand(
                        "UPDATE students " &
                        "SET student_number = @studentNumber, " &
                        "first_name = @firstName, " &
                        "middle_name = @middleName, " &
                        "last_name = @lastName, " &
                        "course_id = @courseId, " &
                        "year_level = @yearLevel, " &
                        "section_name = @sectionName, " &
                        "photo_path = @photoPath " &
                        "WHERE student_id = @studentRecordId;",
                        connection,
                        transaction)
                        studentCommand.Parameters.AddWithValue("@studentNumber", request.StudentNumber.Trim())
                        studentCommand.Parameters.AddWithValue("@firstName", request.FirstName.Trim())
                        studentCommand.Parameters.AddWithValue("@middleName", NormalizeNullableValue(request.MiddleName))
                        studentCommand.Parameters.AddWithValue("@lastName", request.LastName.Trim())
                        studentCommand.Parameters.AddWithValue("@courseId", ResolveCourseId(connection,
                                                                                          transaction,
                                                                                          request.CourseText))
                        studentCommand.Parameters.AddWithValue("@yearLevel", NormalizeNullableInteger(request.YearLevel))
                        studentCommand.Parameters.AddWithValue("@sectionName", NormalizeNullableValue(request.SectionName))
                        studentCommand.Parameters.AddWithValue("@photoPath", NormalizeNullableValue(request.PhotoPath))
                        studentCommand.Parameters.AddWithValue("@studentRecordId", existingRecord.StudentRecordId)
                        studentCommand.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetByStudentNumber(request.StudentNumber)
        End Function

        Public Function DeleteByStudentNumber(studentNumber As String) As Boolean
            If String.IsNullOrWhiteSpace(studentNumber) Then
                Return False
            End If

            Dim existingRecord As StudentRecord = GetByStudentNumber(studentNumber)
            If existingRecord Is Nothing Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "DELETE FROM users WHERE user_id = @userId AND role_key = 'student';",
                    connection)
                    command.Parameters.AddWithValue("@userId", existingRecord.UserId)
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Public Function UpdateSection(studentRecordId As Integer,
                                      sectionName As String) As Boolean
            If studentRecordId <= 0 Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "UPDATE students " &
                    "SET section_name = @sectionName " &
                    "WHERE student_id = @studentRecordId;",
                    connection)
                    command.Parameters.AddWithValue("@sectionName",
                                                    NormalizeNullableValue(sectionName))
                    command.Parameters.AddWithValue("@studentRecordId", studentRecordId)
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Private Function MapStudent(reader As MySqlDataReader) As StudentRecord
            Dim yearLevelOrdinal As Integer = reader.GetOrdinal("year_level")
            Dim courseIdOrdinal As Integer = reader.GetOrdinal("course_id")

            Return New StudentRecord() With {
                .StudentRecordId = Convert.ToInt32(reader("student_id")),
                .UserId = Convert.ToInt32(reader("user_id")),
                .StudentNumber = Convert.ToString(reader("student_number")),
                .FirstName = Convert.ToString(reader("first_name")),
                .MiddleName = Convert.ToString(reader("middle_name")),
                .LastName = Convert.ToString(reader("last_name")),
                .YearLevel = If(reader.IsDBNull(yearLevelOrdinal),
                                CType(Nothing, Integer?),
                                Convert.ToInt32(reader("year_level"))),
                .CourseId = If(reader.IsDBNull(courseIdOrdinal),
                               CType(Nothing, Integer?),
                               Convert.ToInt32(reader("course_id"))),
                .CourseCode = Convert.ToString(reader("course_code")),
                .CourseName = Convert.ToString(reader("course_name")),
                .SectionName = Convert.ToString(reader("section_name")),
                .PhotoPath = Convert.ToString(reader("photo_path")),
                .Email = Convert.ToString(reader("email"))
            }
        End Function

        Private Function ResolveCourseId(connection As MySqlConnection,
                                         transaction As MySqlTransaction,
                                         courseText As String) As Object
            Dim normalizedCourseText As String = If(courseText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedCourseText) Then
                Return DBNull.Value
            End If

            Using command As New MySqlCommand(
                "SELECT course_id " &
                "FROM courses " &
                "WHERE course_code = @courseText OR course_name = @courseText " &
                "ORDER BY CASE WHEN course_code = @courseText THEN 0 ELSE 1 END " &
                "LIMIT 1;",
                connection,
                transaction)
                command.Parameters.AddWithValue("@courseText", normalizedCourseText)

                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Throw New InvalidOperationException(
                        "Course not found. Use an existing course code or course name.")
                End If

                Return Convert.ToInt32(result)
            End Using
        End Function

        Private Function NormalizeNullableValue(value As String) As Object
            Dim normalizedValue As String = If(value, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Return DBNull.Value
            End If

            Return normalizedValue
        End Function

        Private Function NormalizeNullableInteger(value As Integer?) As Object
            If value.HasValue Then
                Return value.Value
            End If

            Return DBNull.Value
        End Function

        Private Function BuildDisplayName(request As StudentSaveRequest) As String
            Dim parts As New List(Of String)()

            If Not String.IsNullOrWhiteSpace(request.FirstName) Then
                parts.Add(request.FirstName.Trim())
            End If

            If Not String.IsNullOrWhiteSpace(request.MiddleName) Then
                parts.Add(request.MiddleName.Trim())
            End If

            If Not String.IsNullOrWhiteSpace(request.LastName) Then
                parts.Add(request.LastName.Trim())
            End If

            Return String.Join(" ", parts)
        End Function
    End Class
End Namespace
