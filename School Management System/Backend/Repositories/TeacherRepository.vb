Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models

Namespace Backend.Repositories
    Public Class TeacherRepository
        Private Const SelectColumnsSql As String =
            "SELECT " &
            "t.teacher_id, " &
            "t.user_id, " &
            "t.employee_number, " &
            "t.first_name, " &
            "COALESCE(t.middle_name, '') AS middle_name, " &
            "t.last_name, " &
            "t.department_id, " &
            "COALESCE(t.department_label, '') AS department_label, " &
            "COALESCE(d.department_code, '') AS department_code, " &
            "COALESCE(d.department_name, '') AS department_name, " &
            "COALESCE(t.position_title, '') AS position_title, " &
            "COALESCE(t.advisory_section, '') AS advisory_section, " &
            "COALESCE(t.photo_path, '') AS photo_path, " &
            "u.email " &
            "FROM teachers t " &
            "INNER JOIN users u ON u.user_id = t.user_id " &
            "LEFT JOIN departments d ON d.department_id = t.department_id " &
            "WHERE u.role_key = 'teacher' "

        Public Function GetAll() As List(Of TeacherRecord)
            Dim teachers As New List(Of TeacherRecord)()

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "ORDER BY t.employee_number;",
                    connection)
                    Using reader As MySqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            teachers.Add(MapTeacher(reader))
                        End While
                    End Using
                End Using
            End Using

            Return teachers
        End Function

        Public Function GetByEmployeeNumber(employeeNumber As String) As TeacherRecord
            If String.IsNullOrWhiteSpace(employeeNumber) Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "AND t.employee_number = @employeeNumber LIMIT 1;",
                    connection)
                    command.Parameters.AddWithValue("@employeeNumber", employeeNumber.Trim())

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        If Not reader.Read() Then
                            Return Nothing
                        End If

                        Return MapTeacher(reader)
                    End Using
                End Using
            End Using
        End Function

        Public Function Create(request As TeacherSaveRequest,
                               passwordHash As String,
                               emailAddress As String) As TeacherRecord
            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using transaction As MySqlTransaction = connection.BeginTransaction()
                    Dim userId As Integer
                    Dim departmentInfo As Tuple(Of Object, Object) =
                        ResolveDepartmentInfo(connection, transaction, request.DepartmentText)

                    Using userCommand As New MySqlCommand(
                        "INSERT INTO users (" &
                        "role_key, username, email, password_hash, is_active" &
                        ") VALUES (" &
                        "'teacher', @username, @email, @passwordHash, 1" &
                        ");",
                        connection,
                        transaction)
                        userCommand.Parameters.AddWithValue("@username", BuildDisplayName(request))
                        userCommand.Parameters.AddWithValue("@email", emailAddress)
                        userCommand.Parameters.AddWithValue("@passwordHash", passwordHash)
                        userCommand.ExecuteNonQuery()
                        userId = CInt(userCommand.LastInsertedId)
                    End Using

                    Using teacherCommand As New MySqlCommand(
                        "INSERT INTO teachers (" &
                        "user_id, employee_number, first_name, middle_name, last_name, " &
                        "department_id, department_label, position_title, advisory_section, photo_path" &
                        ") VALUES (" &
                        "@userId, @employeeNumber, @firstName, @middleName, @lastName, " &
                        "@departmentId, @departmentLabel, @positionTitle, @advisorySection, @photoPath" &
                        ");",
                        connection,
                        transaction)
                        teacherCommand.Parameters.AddWithValue("@userId", userId)
                        teacherCommand.Parameters.AddWithValue("@employeeNumber", request.EmployeeNumber.Trim())
                        teacherCommand.Parameters.AddWithValue("@firstName", request.FirstName.Trim())
                        teacherCommand.Parameters.AddWithValue("@middleName", NormalizeNullableValue(request.MiddleName))
                        teacherCommand.Parameters.AddWithValue("@lastName", request.LastName.Trim())
                        teacherCommand.Parameters.AddWithValue("@departmentId", departmentInfo.Item1)
                        teacherCommand.Parameters.AddWithValue("@departmentLabel", departmentInfo.Item2)
                        teacherCommand.Parameters.AddWithValue("@positionTitle", NormalizeNullableValue(request.PositionTitle))
                        teacherCommand.Parameters.AddWithValue("@advisorySection", NormalizeNullableValue(request.AdvisorySection))
                        teacherCommand.Parameters.AddWithValue("@photoPath", NormalizeNullableValue(request.PhotoPath))
                        teacherCommand.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetByEmployeeNumber(request.EmployeeNumber)
        End Function

        Public Function Update(existingRecord As TeacherRecord,
                               request As TeacherSaveRequest,
                               emailAddress As String,
                               shouldUpdatePassword As Boolean,
                               passwordHash As String) As TeacherRecord
            If existingRecord Is Nothing Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using transaction As MySqlTransaction = connection.BeginTransaction()
                    Dim departmentInfo As Tuple(Of Object, Object) =
                        ResolveDepartmentInfo(connection, transaction, request.DepartmentText)
                    Dim updateUsersSql As String =
                        "UPDATE users " &
                        "SET username = @username, " &
                        "email = @email"

                    If shouldUpdatePassword Then
                        updateUsersSql &= ", password_hash = @passwordHash"
                    End If

                    updateUsersSql &= " WHERE user_id = @userId AND role_key = 'teacher';"

                    Using userCommand As New MySqlCommand(updateUsersSql, connection, transaction)
                        userCommand.Parameters.AddWithValue("@username", BuildDisplayName(request))
                        userCommand.Parameters.AddWithValue("@email", emailAddress)
                        userCommand.Parameters.AddWithValue("@userId", existingRecord.UserId)

                        If shouldUpdatePassword Then
                            userCommand.Parameters.AddWithValue("@passwordHash", passwordHash)
                        End If

                        userCommand.ExecuteNonQuery()
                    End Using

                    Using teacherCommand As New MySqlCommand(
                        "UPDATE teachers " &
                        "SET employee_number = @employeeNumber, " &
                        "first_name = @firstName, " &
                        "middle_name = @middleName, " &
                        "last_name = @lastName, " &
                        "department_id = @departmentId, " &
                        "department_label = @departmentLabel, " &
                        "position_title = @positionTitle, " &
                        "advisory_section = @advisorySection, " &
                        "photo_path = @photoPath " &
                        "WHERE teacher_id = @teacherRecordId;",
                        connection,
                        transaction)
                        teacherCommand.Parameters.AddWithValue("@employeeNumber", request.EmployeeNumber.Trim())
                        teacherCommand.Parameters.AddWithValue("@firstName", request.FirstName.Trim())
                        teacherCommand.Parameters.AddWithValue("@middleName", NormalizeNullableValue(request.MiddleName))
                        teacherCommand.Parameters.AddWithValue("@lastName", request.LastName.Trim())
                        teacherCommand.Parameters.AddWithValue("@departmentId", departmentInfo.Item1)
                        teacherCommand.Parameters.AddWithValue("@departmentLabel", departmentInfo.Item2)
                        teacherCommand.Parameters.AddWithValue("@positionTitle", NormalizeNullableValue(request.PositionTitle))
                        teacherCommand.Parameters.AddWithValue("@advisorySection", NormalizeNullableValue(request.AdvisorySection))
                        teacherCommand.Parameters.AddWithValue("@photoPath", NormalizeNullableValue(request.PhotoPath))
                        teacherCommand.Parameters.AddWithValue("@teacherRecordId", existingRecord.TeacherRecordId)
                        teacherCommand.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetByEmployeeNumber(request.EmployeeNumber)
        End Function

        Public Function DeleteByEmployeeNumber(employeeNumber As String) As Boolean
            If String.IsNullOrWhiteSpace(employeeNumber) Then
                Return False
            End If

            Dim existingRecord As TeacherRecord = GetByEmployeeNumber(employeeNumber)
            If existingRecord Is Nothing Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "DELETE FROM users WHERE user_id = @userId AND role_key = 'teacher';",
                    connection)
                    command.Parameters.AddWithValue("@userId", existingRecord.UserId)
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Private Function MapTeacher(reader As MySqlDataReader) As TeacherRecord
            Dim departmentIdOrdinal As Integer = reader.GetOrdinal("department_id")

            Return New TeacherRecord() With {
                .TeacherRecordId = Convert.ToInt32(reader("teacher_id")),
                .UserId = Convert.ToInt32(reader("user_id")),
                .EmployeeNumber = Convert.ToString(reader("employee_number")),
                .FirstName = Convert.ToString(reader("first_name")),
                .MiddleName = Convert.ToString(reader("middle_name")),
                .LastName = Convert.ToString(reader("last_name")),
                .DepartmentId = If(reader.IsDBNull(departmentIdOrdinal),
                                   CType(Nothing, Integer?),
                                   Convert.ToInt32(reader("department_id"))),
                .DepartmentLabel = Convert.ToString(reader("department_label")),
                .DepartmentCode = Convert.ToString(reader("department_code")),
                .DepartmentName = Convert.ToString(reader("department_name")),
                .PositionTitle = Convert.ToString(reader("position_title")),
                .AdvisorySection = Convert.ToString(reader("advisory_section")),
                .PhotoPath = Convert.ToString(reader("photo_path")),
                .Email = Convert.ToString(reader("email"))
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

        Private Function BuildDisplayName(request As TeacherSaveRequest) As String
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
