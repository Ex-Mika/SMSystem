Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models

Namespace Backend.Repositories
    Public Class TeacherScheduleRepository
        Private Const SelectColumnsSql As String =
            "SELECT " &
            "s.schedule_id, " &
            "t.teacher_id AS teacher_record_id, " &
            "t.employee_number, " &
            "TRIM(CONCAT_WS(' ', t.first_name, t.last_name, NULLIF(t.middle_name, ''))) AS teacher_name, " &
            "s.day_name, " &
            "s.session_label, " &
            "COALESCE(s.section_name, '') AS section_name, " &
            "COALESCE(s.subject_code, '') AS subject_code, " &
            "COALESCE(s.subject_name, '') AS subject_name, " &
            "COALESCE(s.room_label, '') AS room_label " &
            "FROM teacher_schedules s " &
            "INNER JOIN teachers t ON t.teacher_id = s.teacher_id "

        Public Function GetAll() As List(Of TeacherScheduleRecord)
            Dim schedules As New List(Of TeacherScheduleRecord)()

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql &
                    "ORDER BY t.employee_number, " &
                    "FIELD(s.day_name, 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'), " &
                    "s.session_label, " &
                    "s.schedule_id;",
                    connection)
                    Using reader As MySqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            schedules.Add(MapSchedule(reader))
                        End While
                    End Using
                End Using
            End Using

            Return schedules
        End Function

        Public Function Save(request As TeacherScheduleSaveRequest) As TeacherScheduleRecord
            If request Is Nothing Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using transaction As MySqlTransaction = connection.BeginTransaction()
                    Dim teacherRecordId As Integer? =
                        ResolveTeacherRecordId(connection, transaction, request.TeacherId)
                    If Not teacherRecordId.HasValue Then
                        Return Nothing
                    End If

                    Using deleteCommand As New MySqlCommand(
                        "DELETE FROM teacher_schedules " &
                        "WHERE teacher_id = @teacherRecordId " &
                        "AND day_name = @dayName " &
                        "AND session_label = @sessionLabel;",
                        connection,
                        transaction)
                        deleteCommand.Parameters.AddWithValue("@teacherRecordId", teacherRecordId.Value)
                        deleteCommand.Parameters.AddWithValue("@dayName", request.Day)
                        deleteCommand.Parameters.AddWithValue("@sessionLabel", request.Session)
                        deleteCommand.ExecuteNonQuery()
                    End Using

                    Using insertCommand As New MySqlCommand(
                        "INSERT INTO teacher_schedules (" &
                        "teacher_id, day_name, session_label, section_name, subject_code, subject_name, room_label" &
                        ") VALUES (" &
                        "@teacherRecordId, @dayName, @sessionLabel, @sectionName, @subjectCode, @subjectName, @roomLabel" &
                        ");",
                        connection,
                        transaction)
                        insertCommand.Parameters.AddWithValue("@teacherRecordId", teacherRecordId.Value)
                        insertCommand.Parameters.AddWithValue("@dayName", request.Day)
                        insertCommand.Parameters.AddWithValue("@sessionLabel", request.Session)
                        insertCommand.Parameters.AddWithValue("@sectionName", NormalizeNullableValue(request.Section))
                        insertCommand.Parameters.AddWithValue("@subjectCode", NormalizeNullableValue(request.SubjectCode))
                        insertCommand.Parameters.AddWithValue("@subjectName", NormalizeNullableValue(request.SubjectName))
                        insertCommand.Parameters.AddWithValue("@roomLabel", NormalizeNullableValue(request.Room))
                        insertCommand.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetByTeacherSlot(request.TeacherId, request.Day, request.Session)
        End Function

        Public Function DeleteByTeacherSlot(teacherId As String,
                                            dayValue As String,
                                            sessionValue As String) As Boolean
            If String.IsNullOrWhiteSpace(teacherId) OrElse
               String.IsNullOrWhiteSpace(dayValue) OrElse
               String.IsNullOrWhiteSpace(sessionValue) Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Dim teacherRecordId As Integer? =
                    ResolveTeacherRecordId(connection, Nothing, teacherId)
                If Not teacherRecordId.HasValue Then
                    Return False
                End If

                Using command As New MySqlCommand(
                    "DELETE FROM teacher_schedules " &
                    "WHERE teacher_id = @teacherRecordId " &
                    "AND day_name = @dayName " &
                    "AND session_label = @sessionLabel;",
                    connection)
                    command.Parameters.AddWithValue("@teacherRecordId", teacherRecordId.Value)
                    command.Parameters.AddWithValue("@dayName", dayValue.Trim())
                    command.Parameters.AddWithValue("@sessionLabel", sessionValue.Trim())
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Public Function GetByTeacherSlot(teacherId As String,
                                         dayValue As String,
                                         sessionValue As String) As TeacherScheduleRecord
            If String.IsNullOrWhiteSpace(teacherId) OrElse
               String.IsNullOrWhiteSpace(dayValue) OrElse
               String.IsNullOrWhiteSpace(sessionValue) Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql &
                    "WHERE t.employee_number = @teacherId " &
                    "AND s.day_name = @dayName " &
                    "AND s.session_label = @sessionLabel " &
                    "LIMIT 1;",
                    connection)
                    command.Parameters.AddWithValue("@teacherId", teacherId.Trim())
                    command.Parameters.AddWithValue("@dayName", dayValue.Trim())
                    command.Parameters.AddWithValue("@sessionLabel", sessionValue.Trim())

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        If Not reader.Read() Then
                            Return Nothing
                        End If

                        Return MapSchedule(reader)
                    End Using
                End Using
            End Using
        End Function

        Private Function ResolveTeacherRecordId(connection As MySqlConnection,
                                                transaction As MySqlTransaction,
                                                teacherId As String) As Integer?
            If connection Is Nothing OrElse String.IsNullOrWhiteSpace(teacherId) Then
                Return Nothing
            End If

            Using command As New MySqlCommand(
                "SELECT teacher_id " &
                "FROM teachers " &
                "WHERE employee_number = @teacherId " &
                "LIMIT 1;",
                connection,
                transaction)
                command.Parameters.AddWithValue("@teacherId", teacherId.Trim())

                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return Nothing
                End If

                Return Convert.ToInt32(result)
            End Using
        End Function

        Private Function MapSchedule(reader As MySqlDataReader) As TeacherScheduleRecord
            Return New TeacherScheduleRecord() With {
                .ScheduleId = Convert.ToInt32(reader("schedule_id")),
                .TeacherRecordId = Convert.ToInt32(reader("teacher_record_id")),
                .TeacherId = Convert.ToString(reader("employee_number")),
                .TeacherName = Convert.ToString(reader("teacher_name")),
                .Day = Convert.ToString(reader("day_name")),
                .Session = Convert.ToString(reader("session_label")),
                .Section = Convert.ToString(reader("section_name")),
                .SubjectCode = Convert.ToString(reader("subject_code")),
                .SubjectName = Convert.ToString(reader("subject_name")),
                .Room = Convert.ToString(reader("room_label"))
            }
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
