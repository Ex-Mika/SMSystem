Imports System.Collections.Generic
Imports MySql.Data.MySqlClient

Namespace Backend.Repositories
    Public Class StudentSubjectEnrollmentRepository
        Public Function GetSubjectIdsByStudentId(studentRecordId As Integer) As List(Of Integer)
            Dim subjectIds As New List(Of Integer)()
            If studentRecordId <= 0 Then
                Return subjectIds
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "SELECT subject_id " &
                    "FROM student_subject_enrollments " &
                    "WHERE student_id = @studentId " &
                    "ORDER BY created_at, enrollment_id;",
                    connection)
                    command.Parameters.AddWithValue("@studentId", studentRecordId)

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            subjectIds.Add(Convert.ToInt32(reader("subject_id")))
                        End While
                    End Using
                End Using
            End Using

            Return subjectIds
        End Function

        Public Function Exists(studentRecordId As Integer, subjectId As Integer) As Boolean
            If studentRecordId <= 0 OrElse subjectId <= 0 Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "SELECT COUNT(*) " &
                    "FROM student_subject_enrollments " &
                    "WHERE student_id = @studentId " &
                    "AND subject_id = @subjectId;",
                    connection)
                    command.Parameters.AddWithValue("@studentId", studentRecordId)
                    command.Parameters.AddWithValue("@subjectId", subjectId)
                    Return Convert.ToInt32(command.ExecuteScalar()) > 0
                End Using
            End Using
        End Function

        Public Function Create(studentRecordId As Integer, subjectId As Integer) As Boolean
            If studentRecordId <= 0 OrElse subjectId <= 0 Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "INSERT INTO student_subject_enrollments (" &
                    "student_id, subject_id" &
                    ") VALUES (" &
                    "@studentId, @subjectId" &
                    ");",
                    connection)
                    command.Parameters.AddWithValue("@studentId", studentRecordId)
                    command.Parameters.AddWithValue("@subjectId", subjectId)
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Public Function Delete(studentRecordId As Integer, subjectId As Integer) As Boolean
            If studentRecordId <= 0 OrElse subjectId <= 0 Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "DELETE FROM student_subject_enrollments " &
                    "WHERE student_id = @studentId " &
                    "AND subject_id = @subjectId;",
                    connection)
                    command.Parameters.AddWithValue("@studentId", studentRecordId)
                    command.Parameters.AddWithValue("@subjectId", subjectId)
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function
    End Class
End Namespace
