Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models

Namespace Backend.Repositories
    Public Class UserRepository
        Private Const GetUserByIdentifierSql As String =
            "SELECT " &
            "u.user_id, " &
            "u.role_key, " &
            "u.username, " &
            "u.email, " &
            "u.password_hash, " &
            "u.is_active, " &
            "COALESCE(s.student_number, t.employee_number, a.admin_code, u.email) AS reference_code " &
            "FROM users u " &
            "LEFT JOIN students s ON s.user_id = u.user_id " &
            "LEFT JOIN teachers t ON t.user_id = u.user_id " &
            "LEFT JOIN administrators a ON a.user_id = u.user_id " &
            "WHERE u.role_key = @roleKey AND (" &
            "(u.role_key = 'student' AND LOWER(s.student_number) = @normalizedIdentifier) OR " &
            "(u.role_key = 'teacher' AND LOWER(t.employee_number) = @normalizedIdentifier) OR " &
            "(u.role_key = 'admin' AND LOWER(u.email) = @normalizedIdentifier)" &
            ") LIMIT 1;"

        Public Function GetByLoginIdentifier(role As UserRole,
                                             identifier As String) As UserAccount
            If String.IsNullOrWhiteSpace(identifier) Then
                Return Nothing
            End If

            Dim normalizedIdentifier As String = identifier.Trim().ToLowerInvariant()

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(GetUserByIdentifierSql, connection)
                    command.Parameters.AddWithValue("@roleKey", Database.DatabaseModule.GetRoleKey(role))
                    command.Parameters.AddWithValue("@normalizedIdentifier", normalizedIdentifier)

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        If Not reader.Read() Then
                            Return Nothing
                        End If

                        Return MapUser(reader)
                    End Using
                End Using
            End Using
        End Function

        Private Function MapUser(reader As MySqlDataReader) As UserAccount
            Return New UserAccount() With {
                .UserId = Convert.ToInt32(reader("user_id")),
                .Role = Database.DatabaseModule.ParseRoleKey(Convert.ToString(reader("role_key"))),
                .Username = Convert.ToString(reader("username")),
                .Email = Convert.ToString(reader("email")),
                .PasswordHash = Convert.ToString(reader("password_hash")),
                .IsActive = Convert.ToBoolean(reader("is_active")),
                .ReferenceCode = Convert.ToString(reader("reference_code"))
            }
        End Function
    End Class
End Namespace
