Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models

Namespace Backend.Repositories
    Public Class AdministratorRepository
        Private Const SelectColumnsSql As String =
            "SELECT " &
            "a.administrator_id, " &
            "a.user_id, " &
            "a.admin_code, " &
            "a.first_name, " &
            "COALESCE(a.middle_name, '') AS middle_name, " &
            "a.last_name, " &
            "COALESCE(a.role_title, '') AS role_title, " &
            "u.email, " &
            "COALESCE(a.photo_path, '') AS photo_path " &
            "FROM administrators a " &
            "INNER JOIN users u ON u.user_id = a.user_id " &
            "WHERE u.role_key = 'admin' "

        Public Function GetAll() As List(Of AdministratorRecord)
            Dim administrators As New List(Of AdministratorRecord)()

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(SelectColumnsSql & "ORDER BY a.admin_code;", connection)
                    Using reader As MySqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            administrators.Add(MapAdministrator(reader))
                        End While
                    End Using
                End Using
            End Using

            Return administrators
        End Function

        Public Function GetByAdministratorCode(administratorCode As String) As AdministratorRecord
            If String.IsNullOrWhiteSpace(administratorCode) Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "AND a.admin_code = @adminCode LIMIT 1;",
                    connection)
                    command.Parameters.AddWithValue("@adminCode", administratorCode.Trim())

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        If Not reader.Read() Then
                            Return Nothing
                        End If

                        Return MapAdministrator(reader)
                    End Using
                End Using
            End Using
        End Function

        Public Function GetByEmail(email As String) As AdministratorRecord
            If String.IsNullOrWhiteSpace(email) Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "AND u.email = @email LIMIT 1;",
                    connection)
                    command.Parameters.AddWithValue("@email", email.Trim())

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        If Not reader.Read() Then
                            Return Nothing
                        End If

                        Return MapAdministrator(reader)
                    End Using
                End Using
            End Using
        End Function

        Public Function Create(request As AdministratorSaveRequest,
                               passwordHash As String) As AdministratorRecord
            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using transaction As MySqlTransaction = connection.BeginTransaction()
                    Dim userId As Integer

                    Using userCommand As New MySqlCommand(
                        "INSERT INTO users (" &
                        "role_key, username, email, password_hash, is_active" &
                        ") VALUES (" &
                        "'admin', @username, @email, @passwordHash, 1" &
                        ");",
                        connection,
                        transaction)
                        userCommand.Parameters.AddWithValue("@username", BuildDisplayName(request))
                        userCommand.Parameters.AddWithValue("@email", request.Email.Trim())
                        userCommand.Parameters.AddWithValue("@passwordHash", passwordHash)
                        userCommand.ExecuteNonQuery()
                        userId = CInt(userCommand.LastInsertedId)
                    End Using

                    Using administratorCommand As New MySqlCommand(
                        "INSERT INTO administrators (" &
                        "user_id, admin_code, first_name, middle_name, last_name, role_title, photo_path" &
                        ") VALUES (" &
                        "@userId, @adminCode, @firstName, @middleName, @lastName, @roleTitle, @photoPath" &
                        ");",
                        connection,
                        transaction)
                        administratorCommand.Parameters.AddWithValue("@userId", userId)
                        administratorCommand.Parameters.AddWithValue("@adminCode", request.AdministratorCode.Trim())
                        administratorCommand.Parameters.AddWithValue("@firstName", request.FirstName.Trim())
                        administratorCommand.Parameters.AddWithValue("@middleName", NormalizeNullableValue(request.MiddleName))
                        administratorCommand.Parameters.AddWithValue("@lastName", request.LastName.Trim())
                        administratorCommand.Parameters.AddWithValue("@roleTitle", NormalizeNullableValue(request.RoleTitle))
                        administratorCommand.Parameters.AddWithValue("@photoPath", NormalizeNullableValue(request.PhotoPath))
                        administratorCommand.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetByAdministratorCode(request.AdministratorCode)
        End Function

        Public Function Update(existingRecord As AdministratorRecord,
                               request As AdministratorSaveRequest,
                               shouldUpdatePassword As Boolean,
                               passwordHash As String) As AdministratorRecord
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

                    updateUsersSql &= " WHERE user_id = @userId AND role_key = 'admin';"

                    Using userCommand As New MySqlCommand(updateUsersSql, connection, transaction)
                        userCommand.Parameters.AddWithValue("@username", BuildDisplayName(request))
                        userCommand.Parameters.AddWithValue("@email", request.Email.Trim())
                        userCommand.Parameters.AddWithValue("@userId", existingRecord.UserId)

                        If shouldUpdatePassword Then
                            userCommand.Parameters.AddWithValue("@passwordHash", passwordHash)
                        End If

                        userCommand.ExecuteNonQuery()
                    End Using

                    Using administratorCommand As New MySqlCommand(
                        "UPDATE administrators " &
                        "SET admin_code = @adminCode, " &
                        "first_name = @firstName, " &
                        "middle_name = @middleName, " &
                        "last_name = @lastName, " &
                        "role_title = @roleTitle, " &
                        "photo_path = @photoPath " &
                        "WHERE administrator_id = @administratorRecordId;",
                        connection,
                        transaction)
                        administratorCommand.Parameters.AddWithValue("@adminCode", request.AdministratorCode.Trim())
                        administratorCommand.Parameters.AddWithValue("@firstName", request.FirstName.Trim())
                        administratorCommand.Parameters.AddWithValue("@middleName", NormalizeNullableValue(request.MiddleName))
                        administratorCommand.Parameters.AddWithValue("@lastName", request.LastName.Trim())
                        administratorCommand.Parameters.AddWithValue("@roleTitle", NormalizeNullableValue(request.RoleTitle))
                        administratorCommand.Parameters.AddWithValue("@photoPath", NormalizeNullableValue(request.PhotoPath))
                        administratorCommand.Parameters.AddWithValue("@administratorRecordId", existingRecord.AdministratorRecordId)
                        administratorCommand.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                End Using
            End Using

            Return GetByAdministratorCode(request.AdministratorCode)
        End Function

        Public Function DeleteByAdministratorCode(administratorCode As String) As Boolean
            If String.IsNullOrWhiteSpace(administratorCode) Then
                Return False
            End If

            Dim existingRecord As AdministratorRecord = GetByAdministratorCode(administratorCode)
            If existingRecord Is Nothing Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "DELETE FROM users WHERE user_id = @userId AND role_key = 'admin';",
                    connection)
                    command.Parameters.AddWithValue("@userId", existingRecord.UserId)
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Private Function MapAdministrator(reader As MySqlDataReader) As AdministratorRecord
            Return New AdministratorRecord() With {
                .AdministratorRecordId = Convert.ToInt32(reader("administrator_id")),
                .UserId = Convert.ToInt32(reader("user_id")),
                .AdministratorCode = Convert.ToString(reader("admin_code")),
                .FirstName = Convert.ToString(reader("first_name")),
                .MiddleName = Convert.ToString(reader("middle_name")),
                .LastName = Convert.ToString(reader("last_name")),
                .RoleTitle = Convert.ToString(reader("role_title")),
                .Email = Convert.ToString(reader("email")),
                .PhotoPath = Convert.ToString(reader("photo_path"))
            }
        End Function

        Private Function BuildDisplayName(request As AdministratorSaveRequest) As String
            Dim parts As New List(Of String)()

            If Not String.IsNullOrWhiteSpace(request.FirstName) Then
                parts.Add(request.FirstName.Trim())
            End If

            If Not String.IsNullOrWhiteSpace(request.LastName) Then
                parts.Add(request.LastName.Trim())
            End If

            If Not String.IsNullOrWhiteSpace(request.MiddleName) Then
                parts.Add(request.MiddleName.Trim())
            End If

            Return String.Join(" ", parts)
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
