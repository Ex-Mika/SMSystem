Imports System.Collections.Generic
Imports System.IO
Imports System.Security.Cryptography
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models

Namespace Database
    Public Module DatabaseModule
        Private Const DefaultHost As String = "127.0.0.1"
        Private Const DefaultPort As UInteger = 3306UI
        Private Const DefaultDatabaseName As String = "school_management_system"
        Private Const DefaultUserName As String = "root"
        Private Const DefaultPassword As String = "admin"
        Private Const PasswordHashIterations As Integer = 100000
        Private ReadOnly SchemaSyncRoot As New Object()
        Private _databaseBootstrapEnsured As Boolean
        Private _legacySchemaEnsured As Boolean

        Public Function CreateConnection() As MySqlConnection
            Return New MySqlConnection(CreateConnectionStringBuilder(includeDatabase:=True).ConnectionString)
        End Function

        Public Function OpenConnection() As MySqlConnection
            EnsureDatabaseBootstrap()

            Dim connection As MySqlConnection = CreateConnection()
            connection.Open()
            EnsureLegacySchema(connection)
            Return connection
        End Function

        Public Function BuildConnectionString() As String
            Return CreateConnectionStringBuilder(includeDatabase:=True).ConnectionString
        End Function

        Public Function GetDatabaseName() As String
            Return GetEnvironmentSetting("SMS_DB_NAME", DefaultDatabaseName)
        End Function

        Public Function GetRoleKey(role As UserRole) As String
            Select Case role
                Case UserRole.Student
                    Return "student"
                Case UserRole.Teacher
                    Return "teacher"
                Case Else
                    Return "admin"
            End Select
        End Function

        Public Function ParseRoleKey(roleKey As String) As UserRole
            If String.IsNullOrWhiteSpace(roleKey) Then
                Return UserRole.Admin
            End If

            Select Case roleKey.Trim().ToLowerInvariant()
                Case "student"
                    Return UserRole.Student
                Case "teacher"
                    Return UserRole.Teacher
                Case Else
                    Return UserRole.Admin
            End Select
        End Function

        Public Function HashPassword(password As String) As String
            If String.IsNullOrWhiteSpace(password) Then
                Return String.Empty
            End If

            Dim salt(15) As Byte
            RandomNumberGenerator.Fill(salt)

            Dim hashBytes As Byte() = Rfc2898DeriveBytes.Pbkdf2(password,
                                                                salt,
                                                                PasswordHashIterations,
                                                                HashAlgorithmName.SHA256,
                                                                32)

            Return $"{PasswordHashIterations}:{Convert.ToBase64String(salt)}:{Convert.ToBase64String(hashBytes)}"
        End Function

        Public Function VerifyPassword(password As String, storedHash As String) As Boolean
            If String.IsNullOrWhiteSpace(password) OrElse String.IsNullOrWhiteSpace(storedHash) Then
                Return False
            End If

            Dim parts As String() = storedHash.Split(":"c)
            If parts.Length <> 3 Then
                Return False
            End If

            Dim iterations As Integer
            If Not Integer.TryParse(parts(0), iterations) Then
                Return False
            End If

            Dim salt As Byte()
            Dim expectedHash As Byte()

            Try
                salt = Convert.FromBase64String(parts(1))
                expectedHash = Convert.FromBase64String(parts(2))
            Catch ex As FormatException
                Return False
            End Try

            Dim actualHash As Byte() = Rfc2898DeriveBytes.Pbkdf2(password,
                                                                 salt,
                                                                 iterations,
                                                                 HashAlgorithmName.SHA256,
                                                                 expectedHash.Length)

            Return CryptographicOperations.FixedTimeEquals(actualHash, expectedHash)
        End Function

        Private Function GetPort() As UInteger
            Dim rawPort As String = GetEnvironmentSetting("SMS_DB_PORT", DefaultPort.ToString())
            Dim parsedPort As UInteger

            If UInteger.TryParse(rawPort, parsedPort) Then
                Return parsedPort
            End If

            Return DefaultPort
        End Function

        Private Function CreateConnectionStringBuilder(includeDatabase As Boolean) As MySqlConnectionStringBuilder
            Dim connectionBuilder As New MySqlConnectionStringBuilder() With {
                .Server = GetEnvironmentSetting("SMS_DB_HOST", DefaultHost),
                .Port = GetPort(),
                .UserID = GetEnvironmentSetting("SMS_DB_USER", DefaultUserName),
                .Password = GetEnvironmentSetting("SMS_DB_PASSWORD", DefaultPassword),
                .SslMode = MySqlSslMode.Disabled,
                .AllowUserVariables = True,
                .ConvertZeroDateTime = True
            }

            If includeDatabase Then
                connectionBuilder.Database = GetDatabaseName()
            End If

            Return connectionBuilder
        End Function

        Private Function CreateServerConnection() As MySqlConnection
            Return New MySqlConnection(CreateConnectionStringBuilder(includeDatabase:=False).ConnectionString)
        End Function

        Private Sub EnsureDatabaseBootstrap()
            If _databaseBootstrapEnsured Then
                Return
            End If

            SyncLock SchemaSyncRoot
                If _databaseBootstrapEnsured Then
                    Return
                End If

                Using connection As MySqlConnection = CreateServerConnection()
                    connection.Open()
                    ExecuteSqlScripts(connection, "Schema")

                    If ShouldRunSeedScripts(connection) Then
                        ExecuteSqlScripts(connection, "Seeds")
                    End If
                End Using

                _databaseBootstrapEnsured = True
            End SyncLock
        End Sub

        Private Sub EnsureLegacySchema(connection As MySqlConnection)
            If connection Is Nothing OrElse _legacySchemaEnsured Then
                Return
            End If

            SyncLock SchemaSyncRoot
                If _legacySchemaEnsured Then
                    Return
                End If

                ' Keep older local databases compatible with the current repository queries.
                EnsureOptionalColumns(connection,
                                      "departments",
                                      {
                                          Tuple.Create(
                                              "head_name",
                                              "ALTER TABLE `departments` ADD COLUMN `head_name` VARCHAR(150) NULL AFTER `department_name`;")
                                      })

                EnsureOptionalColumns(connection,
                                      "students",
                                      {
                                          Tuple.Create(
                                              "middle_name",
                                              "ALTER TABLE `students` ADD COLUMN `middle_name` VARCHAR(100) NULL AFTER `first_name`;"),
                                          Tuple.Create(
                                              "year_level",
                                              "ALTER TABLE `students` ADD COLUMN `year_level` TINYINT NULL AFTER `course_id`;"),
                                          Tuple.Create(
                                              "section_name",
                                              "ALTER TABLE `students` ADD COLUMN `section_name` VARCHAR(60) NULL AFTER `year_level`;"),
                                          Tuple.Create(
                                              "photo_path",
                                              "ALTER TABLE `students` ADD COLUMN `photo_path` VARCHAR(500) NULL AFTER `section_name`;")
                                      })

                EnsureOptionalColumns(connection,
                                      "administrators",
                                      {
                                          Tuple.Create(
                                              "middle_name",
                                              "ALTER TABLE `administrators` ADD COLUMN `middle_name` VARCHAR(100) NULL AFTER `first_name`;"),
                                          Tuple.Create(
                                              "role_title",
                                              "ALTER TABLE `administrators` ADD COLUMN `role_title` VARCHAR(100) NULL AFTER `last_name`;"),
                                          Tuple.Create(
                                              "photo_path",
                                              "ALTER TABLE `administrators` ADD COLUMN `photo_path` VARCHAR(500) NULL AFTER `role_title`;")
                                      })

                EnsureOptionalColumns(connection,
                                      "teachers",
                                      {
                                          Tuple.Create(
                                              "middle_name",
                                              "ALTER TABLE `teachers` ADD COLUMN `middle_name` VARCHAR(100) NULL AFTER `first_name`;"),
                                          Tuple.Create(
                                              "department_label",
                                              "ALTER TABLE `teachers` ADD COLUMN `department_label` VARCHAR(120) NULL AFTER `department_id`;"),
                                          Tuple.Create(
                                              "position_title",
                                              "ALTER TABLE `teachers` ADD COLUMN `position_title` VARCHAR(100) NULL AFTER `department_label`;"),
                                          Tuple.Create(
                                              "advisory_section",
                                              "ALTER TABLE `teachers` ADD COLUMN `advisory_section` VARCHAR(60) NULL AFTER `position_title`;"),
                                          Tuple.Create(
                                              "photo_path",
                                              "ALTER TABLE `teachers` ADD COLUMN `photo_path` VARCHAR(500) NULL AFTER `advisory_section`;")
                                      })

                _legacySchemaEnsured = True
            End SyncLock
        End Sub

        Private Sub ExecuteSqlScripts(connection As MySqlConnection, scriptFolderName As String)
            For Each scriptPath As String In GetSqlScriptPaths(scriptFolderName)
                ExecuteSqlScript(connection, scriptPath)
            Next
        End Sub

        Private Function GetSqlScriptPaths(scriptFolderName As String) As String()
            Dim scriptFolderPath As String =
                Path.Combine(AppContext.BaseDirectory, "Database", scriptFolderName)

            If Not Directory.Exists(scriptFolderPath) Then
                Throw New DirectoryNotFoundException(
                    "Database script folder was not found: " & scriptFolderPath)
            End If

            Dim scriptPaths As String() =
                Directory.GetFiles(scriptFolderPath, "*.sql", SearchOption.TopDirectoryOnly)
            Array.Sort(scriptPaths, StringComparer.OrdinalIgnoreCase)
            Return scriptPaths
        End Function

        Private Sub ExecuteSqlScript(connection As MySqlConnection, scriptPath As String)
            Dim scriptText As String = File.ReadAllText(scriptPath)
            If String.IsNullOrWhiteSpace(scriptText) Then
                Return
            End If

            Dim resolvedDatabaseName As String = EscapeMySqlIdentifier(GetDatabaseName())
            scriptText = scriptText.Replace("`" & DefaultDatabaseName & "`",
                                            "`" & resolvedDatabaseName & "`")

            Dim script As New MySqlScript(connection, scriptText)
            script.Execute()
        End Sub

        Private Function ShouldRunSeedScripts(connection As MySqlConnection) As Boolean
            If connection Is Nothing OrElse Not TableExists(connection, "users") Then
                Return False
            End If

            Using command As New MySqlCommand("SELECT COUNT(*) FROM `users`;", connection)
                Return Convert.ToInt32(command.ExecuteScalar()) = 0
            End Using
        End Function

        Private Sub EnsureOptionalColumns(connection As MySqlConnection,
                                          tableName As String,
                                          columnDefinitions As IEnumerable(Of Tuple(Of String, String)))
            If connection Is Nothing OrElse
               String.IsNullOrWhiteSpace(tableName) OrElse
               columnDefinitions Is Nothing OrElse
               Not TableExists(connection, tableName) Then
                Return
            End If

            For Each columnDefinition As Tuple(Of String, String) In columnDefinitions
                If columnDefinition Is Nothing OrElse
                   String.IsNullOrWhiteSpace(columnDefinition.Item1) OrElse
                   String.IsNullOrWhiteSpace(columnDefinition.Item2) OrElse
                   ColumnExists(connection, tableName, columnDefinition.Item1) Then
                    Continue For
                End If

                Using alterCommand As New MySqlCommand(columnDefinition.Item2, connection)
                    alterCommand.ExecuteNonQuery()
                End Using
            Next
        End Sub

        Private Function TableExists(connection As MySqlConnection, tableName As String) As Boolean
            Using command As New MySqlCommand(
                "SELECT COUNT(*) " &
                "FROM information_schema.TABLES " &
                "WHERE TABLE_SCHEMA = DATABASE() " &
                "AND TABLE_NAME = @tableName;",
                connection)
                command.Parameters.AddWithValue("@tableName", tableName)
                Return Convert.ToInt32(command.ExecuteScalar()) > 0
            End Using
        End Function

        Private Function ColumnExists(connection As MySqlConnection,
                                      tableName As String,
                                      columnName As String) As Boolean
            Using command As New MySqlCommand(
                "SELECT COUNT(*) " &
                "FROM information_schema.COLUMNS " &
                "WHERE TABLE_SCHEMA = DATABASE() " &
                "AND TABLE_NAME = @tableName " &
                "AND COLUMN_NAME = @columnName;",
                connection)
                command.Parameters.AddWithValue("@tableName", tableName)
                command.Parameters.AddWithValue("@columnName", columnName)
                Return Convert.ToInt32(command.ExecuteScalar()) > 0
            End Using
        End Function

        Private Function GetEnvironmentSetting(variableName As String, defaultValue As String) As String
            Dim value As String = Environment.GetEnvironmentVariable(variableName)

            If String.IsNullOrWhiteSpace(value) Then
                Return defaultValue
            End If

            Return value.Trim()
        End Function

        Private Function EscapeMySqlIdentifier(value As String) As String
            Return If(value, String.Empty).Replace("`", "``")
        End Function
    End Module
End Namespace
