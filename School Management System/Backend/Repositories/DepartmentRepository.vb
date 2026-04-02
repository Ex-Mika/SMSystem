Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models

Namespace Backend.Repositories
    Public Class DepartmentRepository
        Private Const SelectColumnsSql As String =
            "SELECT " &
            "d.department_id, " &
            "d.department_code, " &
            "d.department_name, " &
            "COALESCE(d.head_name, '') AS head_name " &
            "FROM departments d "

        Public Function GetAll() As List(Of DepartmentRecord)
            Dim departments As New List(Of DepartmentRecord)()

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "ORDER BY d.department_code, d.department_name;",
                    connection)
                    Using reader As MySqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            departments.Add(MapDepartment(reader))
                        End While
                    End Using
                End Using
            End Using

            Return departments
        End Function

        Public Function GetByDepartmentCode(departmentCode As String) As DepartmentRecord
            If String.IsNullOrWhiteSpace(departmentCode) Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    SelectColumnsSql & "WHERE d.department_code = @departmentCode LIMIT 1;",
                    connection)
                    command.Parameters.AddWithValue("@departmentCode", departmentCode.Trim())

                    Using reader As MySqlDataReader = command.ExecuteReader()
                        If Not reader.Read() Then
                            Return Nothing
                        End If

                        Return MapDepartment(reader)
                    End Using
                End Using
            End Using
        End Function

        Public Function Create(request As DepartmentSaveRequest) As DepartmentRecord
            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "INSERT INTO departments (" &
                    "department_code, department_name, head_name" &
                    ") VALUES (" &
                    "@departmentCode, @departmentName, @headName" &
                    ");",
                    connection)
                    command.Parameters.AddWithValue("@departmentCode", request.DepartmentCode.Trim())
                    command.Parameters.AddWithValue("@departmentName", request.DepartmentName.Trim())
                    command.Parameters.AddWithValue("@headName", NormalizeNullableValue(request.HeadName))
                    command.ExecuteNonQuery()
                End Using
            End Using

            Return GetByDepartmentCode(request.DepartmentCode)
        End Function

        Public Function Update(existingRecord As DepartmentRecord,
                               request As DepartmentSaveRequest) As DepartmentRecord
            If existingRecord Is Nothing Then
                Return Nothing
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "UPDATE departments " &
                    "SET department_code = @departmentCode, " &
                    "department_name = @departmentName, " &
                    "head_name = @headName " &
                    "WHERE department_id = @departmentRecordId;",
                    connection)
                    command.Parameters.AddWithValue("@departmentCode", request.DepartmentCode.Trim())
                    command.Parameters.AddWithValue("@departmentName", request.DepartmentName.Trim())
                    command.Parameters.AddWithValue("@headName", NormalizeNullableValue(request.HeadName))
                    command.Parameters.AddWithValue("@departmentRecordId", existingRecord.DepartmentRecordId)
                    command.ExecuteNonQuery()
                End Using
            End Using

            Return GetByDepartmentCode(request.DepartmentCode)
        End Function

        Public Function DeleteByDepartmentCode(departmentCode As String) As Boolean
            If String.IsNullOrWhiteSpace(departmentCode) Then
                Return False
            End If

            Dim existingRecord As DepartmentRecord = GetByDepartmentCode(departmentCode)
            If existingRecord Is Nothing Then
                Return False
            End If

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                Using command As New MySqlCommand(
                    "DELETE FROM departments WHERE department_id = @departmentRecordId;",
                    connection)
                    command.Parameters.AddWithValue("@departmentRecordId", existingRecord.DepartmentRecordId)
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Private Function MapDepartment(reader As MySqlDataReader) As DepartmentRecord
            Return New DepartmentRecord() With {
                .DepartmentRecordId = Convert.ToInt32(reader("department_id")),
                .DepartmentCode = Convert.ToString(reader("department_code")),
                .DepartmentName = Convert.ToString(reader("department_name")),
                .HeadName = Convert.ToString(reader("head_name"))
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
