Namespace Backend.Models
    Public Class AdministratorRecord
        Public Property AdministratorRecordId As Integer
        Public Property UserId As Integer
        Public Property AdministratorCode As String = String.Empty
        Public Property FirstName As String = String.Empty
        Public Property MiddleName As String = String.Empty
        Public Property LastName As String = String.Empty
        Public Property RoleTitle As String = String.Empty
        Public Property Email As String = String.Empty
        Public Property PhotoPath As String = String.Empty

        Public ReadOnly Property FullName As String
            Get
                Dim parts As New List(Of String)()

                If Not String.IsNullOrWhiteSpace(FirstName) Then
                    parts.Add(FirstName.Trim())
                End If

                If Not String.IsNullOrWhiteSpace(LastName) Then
                    parts.Add(LastName.Trim())
                End If

                If Not String.IsNullOrWhiteSpace(MiddleName) Then
                    parts.Add(MiddleName.Trim())
                End If

                Return String.Join(" ", parts)
            End Get
        End Property
    End Class
End Namespace
