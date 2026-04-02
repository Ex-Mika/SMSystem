Namespace Backend.Models
    Public Class UserAccount
        Public Property UserId As Integer
        Public Property Role As UserRole
        Public Property Username As String = String.Empty
        Public Property Email As String = String.Empty
        Public Property PasswordHash As String = String.Empty
        Public Property IsActive As Boolean
        Public Property ReferenceCode As String = String.Empty
    End Class
End Namespace
