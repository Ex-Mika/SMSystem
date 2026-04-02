Namespace Backend.Models
    Public Class LoginRequest
        Public Property Role As UserRole
        Public Property Identifier As String = String.Empty
        Public Property Password As String = String.Empty
    End Class
End Namespace
