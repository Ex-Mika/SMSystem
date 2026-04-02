Namespace Backend.Models
    Public Class StudentSaveRequest
        Public Property OriginalStudentNumber As String = String.Empty
        Public Property StudentNumber As String = String.Empty
        Public Property FirstName As String = String.Empty
        Public Property MiddleName As String = String.Empty
        Public Property LastName As String = String.Empty
        Public Property YearLevel As Integer?
        Public Property CourseText As String = String.Empty
        Public Property SectionName As String = String.Empty
        Public Property PhotoPath As String = String.Empty
        Public Property Password As String = String.Empty
    End Class
End Namespace
