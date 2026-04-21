Namespace Backend.Models
    Public Class StudentGradeSaveRequest
        Public Property TeacherId As String = String.Empty
        Public Property StudentNumber As String = String.Empty
        Public Property SubjectCode As String = String.Empty
        Public Property QuizScore As Decimal
        Public Property ProjectScore As Decimal
        Public Property MidtermScore As Decimal
        Public Property FinalExamScore As Decimal
    End Class
End Namespace
