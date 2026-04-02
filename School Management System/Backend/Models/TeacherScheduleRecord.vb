Namespace Backend.Models
    Public Class TeacherScheduleRecord
        Public Property ScheduleId As Integer
        Public Property TeacherRecordId As Integer
        Public Property TeacherId As String = String.Empty
        Public Property TeacherName As String = String.Empty
        Public Property Day As String = String.Empty
        Public Property Session As String = String.Empty
        Public Property Section As String = String.Empty
        Public Property SubjectCode As String = String.Empty
        Public Property SubjectName As String = String.Empty
        Public Property Room As String = String.Empty
    End Class
End Namespace
