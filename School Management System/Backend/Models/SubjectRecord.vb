Namespace Backend.Models
    Public Class SubjectRecord
        Public Property SubjectId As Integer
        Public Property SubjectCode As String = String.Empty
        Public Property SubjectName As String = String.Empty
        Public Property Units As String = String.Empty
        Public Property DepartmentId As Integer?
        Public Property DepartmentCode As String = String.Empty
        Public Property DepartmentName As String = String.Empty
        Public Property CourseId As Integer?
        Public Property CourseLabel As String = String.Empty
        Public Property CourseCode As String = String.Empty
        Public Property CourseName As String = String.Empty
        Public Property YearLevel As String = String.Empty
        Public Property Description As String = String.Empty

        Public ReadOnly Property DepartmentDisplayName As String
            Get
                If Not String.IsNullOrWhiteSpace(DepartmentName) Then
                    Return DepartmentName.Trim()
                End If

                Return If(DepartmentCode, String.Empty).Trim()
            End Get
        End Property

        Public ReadOnly Property CourseDisplayName As String
            Get
                If Not String.IsNullOrWhiteSpace(CourseLabel) Then
                    Return CourseLabel.Trim()
                End If

                If Not String.IsNullOrWhiteSpace(CourseCode) Then
                    Return CourseCode.Trim()
                End If

                Return If(CourseName, String.Empty).Trim()
            End Get
        End Property
    End Class
End Namespace
