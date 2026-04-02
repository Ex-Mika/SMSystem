Namespace Backend.Models
    Public Class CourseRecord
        Public Property CourseId As Integer
        Public Property CourseCode As String = String.Empty
        Public Property CourseName As String = String.Empty
        Public Property DepartmentId As Integer?
        Public Property DepartmentLabel As String = String.Empty
        Public Property DepartmentCode As String = String.Empty
        Public Property DepartmentName As String = String.Empty
        Public Property Units As String = String.Empty

        Public ReadOnly Property DepartmentDisplayName As String
            Get
                If Not String.IsNullOrWhiteSpace(DepartmentLabel) Then
                    Return DepartmentLabel.Trim()
                End If

                If Not String.IsNullOrWhiteSpace(DepartmentName) Then
                    Return DepartmentName.Trim()
                End If

                Return If(DepartmentCode, String.Empty).Trim()
            End Get
        End Property
    End Class
End Namespace
