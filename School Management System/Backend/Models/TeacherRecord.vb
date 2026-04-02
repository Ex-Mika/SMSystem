Namespace Backend.Models
    Public Class TeacherRecord
        Public Property TeacherRecordId As Integer
        Public Property UserId As Integer
        Public Property EmployeeNumber As String = String.Empty
        Public Property FirstName As String = String.Empty
        Public Property MiddleName As String = String.Empty
        Public Property LastName As String = String.Empty
        Public Property DepartmentId As Integer?
        Public Property DepartmentLabel As String = String.Empty
        Public Property DepartmentCode As String = String.Empty
        Public Property DepartmentName As String = String.Empty
        Public Property PositionTitle As String = String.Empty
        Public Property AdvisorySection As String = String.Empty
        Public Property PhotoPath As String = String.Empty
        Public Property Email As String = String.Empty

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
