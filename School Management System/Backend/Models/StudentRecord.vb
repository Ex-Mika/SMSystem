Imports System.Globalization

Namespace Backend.Models
    Public Class StudentRecord
        Public Property StudentRecordId As Integer
        Public Property UserId As Integer
        Public Property StudentNumber As String = String.Empty
        Public Property FirstName As String = String.Empty
        Public Property MiddleName As String = String.Empty
        Public Property LastName As String = String.Empty
        Public Property YearLevel As Integer?
        Public Property CourseId As Integer?
        Public Property CourseCode As String = String.Empty
        Public Property CourseName As String = String.Empty
        Public Property SectionName As String = String.Empty
        Public Property PhotoPath As String = String.Empty
        Public Property Email As String = String.Empty

        Public ReadOnly Property FullName As String
            Get
                Dim parts As New List(Of String)()

                If Not String.IsNullOrWhiteSpace(FirstName) Then
                    parts.Add(FirstName.Trim())
                End If

                If Not String.IsNullOrWhiteSpace(MiddleName) Then
                    parts.Add(MiddleName.Trim())
                End If

                If Not String.IsNullOrWhiteSpace(LastName) Then
                    parts.Add(LastName.Trim())
                End If

                Return String.Join(" ", parts)
            End Get
        End Property

        Public ReadOnly Property CourseDisplayName As String
            Get
                If Not String.IsNullOrWhiteSpace(CourseCode) Then
                    Return CourseCode.Trim()
                End If

                Return If(CourseName, String.Empty).Trim()
            End Get
        End Property

        Public ReadOnly Property YearLevelLabel As String
            Get
                If Not YearLevel.HasValue Then
                    Return String.Empty
                End If

                Return YearLevel.Value.ToString(CultureInfo.InvariantCulture)
            End Get
        End Property
    End Class
End Namespace
