Imports School_Management_System.Backend.Common

Namespace Backend.Models
    Public Class TeacherStudentListExportTarget
        Public Property SubjectCode As String = String.Empty
        Public Property SubjectName As String = String.Empty
        Public Property SectionValue As String = String.Empty
        Public Property YearLevel As String = String.Empty

        Public ReadOnly Property SubjectLabel As String
            Get
                Dim normalizedCode As String = NormalizeText(SubjectCode)
                Dim normalizedName As String = NormalizeText(SubjectName)

                If normalizedCode <> String.Empty AndAlso
                   normalizedName <> String.Empty AndAlso
                   Not String.Equals(normalizedCode,
                                     normalizedName,
                                     StringComparison.OrdinalIgnoreCase) Then
                    Return normalizedCode & " - " & normalizedName
                End If

                If normalizedCode <> String.Empty Then
                    Return normalizedCode
                End If

                If normalizedName <> String.Empty Then
                    Return normalizedName
                End If

                Return "Untitled Subject"
            End Get
        End Property

        Public ReadOnly Property SectionLabel As String
            Get
                Return StudentScheduleHelper.BuildCompactSectionValue(
                    NormalizeText(SectionValue),
                    NormalizeText(YearLevel),
                    "Section TBA")
            End Get
        End Property

        Public ReadOnly Property DisplayLabel As String
            Get
                Return SubjectLabel & " | " & SectionLabel
            End Get
        End Property

        Private Shared Function NormalizeText(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function
    End Class
End Namespace
