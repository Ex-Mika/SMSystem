Imports System.Globalization
Imports School_Management_System.Backend.Common

Namespace Backend.Models
    Public Class StudentSubjectGradeRecord
        Public Property GradeRecordId As Integer
        Public Property TeacherRecordId As Integer
        Public Property TeacherId As String = String.Empty
        Public Property TeacherName As String = String.Empty
        Public Property StudentRecordId As Integer
        Public Property StudentNumber As String = String.Empty
        Public Property StudentName As String = String.Empty
        Public Property StudentSection As String = String.Empty
        Public Property StudentYearLevel As String = String.Empty
        Public Property SubjectId As Integer
        Public Property SubjectCode As String = String.Empty
        Public Property SubjectName As String = String.Empty
        Public Property SubjectUnits As String = String.Empty
        Public Property SectionName As String = String.Empty
        Public Property QuizScore As Decimal?
        Public Property ProjectScore As Decimal?
        Public Property MidtermScore As Decimal?
        Public Property FinalExamScore As Decimal?
        Public Property FinalGrade As Decimal?
        Public Property Remarks As String = String.Empty
        Public Property UpdatedAt As DateTime?

        Public ReadOnly Property HasPublishedGrade As Boolean
            Get
                Return GradeRecordId > 0 AndAlso
                    FinalGrade.HasValue
            End Get
        End Property

        Public ReadOnly Property HasPassingGrade As Boolean
            Get
                Return HasPublishedGrade AndAlso
                    FinalGrade.Value >= 75D
            End Get
        End Property

        Public ReadOnly Property SubjectDisplayLabel As String
            Get
                Dim subjectCodeValue As String = NormalizeText(SubjectCode)
                Dim subjectNameValue As String = NormalizeText(SubjectName)

                If subjectCodeValue <> String.Empty AndAlso
                   subjectNameValue <> String.Empty AndAlso
                   Not String.Equals(subjectCodeValue,
                                     subjectNameValue,
                                     StringComparison.OrdinalIgnoreCase) Then
                    Return subjectCodeValue & " - " & subjectNameValue
                End If

                If subjectCodeValue <> String.Empty Then
                    Return subjectCodeValue
                End If

                If subjectNameValue <> String.Empty Then
                    Return subjectNameValue
                End If

                Return "--"
            End Get
        End Property

        Public ReadOnly Property SectionDisplayLabel As String
            Get
                Dim sectionValue As String = NormalizeText(SectionName)
                If sectionValue <> String.Empty Then
                    Return sectionValue
                End If

                Return NormalizeText(StudentSection, "--")
            End Get
        End Property

        Public ReadOnly Property YearSectionDisplayLabel As String
            Get
                Dim resolvedSection As String = NormalizeText(SectionName)
                If resolvedSection = String.Empty Then
                    resolvedSection = NormalizeText(StudentSection)
                End If

                Return StudentScheduleHelper.BuildCompactSectionValue(
                    resolvedSection,
                    NormalizeText(StudentYearLevel),
                    "--")
            End Get
        End Property

        Public ReadOnly Property QuizScoreLabel As String
            Get
                Return FormatNullableDecimal(QuizScore)
            End Get
        End Property

        Public ReadOnly Property ProjectScoreLabel As String
            Get
                Return FormatNullableDecimal(ProjectScore)
            End Get
        End Property

        Public ReadOnly Property MidtermScoreLabel As String
            Get
                Return FormatNullableDecimal(MidtermScore)
            End Get
        End Property

        Public ReadOnly Property FinalExamScoreLabel As String
            Get
                Return FormatNullableDecimal(FinalExamScore)
            End Get
        End Property

        Public ReadOnly Property FinalGradeLabel As String
            Get
                Return FormatNullableDecimal(FinalGrade)
            End Get
        End Property

        Public ReadOnly Property RemarksLabel As String
            Get
                Dim remarksValue As String = NormalizeText(Remarks)
                If remarksValue <> String.Empty Then
                    Return remarksValue
                End If

                If Not HasPublishedGrade Then
                    Return "Pending"
                End If

                Return If(HasPassingGrade, "Passed", "Failed")
            End Get
        End Property

        Public ReadOnly Property UpdatedAtLabel As String
            Get
                If Not UpdatedAt.HasValue Then
                    Return "--"
                End If

                Return UpdatedAt.Value.ToString("MMM dd, yyyy hh:mm tt",
                                                CultureInfo.InvariantCulture)
            End Get
        End Property

        Public ReadOnly Property UnitsValue As Decimal
            Get
                Dim normalizedUnits As String = NormalizeText(SubjectUnits)
                If normalizedUnits = String.Empty Then
                    Return 0D
                End If

                Dim parsedUnits As Decimal
                If Decimal.TryParse(normalizedUnits,
                                    NumberStyles.Number,
                                    CultureInfo.InvariantCulture,
                                    parsedUnits) OrElse
                   Decimal.TryParse(normalizedUnits,
                                    NumberStyles.Number,
                                    CultureInfo.CurrentCulture,
                                    parsedUnits) Then
                    Return parsedUnits
                End If

                Return 0D
            End Get
        End Property

        Private Function FormatNullableDecimal(value As Decimal?) As String
            If Not value.HasValue Then
                Return "--"
            End If

            Return value.Value.ToString("0.00", CultureInfo.InvariantCulture)
        End Function

        Private Function NormalizeText(value As String,
                                       Optional fallbackValue As String = "") As String
            Dim normalizedValue As String = If(value, String.Empty).Trim()
            If normalizedValue = String.Empty Then
                Return fallbackValue
            End If

            Return normalizedValue
        End Function
    End Class
End Namespace
