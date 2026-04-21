Imports System.Collections.Generic
Imports System.Globalization

Namespace Backend.Models
    Public Class StudentGradeSnapshot
        Public Property Student As StudentRecord
        Public Property Grades As New List(Of StudentSubjectGradeRecord)()
        Public Property NoticeMessage As String = String.Empty

        Public ReadOnly Property PostedGradeCount As Integer
            Get
                Dim count As Integer = 0

                For Each grade As StudentSubjectGradeRecord In Grades
                    If grade IsNot Nothing AndAlso grade.HasPublishedGrade Then
                        count += 1
                    End If
                Next

                Return count
            End Get
        End Property

        Public ReadOnly Property HasGrades As Boolean
            Get
                Return PostedGradeCount > 0
            End Get
        End Property

        Public ReadOnly Property RatedUnits As Decimal
            Get
                Dim total As Decimal = 0D

                For Each grade As StudentSubjectGradeRecord In Grades
                    If grade IsNot Nothing AndAlso grade.HasPublishedGrade Then
                        total += grade.UnitsValue
                    End If
                Next

                Return total
            End Get
        End Property

        Public ReadOnly Property RatedUnitsLabel As String
            Get
                Return RatedUnits.ToString("0.#", CultureInfo.InvariantCulture)
            End Get
        End Property

        Public ReadOnly Property EarnedUnits As Decimal
            Get
                Dim total As Decimal = 0D

                For Each grade As StudentSubjectGradeRecord In Grades
                    If grade IsNot Nothing AndAlso grade.HasPassingGrade Then
                        total += grade.UnitsValue
                    End If
                Next

                Return total
            End Get
        End Property

        Public ReadOnly Property EarnedUnitsLabel As String
            Get
                Return EarnedUnits.ToString("0.#", CultureInfo.InvariantCulture)
            End Get
        End Property

        Public ReadOnly Property PassedSubjectCount As Integer
            Get
                Dim count As Integer = 0

                For Each grade As StudentSubjectGradeRecord In Grades
                    If grade IsNot Nothing AndAlso grade.HasPassingGrade Then
                        count += 1
                    End If
                Next

                Return count
            End Get
        End Property

        Public ReadOnly Property CurrentGwa As Decimal
            Get
                If Not HasGrades Then
                    Return 0D
                End If

                Dim totalWeightedGrades As Decimal = 0D
                Dim totalUnits As Decimal = 0D

                For Each grade As StudentSubjectGradeRecord In Grades
                    If grade Is Nothing OrElse
                       Not grade.HasPublishedGrade OrElse
                       Not grade.FinalGrade.HasValue Then
                        Continue For
                    End If

                    Dim unitsValue As Decimal = grade.UnitsValue
                    If unitsValue <= 0D Then
                        unitsValue = 1D
                    End If

                    totalWeightedGrades += grade.FinalGrade.Value * unitsValue
                    totalUnits += unitsValue
                Next

                If totalUnits <= 0D Then
                    Return 0D
                End If

                Return Decimal.Round(totalWeightedGrades / totalUnits,
                                     2,
                                     MidpointRounding.AwayFromZero)
            End Get
        End Property

        Public ReadOnly Property CurrentGwaLabel As String
            Get
                Return CurrentGwa.ToString("0.00", CultureInfo.InvariantCulture)
            End Get
        End Property

        Public ReadOnly Property QuizAverage As Decimal
            Get
                Return ComputeAverage(Function(grade) grade.QuizScore)
            End Get
        End Property

        Public ReadOnly Property QuizAverageLabel As String
            Get
                Return QuizAverage.ToString("0.00", CultureInfo.InvariantCulture)
            End Get
        End Property

        Public ReadOnly Property ProjectAverage As Decimal
            Get
                Return ComputeAverage(Function(grade) grade.ProjectScore)
            End Get
        End Property

        Public ReadOnly Property ProjectAverageLabel As String
            Get
                Return ProjectAverage.ToString("0.00", CultureInfo.InvariantCulture)
            End Get
        End Property

        Public ReadOnly Property MidtermAverage As Decimal
            Get
                Return ComputeAverage(Function(grade) grade.MidtermScore)
            End Get
        End Property

        Public ReadOnly Property MidtermAverageLabel As String
            Get
                Return MidtermAverage.ToString("0.00", CultureInfo.InvariantCulture)
            End Get
        End Property

        Public ReadOnly Property FinalExamAverage As Decimal
            Get
                Return ComputeAverage(Function(grade) grade.FinalExamScore)
            End Get
        End Property

        Public ReadOnly Property FinalExamAverageLabel As String
            Get
                Return FinalExamAverage.ToString("0.00", CultureInfo.InvariantCulture)
            End Get
        End Property

        Public ReadOnly Property TermCount As Integer
            Get
                Return If(HasGrades, 1, 0)
            End Get
        End Property

        Public ReadOnly Property AcademicStandingLabel As String
            Get
                If Not HasGrades Then
                    Return "Pending"
                End If

                If PassedSubjectCount = PostedGradeCount Then
                    Return "Good"
                End If

                If PassedSubjectCount = 0 Then
                    Return "At Risk"
                End If

                Return "Needs Attention"
            End Get
        End Property

        Private Function ComputeAverage(selector As Func(Of StudentSubjectGradeRecord, Decimal?)) As Decimal
            Dim total As Decimal = 0D
            Dim count As Integer = 0

            If selector Is Nothing Then
                Return 0D
            End If

            For Each grade As StudentSubjectGradeRecord In Grades
                If grade Is Nothing OrElse Not grade.HasPublishedGrade Then
                    Continue For
                End If

                Dim scoreValue As Decimal? = selector(grade)
                If Not scoreValue.HasValue Then
                    Continue For
                End If

                total += scoreValue.Value
                count += 1
            Next

            If count = 0 Then
                Return 0D
            End If

            Return Decimal.Round(total / count,
                                 2,
                                 MidpointRounding.AwayFromZero)
        End Function
    End Class
End Namespace
