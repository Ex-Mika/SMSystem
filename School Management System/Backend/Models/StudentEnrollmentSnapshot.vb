Imports System.Collections.Generic
Imports System.Globalization

Namespace Backend.Models
    Public Class StudentEnrollmentSnapshot
        Public Property Student As StudentRecord
        Public Property SelectedSubjects As New List(Of SubjectRecord)()
        Public Property AvailableSubjects As New List(Of SubjectRecord)()
        Public Property AvailableSections As New List(Of StudentSectionOption)()
        Public Property NoticeMessage As String = String.Empty

        Public ReadOnly Property SelectedSubjectCount As Integer
            Get
                Return CountSubjects(SelectedSubjects)
            End Get
        End Property

        Public ReadOnly Property AvailableSubjectCount As Integer
            Get
                Return CountSubjects(AvailableSubjects)
            End Get
        End Property

        Public ReadOnly Property HasAssignedYearLevel As Boolean
            Get
                Return Student IsNot Nothing AndAlso Student.YearLevel.HasValue
            End Get
        End Property

        Public ReadOnly Property HasAssignedSection As Boolean
            Get
                Return Student IsNot Nothing AndAlso
                    Not String.IsNullOrWhiteSpace(If(Student.SectionName, String.Empty).Trim())
            End Get
        End Property

        Public ReadOnly Property HasSelectedSubjects As Boolean
            Get
                Return SelectedSubjectCount > 0
            End Get
        End Property

        Public ReadOnly Property HasRemainingSubjects As Boolean
            Get
                Return AvailableSubjectCount > 0
            End Get
        End Property

        Public ReadOnly Property IsReadyForReview As Boolean
            Get
                Return HasAssignedYearLevel AndAlso
                    HasAssignedSection AndAlso
                    HasSelectedSubjects
            End Get
        End Property

        Public ReadOnly Property IsFullyEnrolled As Boolean
            Get
                Return IsReadyForReview AndAlso Not HasRemainingSubjects
            End Get
        End Property

        Public ReadOnly Property SelectedTotalUnits As Decimal
            Get
                Return SumSubjectUnits(SelectedSubjects)
            End Get
        End Property

        Public ReadOnly Property SelectedTotalUnitsLabel As String
            Get
                Return SelectedTotalUnits.ToString("0.#", CultureInfo.InvariantCulture)
            End Get
        End Property

        Private Function CountSubjects(subjects As IEnumerable(Of SubjectRecord)) As Integer
            Dim count As Integer = 0

            If subjects Is Nothing Then
                Return count
            End If

            For Each subject As SubjectRecord In subjects
                If subject IsNot Nothing Then
                    count += 1
                End If
            Next

            Return count
        End Function

        Private Function SumSubjectUnits(subjects As IEnumerable(Of SubjectRecord)) As Decimal
            Dim total As Decimal = 0D

            If subjects Is Nothing Then
                Return total
            End If

            For Each subject As SubjectRecord In subjects
                If subject Is Nothing Then
                    Continue For
                End If

                total += ParseUnits(subject.Units)
            Next

            Return total
        End Function

        Private Function ParseUnits(value As String) As Decimal
            Dim normalizedValue As String = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Return 0D
            End If

            Dim parsedValue As Decimal
            If Decimal.TryParse(normalizedValue,
                                NumberStyles.Number,
                                CultureInfo.InvariantCulture,
                                parsedValue) OrElse
               Decimal.TryParse(normalizedValue,
                                NumberStyles.Number,
                                CultureInfo.CurrentCulture,
                                parsedValue) Then
                Return parsedValue
            End If

            Return 0D
        End Function
    End Class
End Namespace
