Imports System.Collections.Generic
Imports School_Management_System.Backend.Models

Class StudentGradeBreakdownTab
    Private NotInheritable Class GradeBreakdownListItem
        Public Property SubjectName As String = String.Empty
        Public Property QuizScore As String = String.Empty
        Public Property ProjectScore As String = String.Empty
        Public Property MidtermScore As String = String.Empty
        Public Property FinalExamScore As String = String.Empty
    End Class

    Public Sub New()
        InitializeComponent()
        SetGradeSnapshot(Nothing)
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        GradeBreakdownStudentIdTextBlock.Text = If(studentId, String.Empty).Trim()
    End Sub

    Public Sub SetGradeSnapshot(snapshot As StudentGradeSnapshot,
                                Optional statusMessage As String = "")
        Dim resolvedSnapshot As StudentGradeSnapshot =
            If(snapshot, New StudentGradeSnapshot())
        Dim breakdownItems As List(Of GradeBreakdownListItem) =
            BuildBreakdownItems(resolvedSnapshot)

        QuizAverageTextBlock.Text = resolvedSnapshot.QuizAverageLabel
        ProjectAverageTextBlock.Text = resolvedSnapshot.ProjectAverageLabel
        ExamAverageTextBlock.Text = resolvedSnapshot.FinalExamAverageLabel
        GradeBreakdownStatusTextBlock.Text =
            BuildStatusMessage(resolvedSnapshot, statusMessage)
        GradeBreakdownEmptyTextBlock.Text =
            BuildEmptyStateMessage(resolvedSnapshot, statusMessage)

        GradeBreakdownItemsControl.ItemsSource = breakdownItems
        GradeBreakdownItemsControl.Visibility =
            If(breakdownItems.Count = 0, Visibility.Collapsed, Visibility.Visible)
        GradeBreakdownEmptyStateBorder.Visibility =
            If(breakdownItems.Count = 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Function BuildBreakdownItems(snapshot As StudentGradeSnapshot) As List(Of GradeBreakdownListItem)
        Dim items As New List(Of GradeBreakdownListItem)()
        If snapshot Is Nothing OrElse snapshot.Grades Is Nothing Then
            Return items
        End If

        For Each grade As StudentSubjectGradeRecord In snapshot.Grades
            If grade Is Nothing OrElse Not grade.HasPublishedGrade Then
                Continue For
            End If

            items.Add(New GradeBreakdownListItem() With {
                .SubjectName = grade.SubjectDisplayLabel,
                .QuizScore = grade.QuizScoreLabel,
                .ProjectScore = grade.ProjectScoreLabel,
                .MidtermScore = grade.MidtermScoreLabel,
                .FinalExamScore = grade.FinalExamScoreLabel
            })
        Next

        Return items
    End Function

    Private Function BuildStatusMessage(snapshot As StudentGradeSnapshot,
                                        statusMessage As String) As String
        Dim normalizedStatusMessage As String = If(statusMessage, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(normalizedStatusMessage) Then
            Return normalizedStatusMessage
        End If

        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "Load your student record to review your score components."
        End If

        If Not snapshot.HasGrades Then
            Return snapshot.NoticeMessage
        End If

        Return "Component scores are available for " &
            snapshot.PostedGradeCount.ToString() &
            " posted subject(s)."
    End Function

    Private Function BuildEmptyStateMessage(snapshot As StudentGradeSnapshot,
                                            statusMessage As String) As String
        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "No breakdown data loaded."
        End If

        If Not snapshot.HasGrades AndAlso
           Not String.IsNullOrWhiteSpace(snapshot.NoticeMessage) Then
            Return snapshot.NoticeMessage
        End If

        Return BuildStatusMessage(snapshot, statusMessage)
    End Function
End Class
