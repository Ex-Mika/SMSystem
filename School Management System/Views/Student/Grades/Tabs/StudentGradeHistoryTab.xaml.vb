Imports System.Collections.Generic
Imports School_Management_System.Backend.Models

Class StudentGradeHistoryTab
    Private NotInheritable Class GradeHistoryListItem
        Public Property TermLabel As String = String.Empty
        Public Property Units As String = String.Empty
        Public Property Gwa As String = String.Empty
        Public Property Standing As String = String.Empty
    End Class

    Public Sub New()
        InitializeComponent()
        SetGradeSnapshot(Nothing)
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        GradeHistoryStudentIdTextBlock.Text = If(studentId, String.Empty).Trim()
    End Sub

    Public Sub SetGradeSnapshot(snapshot As StudentGradeSnapshot,
                                Optional statusMessage As String = "")
        Dim resolvedSnapshot As StudentGradeSnapshot =
            If(snapshot, New StudentGradeSnapshot())
        Dim historyItems As List(Of GradeHistoryListItem) =
            BuildHistoryItems(resolvedSnapshot)

        TermsCompletedTextBlock.Text = resolvedSnapshot.TermCount.ToString()
        CumulativeGwaTextBlock.Text = resolvedSnapshot.CurrentGwaLabel
        EarnedUnitsTextBlock.Text = resolvedSnapshot.EarnedUnitsLabel
        GradeHistoryStatusTextBlock.Text =
            BuildStatusMessage(resolvedSnapshot, statusMessage)
        GradeHistoryEmptyTextBlock.Text =
            BuildEmptyStateMessage(resolvedSnapshot, statusMessage)

        GradeHistoryItemsControl.ItemsSource = historyItems
        GradeHistoryItemsControl.Visibility =
            If(historyItems.Count = 0, Visibility.Collapsed, Visibility.Visible)
        GradeHistoryEmptyStateBorder.Visibility =
            If(historyItems.Count = 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Function BuildHistoryItems(snapshot As StudentGradeSnapshot) As List(Of GradeHistoryListItem)
        Dim items As New List(Of GradeHistoryListItem)()
        If snapshot Is Nothing OrElse Not snapshot.HasGrades Then
            Return items
        End If

        items.Add(New GradeHistoryListItem() With {
            .TermLabel = "Current Term",
            .Units = snapshot.EarnedUnitsLabel,
            .Gwa = snapshot.CurrentGwaLabel,
            .Standing = snapshot.AcademicStandingLabel
        })

        Return items
    End Function

    Private Function BuildStatusMessage(snapshot As StudentGradeSnapshot,
                                        statusMessage As String) As String
        Dim normalizedStatusMessage As String = If(statusMessage, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(normalizedStatusMessage) Then
            Return normalizedStatusMessage
        End If

        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "Load your student record to review your grade history."
        End If

        If Not snapshot.HasGrades Then
            Return snapshot.NoticeMessage
        End If

        Return "Your current posted grades are summarized here until term archiving is available."
    End Function

    Private Function BuildEmptyStateMessage(snapshot As StudentGradeSnapshot,
                                            statusMessage As String) As String
        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "No term history loaded."
        End If

        If Not snapshot.HasGrades AndAlso
           Not String.IsNullOrWhiteSpace(snapshot.NoticeMessage) Then
            Return snapshot.NoticeMessage
        End If

        Return BuildStatusMessage(snapshot, statusMessage)
    End Function
End Class
