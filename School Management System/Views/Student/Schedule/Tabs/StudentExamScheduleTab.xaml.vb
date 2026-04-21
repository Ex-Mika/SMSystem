Imports School_Management_System.Backend.Models

Class StudentExamScheduleTab
    Public Class StudentSubjectActionEventArgs
        Inherits EventArgs

        Public Sub New(subjectCode As String)
            Me.SubjectCode = If(subjectCode, String.Empty).Trim()
        End Sub

        Public ReadOnly Property SubjectCode As String
    End Class

    Public Event EnrollRequested As EventHandler(Of StudentSubjectActionEventArgs)
    Public Event RemoveRequested As EventHandler(Of StudentSubjectActionEventArgs)

    Public Sub New()
        InitializeComponent()
        SetEnrollmentSnapshot(Nothing)
    End Sub

    Public Sub SetEnrollmentSnapshot(snapshot As StudentEnrollmentSnapshot,
                                     Optional statusMessage As String = "")
        Dim resolvedSnapshot As StudentEnrollmentSnapshot =
            If(snapshot, New StudentEnrollmentSnapshot())
        Dim resolvedStatusMessage As String = If(statusMessage, String.Empty).Trim()

        AvailableSubjectsCountTextBlock.Text = resolvedSnapshot.AvailableSubjectCount.ToString()
        SelectedSubjectsCountTextBlock.Text = resolvedSnapshot.SelectedSubjectCount.ToString()
        SelectedSubjectsUnitsTextBlock.Text = resolvedSnapshot.SelectedTotalUnitsLabel

        If String.IsNullOrWhiteSpace(resolvedStatusMessage) Then
            resolvedStatusMessage = resolvedSnapshot.NoticeMessage
        End If

        EnrollmentStatusTextBlock.Text = If(String.IsNullOrWhiteSpace(resolvedStatusMessage),
                                            "Choose the subjects you want to include in your current load.",
                                            resolvedStatusMessage)

        SelectedSubjectsStatusTextBlock.Text =
            BuildSelectedSubjectsStatus(resolvedSnapshot)

        AvailableSubjectsItemsControl.ItemsSource = resolvedSnapshot.AvailableSubjects
        SelectedSubjectsItemsControl.ItemsSource = resolvedSnapshot.SelectedSubjects

        AvailableSubjectsEmptyTextBlock.Visibility =
            If(resolvedSnapshot.AvailableSubjectCount = 0,
               Visibility.Visible,
               Visibility.Collapsed)
        SelectedSubjectsEmptyTextBlock.Visibility =
            If(resolvedSnapshot.SelectedSubjectCount = 0,
               Visibility.Visible,
               Visibility.Collapsed)
    End Sub

    Private Sub AvailableSubjectEnrollButton_Click(sender As Object, e As RoutedEventArgs)
        Dim subjectCode As String = ReadSubjectCodeFromSender(sender)
        If String.IsNullOrWhiteSpace(subjectCode) Then
            Return
        End If

        RaiseEvent EnrollRequested(Me, New StudentSubjectActionEventArgs(subjectCode))
    End Sub

    Private Sub SelectedSubjectRemoveButton_Click(sender As Object, e As RoutedEventArgs)
        Dim subjectCode As String = ReadSubjectCodeFromSender(sender)
        If String.IsNullOrWhiteSpace(subjectCode) Then
            Return
        End If

        RaiseEvent RemoveRequested(Me, New StudentSubjectActionEventArgs(subjectCode))
    End Sub

    Private Function ReadSubjectCodeFromSender(sender As Object) As String
        Dim button As Button = TryCast(sender, Button)
        If button Is Nothing OrElse button.Tag Is Nothing Then
            Return String.Empty
        End If

        Return button.Tag.ToString().Trim()
    End Function

    Private Function BuildSelectedSubjectsStatus(snapshot As StudentEnrollmentSnapshot) As String
        If snapshot Is Nothing OrElse snapshot.SelectedSubjectCount = 0 Then
            Return "No subjects have been added to your current load yet."
        End If

        If snapshot.HasRemainingSubjects Then
            Return snapshot.SelectedSubjectCount.ToString() &
                " subjects saved in your current load. " &
                snapshot.AvailableSubjectCount.ToString() &
                " still need to be added."
        End If

        Return snapshot.SelectedSubjectCount.ToString() &
            " subjects saved in your finalized load."
    End Function
End Class
