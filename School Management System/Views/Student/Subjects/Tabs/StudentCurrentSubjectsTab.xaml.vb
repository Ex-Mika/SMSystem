Imports System.Collections.Generic
Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models

Class StudentCurrentSubjectsTab
    Private NotInheritable Class CurrentSubjectListItem
        Public Property SubjectName As String = String.Empty
        Public Property SubjectCode As String = String.Empty
        Public Property Units As String = String.Empty
        Public Property StatusText As String = String.Empty
    End Class

    Public Sub New()
        InitializeComponent()
        SetEnrollmentSnapshot(Nothing)
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        CurrentStudentIdTextBlock.Text = If(studentId, String.Empty).Trim()
    End Sub

    Public Sub SetEnrollmentSnapshot(snapshot As StudentEnrollmentSnapshot,
                                     Optional statusMessage As String = "")
        Dim resolvedSnapshot As StudentEnrollmentSnapshot =
            If(snapshot, New StudentEnrollmentSnapshot())
        Dim currentSubjectItems As List(Of CurrentSubjectListItem) =
            BuildCurrentSubjectItems(resolvedSnapshot)

        CurrentTotalUnitsTextBlock.Text = resolvedSnapshot.SelectedTotalUnitsLabel
        CurrentSectionTextBlock.Text = BuildSectionLabel(resolvedSnapshot.Student)
        CurrentSubjectsStatusTextBlock.Text =
            BuildStatusMessage(resolvedSnapshot, statusMessage)
        CurrentSubjectsEmptyTextBlock.Text =
            BuildEmptyStateMessage(resolvedSnapshot, statusMessage)

        CurrentSubjectsItemsControl.ItemsSource = currentSubjectItems
        CurrentSubjectsItemsControl.Visibility =
            If(currentSubjectItems.Count = 0, Visibility.Collapsed, Visibility.Visible)
        CurrentSubjectsEmptyStateBorder.Visibility =
            If(currentSubjectItems.Count = 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Function BuildCurrentSubjectItems(snapshot As StudentEnrollmentSnapshot) As List(Of CurrentSubjectListItem)
        Dim items As New List(Of CurrentSubjectListItem)()
        If snapshot Is Nothing OrElse snapshot.SelectedSubjects Is Nothing Then
            Return items
        End If

        Dim subjectStatusText As String = BuildSubjectStatusText(snapshot)

        For Each subject As SubjectRecord In snapshot.SelectedSubjects
            If subject Is Nothing Then
                Continue For
            End If

            items.Add(New CurrentSubjectListItem() With {
                .SubjectName = ResolveSubjectName(subject),
                .SubjectCode = ResolveSubjectCode(subject),
                .Units = ResolveUnits(subject),
                .StatusText = subjectStatusText
            })
        Next

        Return items
    End Function

    Private Function BuildStatusMessage(snapshot As StudentEnrollmentSnapshot,
                                        statusMessage As String) As String
        Dim normalizedStatusMessage As String = If(statusMessage, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(normalizedStatusMessage) Then
            Return normalizedStatusMessage
        End If

        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "Load your student record to see your active subjects."
        End If

        If Not snapshot.HasSelectedSubjects Then
            Return "Enroll subjects in Class Schedule > Enrollment to populate this tab."
        End If

        If Not snapshot.HasAssignedSection Then
            Return "Your selected subjects are saved. Choose a section in Class Schedule > Submission to finalize them."
        End If

        If snapshot.HasRemainingSubjects Then
            Return "Your saved subjects are listed below. Add the remaining eligible subjects to finish enrollment."
        End If

        Return "Your enrolled subjects are now active under " &
            BuildSectionLabel(snapshot.Student) & "."
    End Function

    Private Function BuildEmptyStateMessage(snapshot As StudentEnrollmentSnapshot,
                                            statusMessage As String) As String
        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "No subjects loaded yet."
        End If

        If Not snapshot.HasSelectedSubjects AndAlso snapshot.AvailableSubjectCount > 0 Then
            Return "No subjects are in your current load yet. Add them from Class Schedule > Enrollment."
        End If

        If Not snapshot.HasSelectedSubjects AndAlso
           Not String.IsNullOrWhiteSpace(snapshot.NoticeMessage) Then
            Return snapshot.NoticeMessage
        End If

        Return BuildStatusMessage(snapshot, statusMessage)
    End Function

    Private Function BuildSubjectStatusText(snapshot As StudentEnrollmentSnapshot) As String
        If snapshot Is Nothing OrElse Not snapshot.HasAssignedSection Then
            Return "Section Needed"
        End If

        If snapshot.HasRemainingSubjects Then
            Return "Load Saved"
        End If

        Return "Enrolled"
    End Function

    Private Function BuildSectionLabel(student As StudentRecord) As String
        Return StudentScheduleHelper.BuildStudentSectionValue(student,
                                                              "Not Selected")
    End Function

    Private Function ResolveSubjectName(subject As SubjectRecord) As String
        Dim subjectName As String = If(subject.SubjectName, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(subjectName) Then
            Return subjectName
        End If

        Return "Unnamed Subject"
    End Function

    Private Function ResolveSubjectCode(subject As SubjectRecord) As String
        Dim subjectCode As String = If(subject.SubjectCode, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(subjectCode) Then
            Return subjectCode
        End If

        Return "--"
    End Function

    Private Function ResolveUnits(subject As SubjectRecord) As String
        Dim units As String = If(subject.Units, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(units) Then
            Return units
        End If

        Return "0"
    End Function
End Class
