Imports System.Collections.Generic
Imports School_Management_System.Backend.Models

Class StudentCurrentGradesTab
    Private NotInheritable Class CurrentGradeListItem
        Public Property SubjectName As String = String.Empty
        Public Property SubjectCode As String = String.Empty
        Public Property Units As String = String.Empty
        Public Property FinalGrade As String = String.Empty
        Public Property Remarks As String = String.Empty
    End Class

    Public Sub New()
        InitializeComponent()
        SetGradeSnapshot(Nothing)
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        CurrentGradesStudentIdTextBlock.Text = If(studentId, String.Empty).Trim()
    End Sub

    Public Sub SetGradeSnapshot(snapshot As StudentGradeSnapshot,
                                Optional statusMessage As String = "")
        Dim resolvedSnapshot As StudentGradeSnapshot =
            If(snapshot, New StudentGradeSnapshot())
        Dim currentGradeItems As List(Of CurrentGradeListItem) =
            BuildCurrentGradeItems(resolvedSnapshot)

        SubjectsPostedTextBlock.Text = resolvedSnapshot.PostedGradeCount.ToString()
        UnitsRatedTextBlock.Text = resolvedSnapshot.RatedUnitsLabel
        CurrentGradesStatusTextBlock.Text =
            BuildStatusMessage(resolvedSnapshot, statusMessage)
        CurrentGradesEmptyTextBlock.Text =
            BuildEmptyStateMessage(resolvedSnapshot, statusMessage)

        CurrentGradesItemsControl.ItemsSource = currentGradeItems
        CurrentGradesItemsControl.Visibility =
            If(currentGradeItems.Count = 0, Visibility.Collapsed, Visibility.Visible)
        CurrentGradesEmptyStateBorder.Visibility =
            If(currentGradeItems.Count = 0, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Function BuildCurrentGradeItems(snapshot As StudentGradeSnapshot) As List(Of CurrentGradeListItem)
        Dim items As New List(Of CurrentGradeListItem)()
        If snapshot Is Nothing OrElse snapshot.Grades Is Nothing Then
            Return items
        End If

        For Each grade As StudentSubjectGradeRecord In snapshot.Grades
            If grade Is Nothing OrElse Not grade.HasPublishedGrade Then
                Continue For
            End If

            items.Add(New CurrentGradeListItem() With {
                .SubjectName = ResolveSubjectName(grade),
                .SubjectCode = ResolveSubjectCode(grade),
                .Units = ResolveUnits(grade),
                .FinalGrade = grade.FinalGradeLabel,
                .Remarks = grade.RemarksLabel
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
            Return "Load your student record to see posted grades."
        End If

        If Not snapshot.HasGrades Then
            Return snapshot.NoticeMessage
        End If

        Return snapshot.PostedGradeCount.ToString() &
            " posted grade(s) are available for your current subjects."
    End Function

    Private Function BuildEmptyStateMessage(snapshot As StudentGradeSnapshot,
                                            statusMessage As String) As String
        If snapshot Is Nothing OrElse snapshot.Student Is Nothing Then
            Return "No grades loaded yet."
        End If

        If Not snapshot.HasGrades AndAlso
           Not String.IsNullOrWhiteSpace(snapshot.NoticeMessage) Then
            Return snapshot.NoticeMessage
        End If

        Return BuildStatusMessage(snapshot, statusMessage)
    End Function

    Private Function ResolveSubjectName(grade As StudentSubjectGradeRecord) As String
        If grade Is Nothing Then
            Return "Unnamed Subject"
        End If

        Dim subjectName As String = If(grade.SubjectName, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(subjectName) Then
            Return subjectName
        End If

        Return "Unnamed Subject"
    End Function

    Private Function ResolveSubjectCode(grade As StudentSubjectGradeRecord) As String
        If grade Is Nothing Then
            Return "--"
        End If

        Dim subjectCode As String = If(grade.SubjectCode, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(subjectCode) Then
            Return subjectCode
        End If

        Return "--"
    End Function

    Private Function ResolveUnits(grade As StudentSubjectGradeRecord) As String
        If grade Is Nothing Then
            Return "0"
        End If

        Dim unitsValue As String = If(grade.SubjectUnits, String.Empty).Trim()
        If Not String.IsNullOrWhiteSpace(unitsValue) Then
            Return unitsValue
        End If

        Return "0"
    End Function
End Class
