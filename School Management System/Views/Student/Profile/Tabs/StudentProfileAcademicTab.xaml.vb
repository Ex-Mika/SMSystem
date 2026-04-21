Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class StudentProfileAcademicTab
    Private ReadOnly _studentManagementService As New StudentManagementService()

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetStudentContext(studentId As String, studentName As String)
        Dim normalizedId As String = If(studentId, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedId) Then
            ApplyStudentRecord(Nothing)
            Return
        End If

        Dim result = _studentManagementService.GetStudentByStudentNumber(normalizedId)
        If Not result.IsSuccess Then
            ApplyStudentRecord(Nothing)
            Return
        End If

        ApplyStudentRecord(result.Data)
    End Sub

    Public Sub SetStudentRecord(student As StudentRecord)
        ApplyStudentRecord(student)
    End Sub

    Private Sub ApplyStudentRecord(student As StudentRecord)
        ProgramValueTextBlock.Text = ResolveDisplayValue(If(student Is Nothing,
                                                            String.Empty,
                                                            student.CourseDisplayName))
        YearLevelValueTextBlock.Text = ResolveDisplayValue(If(student Is Nothing,
                                                              String.Empty,
                                                              student.YearLevelLabel))
        SectionValueTextBlock.Text = ResolveDisplayValue(BuildSectionDisplayValue(student))
    End Sub

    Private Function BuildSectionDisplayValue(student As StudentRecord) As String
        Return StudentScheduleHelper.BuildStudentSectionValue(student, String.Empty)
    End Function

    Private Function ResolveDisplayValue(value As String) As String
        Dim normalizedValue As String = If(value, String.Empty).Trim()
        If String.IsNullOrWhiteSpace(normalizedValue) Then
            Return "Not set"
        End If

        Return normalizedValue
    End Function
End Class
