Imports System.Collections.Generic
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models

Namespace Backend.Repositories
    Public Class AdminOverviewRepository
        Private Const MetricsSql As String =
            "SELECT " &
            "(SELECT COUNT(*) FROM students) AS total_students, " &
            "(SELECT COUNT(*) FROM teachers) AS total_teachers, " &
            "(SELECT COUNT(*) FROM users " &
            " WHERE created_at >= DATE_SUB(NOW(), INTERVAL 30 DAY)) AS new_applications;"

        Private Const RecentActivitySql As String =
            "SELECT role_key, username, created_at " &
            "FROM users " &
            "ORDER BY created_at DESC, user_id DESC " &
            "LIMIT 5;"

        Private Const DepartmentPerformanceSql As String =
            "SELECT " &
            "d.department_name, " &
            "d.department_code, " &
            "(SELECT COUNT(*) FROM teachers t " &
            " WHERE t.department_id = d.department_id) AS teacher_count, " &
            "(SELECT COUNT(*) FROM courses c " &
            " WHERE c.department_id = d.department_id) AS course_count, " &
            "(SELECT COUNT(*) " &
            " FROM students s " &
            " INNER JOIN courses c ON c.course_id = s.course_id " &
            " WHERE c.department_id = d.department_id) AS student_count " &
            "FROM departments d " &
            "ORDER BY student_count DESC, teacher_count DESC, d.department_name ASC " &
            "LIMIT 5;"

        Public Function GetDashboardOverview() As AdminDashboardOverview
            Dim overview As New AdminDashboardOverview()

            Using connection As MySqlConnection = Database.DatabaseModule.OpenConnection()
                LoadMetrics(connection, overview)
                overview.RecentActivities = LoadRecentActivities(connection)
                overview.DepartmentPerformance = LoadDepartmentPerformance(connection)
            End Using

            Return overview
        End Function

        Private Sub LoadMetrics(connection As MySqlConnection,
                                overview As AdminDashboardOverview)
            Using command As New MySqlCommand(MetricsSql, connection)
                Using reader As MySqlDataReader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Return
                    End If

                    overview.TotalStudents = Convert.ToInt32(reader("total_students"))
                    overview.TotalTeachers = Convert.ToInt32(reader("total_teachers"))
                    overview.NewApplications = Convert.ToInt32(reader("new_applications"))
                End Using
            End Using
        End Sub

        Private Function LoadRecentActivities(connection As MySqlConnection) As List(Of AdminRecentActivityItem)
            Dim items As New List(Of AdminRecentActivityItem)()

            Using command As New MySqlCommand(RecentActivitySql, connection)
                Using reader As MySqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim roleKey As String = Convert.ToString(reader("role_key"))
                        Dim username As String = Convert.ToString(reader("username"))
                        Dim createdAt As DateTime = Convert.ToDateTime(reader("created_at"))

                        items.Add(New AdminRecentActivityItem() With {
                            .Title = BuildActivityTitle(roleKey, username),
                            .TimestampText = createdAt.ToString("MMM dd, yyyy hh:mm tt")
                        })
                    End While
                End Using
            End Using

            Return items
        End Function

        Private Function LoadDepartmentPerformance(connection As MySqlConnection) As List(Of AdminDepartmentPerformanceItem)
            Dim items As New List(Of AdminDepartmentPerformanceItem)()

            Using command As New MySqlCommand(DepartmentPerformanceSql, connection)
                Using reader As MySqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim departmentName As String = Convert.ToString(reader("department_name"))
                        Dim departmentCode As String = Convert.ToString(reader("department_code"))
                        Dim teacherCount As Integer = Convert.ToInt32(reader("teacher_count"))
                        Dim courseCount As Integer = Convert.ToInt32(reader("course_count"))
                        Dim studentCount As Integer = Convert.ToInt32(reader("student_count"))

                        items.Add(New AdminDepartmentPerformanceItem() With {
                            .Title = $"{departmentName} ({departmentCode})",
                            .Summary = $"{studentCount} students - {teacherCount} teachers - {courseCount} courses"
                        })
                    End While
                End Using
            End Using

            Return items
        End Function

        Private Function BuildActivityTitle(roleKey As String, username As String) As String
            Dim normalizedRole As String = If(roleKey, String.Empty).Trim().ToLowerInvariant()
            Dim displayName As String = If(username, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(displayName) Then
                displayName = "Unnamed Account"
            End If

            Select Case normalizedRole
                Case "student"
                    Return $"Student account created for {displayName}"
                Case "teacher"
                    Return $"Teacher account created for {displayName}"
                Case Else
                    Return $"Administrator account created for {displayName}"
            End Select
        End Function
    End Class
End Namespace
