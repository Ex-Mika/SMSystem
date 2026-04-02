Imports System.Collections.Generic

Namespace Backend.Models
    Public Class AdminDashboardOverview
        Public Property TotalStudents As Integer
        Public Property TotalTeachers As Integer
        Public Property NewApplications As Integer
        Public Property FeesCollectedDisplay As String = "N/A"
        Public Property FeesSummary As String = "Finance module not connected yet."
        Public Property RecentActivities As List(Of AdminRecentActivityItem) =
            New List(Of AdminRecentActivityItem)()
        Public Property DepartmentPerformance As List(Of AdminDepartmentPerformanceItem) =
            New List(Of AdminDepartmentPerformanceItem)()
    End Class
End Namespace
