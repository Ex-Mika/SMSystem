Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class AdminOverviewService
        Private ReadOnly _repository As AdminOverviewRepository

        Public Sub New()
            Me.New(New AdminOverviewRepository())
        End Sub

        Public Sub New(repository As AdminOverviewRepository)
            _repository = repository
        End Sub

        Public Function GetDashboardOverview() As AdminDashboardOverview
            Dim overview As AdminDashboardOverview = _repository.GetDashboardOverview()
            If overview Is Nothing Then
                Return New AdminDashboardOverview()
            End If

            If overview.RecentActivities Is Nothing Then
                overview.RecentActivities = New List(Of AdminRecentActivityItem)()
            End If

            If overview.DepartmentPerformance Is Nothing Then
                overview.DepartmentPerformance = New List(Of AdminDepartmentPerformanceItem)()
            End If

            If String.IsNullOrWhiteSpace(overview.FeesCollectedDisplay) Then
                overview.FeesCollectedDisplay = "N/A"
            End If

            If String.IsNullOrWhiteSpace(overview.FeesSummary) Then
                overview.FeesSummary = "Finance module not connected yet."
            End If

            Return overview
        End Function
    End Class
End Namespace
