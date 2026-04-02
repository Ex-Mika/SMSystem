Imports System.Threading.Tasks
Imports MySql.Data.MySqlClient
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Services

Class AdminOverviewView
    Private ReadOnly _overviewService As New AdminOverviewService()
    Private _hasLoaded As Boolean
    Private _isLoading As Boolean

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Async Sub RefreshData()
        _hasLoaded = True
        Await LoadOverviewAsync()
    End Sub

    Private Async Sub AdminOverviewView_Loaded(sender As Object, e As RoutedEventArgs)
        If _hasLoaded Then
            Return
        End If

        _hasLoaded = True
        Await LoadOverviewAsync()
    End Sub

    Private Async Function LoadOverviewAsync() As Task
        If _isLoading Then
            Return
        End If

        _isLoading = True
        SetLoadingState()

        Try
            Dim overview As AdminDashboardOverview =
                Await Task.Run(Function() _overviewService.GetDashboardOverview())
            ApplyOverview(overview)
        Catch ex As MySqlException
            ApplyErrorState("Unable to connect to MySQL on 127.0.0.1.")
        Catch ex As InvalidOperationException
            ApplyErrorState("Unable to load admin dashboard data.")
        Catch ex As Exception
            ApplyErrorState("Unable to load admin dashboard data.")
        Finally
            _isLoading = False
        End Try
    End Function

    Private Sub SetLoadingState()
        TotalStudentsValueTextBlock.Text = "--"
        TotalTeachersValueTextBlock.Text = "--"
        NewApplicationsValueTextBlock.Text = "--"

        TotalStudentsStatusTextBlock.Text = "Loading student data..."
        TotalTeachersStatusTextBlock.Text = "Loading teacher data..."
        NewApplicationsStatusTextBlock.Text = "Loading recent account activity..."

        RecentActivityItemsControl.ItemsSource = Nothing
        RecentActivityEmptyTextBlock.Text = "Loading recent activity..."
        RecentActivityEmptyTextBlock.Visibility = Visibility.Visible

        DepartmentPerformanceItemsControl.ItemsSource = Nothing
        DepartmentPerformanceEmptyTextBlock.Text = "Loading department data..."
        DepartmentPerformanceEmptyTextBlock.Visibility = Visibility.Visible
    End Sub

    Private Sub ApplyOverview(overview As AdminDashboardOverview)
        Dim resolvedOverview As AdminDashboardOverview = If(overview, New AdminDashboardOverview())

        TotalStudentsValueTextBlock.Text = resolvedOverview.TotalStudents.ToString()
        TotalTeachersValueTextBlock.Text = resolvedOverview.TotalTeachers.ToString()
        NewApplicationsValueTextBlock.Text = resolvedOverview.NewApplications.ToString()

        TotalStudentsStatusTextBlock.Text = BuildCountSummary(resolvedOverview.TotalStudents,
                                                              "registered student",
                                                              "registered students")
        TotalTeachersStatusTextBlock.Text = BuildCountSummary(resolvedOverview.TotalTeachers,
                                                              "active teacher profile",
                                                              "active teacher profiles")
        NewApplicationsStatusTextBlock.Text = "User accounts created in the last 30 days"

        RecentActivityItemsControl.ItemsSource = resolvedOverview.RecentActivities
        RecentActivityEmptyTextBlock.Visibility = If(resolvedOverview.RecentActivities.Count = 0,
                                                     Visibility.Visible,
                                                     Visibility.Collapsed)
        RecentActivityEmptyTextBlock.Text = If(resolvedOverview.RecentActivities.Count = 0,
                                               "No recent activity found.",
                                               String.Empty)

        DepartmentPerformanceItemsControl.ItemsSource = resolvedOverview.DepartmentPerformance
        DepartmentPerformanceEmptyTextBlock.Visibility = If(resolvedOverview.DepartmentPerformance.Count = 0,
                                                            Visibility.Visible,
                                                            Visibility.Collapsed)
        DepartmentPerformanceEmptyTextBlock.Text = If(resolvedOverview.DepartmentPerformance.Count = 0,
                                                      "No departments found.",
                                                      String.Empty)
    End Sub

    Private Sub ApplyErrorState(message As String)
        TotalStudentsValueTextBlock.Text = "--"
        TotalTeachersValueTextBlock.Text = "--"
        NewApplicationsValueTextBlock.Text = "--"

        TotalStudentsStatusTextBlock.Text = message
        TotalTeachersStatusTextBlock.Text = message
        NewApplicationsStatusTextBlock.Text = message

        RecentActivityItemsControl.ItemsSource = Nothing
        RecentActivityEmptyTextBlock.Text = message
        RecentActivityEmptyTextBlock.Visibility = Visibility.Visible

        DepartmentPerformanceItemsControl.ItemsSource = Nothing
        DepartmentPerformanceEmptyTextBlock.Text = message
        DepartmentPerformanceEmptyTextBlock.Visibility = Visibility.Visible
    End Sub

    Private Function BuildCountSummary(count As Integer,
                                       singularLabel As String,
                                       pluralLabel As String) As String
        If count = 1 Then
            Return $"1 {singularLabel}"
        End If

        Return $"{count} {pluralLabel}"
    End Function
End Class
