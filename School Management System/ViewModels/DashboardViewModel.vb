Imports System.Collections.ObjectModel

Public Class DashboardViewModel
    Inherits ViewModelBase

    Public Sub New()
        RecentActivities = New ObservableCollection(Of String)()

        AddStudentCommand = New RelayCommand(Sub(p)
                                                 ' No-op stub — will navigate to Students page later.
                                             End Sub)
        AddTeacherCommand = New RelayCommand(Sub(p)
                                                 ' No-op stub — will navigate to Teachers page later.
                                             End Sub)
        CreateCourseCommand = New RelayCommand(Sub(p)
                                                   ' No-op stub — will navigate to Courses page later.
                                               End Sub)
        CreateSubjectCommand = New RelayCommand(Sub(p)
                                                    ' No-op stub — will navigate to Subjects page later.
                                                End Sub)
        OpenSchedulingCommand = New RelayCommand(Sub(p)
                                                     ' No-op stub — will navigate to Scheduling page later.
                                                 End Sub)
    End Sub

    ' KPI placeholder values — will be replaced with live data from backend later.
    Public ReadOnly Property TotalStudents As String = "—"
    Public ReadOnly Property TotalTeachers As String = "—"
    Public ReadOnly Property TotalCourses As String = "—"
    Public ReadOnly Property TotalSubjects As String = "—"
    Public ReadOnly Property TotalSections As String = "—"
    Public ReadOnly Property ScheduledClassesToday As String = "—"

    Public Property RecentActivities As ObservableCollection(Of String)

    Public Property AddStudentCommand As RelayCommand
    Public Property AddTeacherCommand As RelayCommand
    Public Property CreateCourseCommand As RelayCommand
    Public Property CreateSubjectCommand As RelayCommand
    Public Property OpenSchedulingCommand As RelayCommand
End Class
