Imports System.Collections.ObjectModel

Public Class CoursesViewModel
    Inherits ViewModelBase

    Private _selectedCourse As Course
    Private _formCourseCode As String
    Private _formCourseName As String
    Private _formStatus As String
    Private _searchText As String
    Private _selectedStatusFilter As String

    Public Sub New()
        Courses = New ObservableCollection(Of Course)()
        StatusOptions = New List(Of String) From {"All", "Active", "Inactive"}
        _selectedStatusFilter = "All"
        _formStatus = "Active"

        NewCommand = New RelayCommand(Sub(p) ClearForm())
        EditCommand = New RelayCommand(Sub(p) CopySelectedToForm(), Function(p) _selectedCourse IsNot Nothing)
        SaveCommand = New RelayCommand(Sub(p)
                                           ' No-op stub — will be wired to data layer later.
                                       End Sub)
        CancelCommand = New RelayCommand(Sub(p) ClearForm())
        DeactivateCommand = New RelayCommand(Sub(p)
                                                 ' No-op stub — will be wired to data layer later.
                                             End Sub)
        ClearFiltersCommand = New RelayCommand(Sub(p)
                                                   SearchText = String.Empty
                                                   SelectedStatusFilter = "All"
                                               End Sub)
        RefreshCommand = New RelayCommand(Sub(p)
                                              ' No-op stub — will be wired to data layer later.
                                          End Sub)
    End Sub

    Public Property Courses As ObservableCollection(Of Course)

    Public Property SelectedCourse As Course
        Get
            Return _selectedCourse
        End Get
        Set(value As Course)
            _selectedCourse = value
            OnPropertyChanged()
            CType(EditCommand, RelayCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Public Property FormCourseCode As String
        Get
            Return _formCourseCode
        End Get
        Set(value As String)
            _formCourseCode = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormCourseName As String
        Get
            Return _formCourseName
        End Get
        Set(value As String)
            _formCourseName = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormStatus As String
        Get
            Return _formStatus
        End Get
        Set(value As String)
            _formStatus = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property SearchText As String
        Get
            Return _searchText
        End Get
        Set(value As String)
            _searchText = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property SelectedStatusFilter As String
        Get
            Return _selectedStatusFilter
        End Get
        Set(value As String)
            _selectedStatusFilter = value
            OnPropertyChanged()
        End Set
    End Property

    Public ReadOnly Property StatusOptions As List(Of String)

    Public Property NewCommand As RelayCommand
    Public Property EditCommand As RelayCommand
    Public Property SaveCommand As RelayCommand
    Public Property CancelCommand As RelayCommand
    Public Property DeactivateCommand As RelayCommand
    Public Property ClearFiltersCommand As RelayCommand
    Public Property RefreshCommand As RelayCommand

    Private Sub ClearForm()
        FormCourseCode = String.Empty
        FormCourseName = String.Empty
        FormStatus = "Active"
    End Sub

    Private Sub CopySelectedToForm()
        If _selectedCourse Is Nothing Then Return
        FormCourseCode = _selectedCourse.CourseCode
        FormCourseName = _selectedCourse.CourseName
        FormStatus = _selectedCourse.Status
    End Sub
End Class
