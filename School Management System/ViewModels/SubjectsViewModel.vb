Imports System.Collections.ObjectModel

Public Class SubjectsViewModel
    Inherits ViewModelBase

    Private _selectedSubject As Subject
    Private _formSubjectCode As String
    Private _formDescription As String
    Private _formUnits As String
    Private _formCourse As String
    Private _formStatus As String
    Private _searchText As String
    Private _selectedStatusFilter As String
    Private _selectedCourseFilter As String

    Public Sub New()
        Subjects = New ObservableCollection(Of Subject)()
        StatusOptions = New List(Of String) From {"All", "Active", "Inactive"}
        CourseOptions = New List(Of String) From {"All"}
        _selectedStatusFilter = "All"
        _selectedCourseFilter = "All"
        _formStatus = "Active"

        NewCommand = New RelayCommand(Sub(p) ClearForm())
        EditCommand = New RelayCommand(Sub(p) CopySelectedToForm(), Function(p) _selectedSubject IsNot Nothing)
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
                                                   SelectedCourseFilter = "All"
                                               End Sub)
        RefreshCommand = New RelayCommand(Sub(p)
                                              ' No-op stub — will be wired to data layer later.
                                          End Sub)
    End Sub

    Public Property Subjects As ObservableCollection(Of Subject)

    Public Property SelectedSubject As Subject
        Get
            Return _selectedSubject
        End Get
        Set(value As Subject)
            _selectedSubject = value
            OnPropertyChanged()
            CType(EditCommand, RelayCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Public Property FormSubjectCode As String
        Get
            Return _formSubjectCode
        End Get
        Set(value As String)
            _formSubjectCode = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormDescription As String
        Get
            Return _formDescription
        End Get
        Set(value As String)
            _formDescription = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormUnits As String
        Get
            Return _formUnits
        End Get
        Set(value As String)
            _formUnits = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormCourse As String
        Get
            Return _formCourse
        End Get
        Set(value As String)
            _formCourse = value
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

    Public Property SelectedCourseFilter As String
        Get
            Return _selectedCourseFilter
        End Get
        Set(value As String)
            _selectedCourseFilter = value
            OnPropertyChanged()
        End Set
    End Property

    Public ReadOnly Property StatusOptions As List(Of String)
    Public ReadOnly Property CourseOptions As List(Of String)

    Public Property NewCommand As RelayCommand
    Public Property EditCommand As RelayCommand
    Public Property SaveCommand As RelayCommand
    Public Property CancelCommand As RelayCommand
    Public Property DeactivateCommand As RelayCommand
    Public Property ClearFiltersCommand As RelayCommand
    Public Property RefreshCommand As RelayCommand

    Private Sub ClearForm()
        FormSubjectCode = String.Empty
        FormDescription = String.Empty
        FormUnits = String.Empty
        FormCourse = String.Empty
        FormStatus = "Active"
    End Sub

    Private Sub CopySelectedToForm()
        If _selectedSubject Is Nothing Then Return
        FormSubjectCode = _selectedSubject.SubjectCode
        FormDescription = _selectedSubject.Description
        FormUnits = _selectedSubject.Units.ToString()
        FormCourse = _selectedSubject.Course
        FormStatus = _selectedSubject.Status
    End Sub
End Class
