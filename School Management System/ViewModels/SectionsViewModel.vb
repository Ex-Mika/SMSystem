Imports System.Collections.ObjectModel

Public Class SectionsViewModel
    Inherits ViewModelBase

    Private _selectedSection As Section
    Private _formSectionName As String
    Private _formCourse As String
    Private _formYearLevel As String
    Private _formStatus As String
    Private _searchText As String
    Private _selectedStatusFilter As String
    Private _selectedCourseFilter As String
    Private _selectedYearLevelFilter As String

    Public Sub New()
        Sections = New ObservableCollection(Of Section)()
        StatusOptions = New List(Of String) From {"All", "Active", "Inactive"}
        CourseOptions = New List(Of String) From {"All"}
        YearLevelOptions = New List(Of String) From {"All", "1st Year", "2nd Year", "3rd Year", "4th Year"}
        _selectedStatusFilter = "All"
        _selectedCourseFilter = "All"
        _selectedYearLevelFilter = "All"
        _formStatus = "Active"

        NewCommand = New RelayCommand(Sub(p) ClearForm())
        EditCommand = New RelayCommand(Sub(p) CopySelectedToForm(), Function(p) _selectedSection IsNot Nothing)
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
                                                   SelectedYearLevelFilter = "All"
                                               End Sub)
        RefreshCommand = New RelayCommand(Sub(p)
                                              ' No-op stub — will be wired to data layer later.
                                          End Sub)
    End Sub

    Public Property Sections As ObservableCollection(Of Section)

    Public Property SelectedSection As Section
        Get
            Return _selectedSection
        End Get
        Set(value As Section)
            _selectedSection = value
            OnPropertyChanged()
            CType(EditCommand, RelayCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Public Property FormSectionName As String
        Get
            Return _formSectionName
        End Get
        Set(value As String)
            _formSectionName = value
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

    Public Property FormYearLevel As String
        Get
            Return _formYearLevel
        End Get
        Set(value As String)
            _formYearLevel = value
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

    Public Property SelectedYearLevelFilter As String
        Get
            Return _selectedYearLevelFilter
        End Get
        Set(value As String)
            _selectedYearLevelFilter = value
            OnPropertyChanged()
        End Set
    End Property

    Public ReadOnly Property StatusOptions As List(Of String)
    Public ReadOnly Property CourseOptions As List(Of String)
    Public ReadOnly Property YearLevelOptions As List(Of String)

    Public Property NewCommand As RelayCommand
    Public Property EditCommand As RelayCommand
    Public Property SaveCommand As RelayCommand
    Public Property CancelCommand As RelayCommand
    Public Property DeactivateCommand As RelayCommand
    Public Property ClearFiltersCommand As RelayCommand
    Public Property RefreshCommand As RelayCommand

    Private Sub ClearForm()
        FormSectionName = String.Empty
        FormCourse = String.Empty
        FormYearLevel = String.Empty
        FormStatus = "Active"
    End Sub

    Private Sub CopySelectedToForm()
        If _selectedSection Is Nothing Then Return
        FormSectionName = _selectedSection.SectionName
        FormCourse = _selectedSection.Course
        FormYearLevel = _selectedSection.YearLevel
        FormStatus = _selectedSection.Status
    End Sub
End Class
