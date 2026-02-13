Imports System.Collections.ObjectModel

Public Class TeachersViewModel
    Inherits ViewModelBase

    Private _selectedTeacher As Teacher
    Private _formTeacherId As String
    Private _formFullName As String
    Private _formEmail As String
    Private _formContactNo As String
    Private _formDepartment As String
    Private _formStatus As String
    Private _searchText As String
    Private _selectedStatusFilter As String
    Private _selectedCourseFilter As String
    Private _selectedDayFilter As String
    Private _selectedSectionFilter As String
    Private _selectedRoomFilter As String

    Public Sub New()
        Teachers = New ObservableCollection(Of Teacher)()
        AssignedClasses = New ObservableCollection(Of ClassSchedule)()
        StatusOptions = New List(Of String) From {"All", "Active", "Inactive"}
        CourseOptions = New List(Of String) From {"All"}
        DayOptions = New List(Of String) From {"All", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
        SectionOptions = New List(Of String) From {"All"}
        RoomOptions = New List(Of String) From {"All"}
        _selectedStatusFilter = "All"
        _selectedCourseFilter = "All"
        _selectedDayFilter = "All"
        _selectedSectionFilter = "All"
        _selectedRoomFilter = "All"
        _formStatus = "Active"

        NewCommand = New RelayCommand(Sub(p) ClearForm())
        EditCommand = New RelayCommand(Sub(p) CopySelectedToForm(), Function(p) _selectedTeacher IsNot Nothing)
        SaveCommand = New RelayCommand(Sub(p)
                                           ' No-op stub — will be wired to data layer later.
                                       End Sub)
        CancelCommand = New RelayCommand(Sub(p) ClearForm())
        BlockCommand = New RelayCommand(Sub(p)
                                            ' No-op stub — will be wired to data layer later.
                                        End Sub)
        ClearFiltersCommand = New RelayCommand(Sub(p)
                                                   SearchText = String.Empty
                                                   SelectedStatusFilter = "All"
                                                   SelectedCourseFilter = "All"
                                                   SelectedDayFilter = "All"
                                                   SelectedSectionFilter = "All"
                                                   SelectedRoomFilter = "All"
                                               End Sub)
        RefreshCommand = New RelayCommand(Sub(p)
                                              ' No-op stub — will be wired to data layer later.
                                          End Sub)
    End Sub

    Public Property Teachers As ObservableCollection(Of Teacher)
    Public Property AssignedClasses As ObservableCollection(Of ClassSchedule)

    Public Property SelectedTeacher As Teacher
        Get
            Return _selectedTeacher
        End Get
        Set(value As Teacher)
            _selectedTeacher = value
            OnPropertyChanged()
            CType(EditCommand, RelayCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Public Property FormTeacherId As String
        Get
            Return _formTeacherId
        End Get
        Set(value As String)
            _formTeacherId = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormFullName As String
        Get
            Return _formFullName
        End Get
        Set(value As String)
            _formFullName = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormEmail As String
        Get
            Return _formEmail
        End Get
        Set(value As String)
            _formEmail = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormContactNo As String
        Get
            Return _formContactNo
        End Get
        Set(value As String)
            _formContactNo = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormDepartment As String
        Get
            Return _formDepartment
        End Get
        Set(value As String)
            _formDepartment = value
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

    Public Property SelectedDayFilter As String
        Get
            Return _selectedDayFilter
        End Get
        Set(value As String)
            _selectedDayFilter = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property SelectedSectionFilter As String
        Get
            Return _selectedSectionFilter
        End Get
        Set(value As String)
            _selectedSectionFilter = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property SelectedRoomFilter As String
        Get
            Return _selectedRoomFilter
        End Get
        Set(value As String)
            _selectedRoomFilter = value
            OnPropertyChanged()
        End Set
    End Property

    Public ReadOnly Property StatusOptions As List(Of String)
    Public ReadOnly Property CourseOptions As List(Of String)
    Public ReadOnly Property DayOptions As List(Of String)
    Public ReadOnly Property SectionOptions As List(Of String)
    Public ReadOnly Property RoomOptions As List(Of String)

    Public Property NewCommand As RelayCommand
    Public Property EditCommand As RelayCommand
    Public Property SaveCommand As RelayCommand
    Public Property CancelCommand As RelayCommand
    Public Property BlockCommand As RelayCommand
    Public Property ClearFiltersCommand As RelayCommand
    Public Property RefreshCommand As RelayCommand

    Private Sub ClearForm()
        FormTeacherId = String.Empty
        FormFullName = String.Empty
        FormEmail = String.Empty
        FormContactNo = String.Empty
        FormDepartment = String.Empty
        FormStatus = "Active"
    End Sub

    Private Sub CopySelectedToForm()
        If _selectedTeacher Is Nothing Then Return
        FormTeacherId = _selectedTeacher.TeacherId
        FormFullName = _selectedTeacher.FullName
        FormEmail = _selectedTeacher.Email
        FormContactNo = _selectedTeacher.ContactNo
        FormDepartment = _selectedTeacher.Department
        FormStatus = _selectedTeacher.Status
    End Sub
End Class
