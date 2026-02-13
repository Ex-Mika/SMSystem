Imports System.Collections.ObjectModel

Public Class SchedulingViewModel
    Inherits ViewModelBase

    Private _selectedSchedule As ClassSchedule
    Private _formSubject As String
    Private _formTeacher As String
    Private _formSection As String
    Private _formRoom As String
    Private _formDay As String
    Private _formStartTime As String
    Private _formEndTime As String
    Private _searchText As String
    Private _selectedCourseFilter As String
    Private _selectedSectionFilter As String
    Private _selectedRoomFilter As String
    Private _selectedDayFilter As String
    Private _selectedTabIndex As Integer

    ' Teacher View
    Private _selectedTeacherFilter As String

    ' Section View
    Private _sectionViewCourseFilter As String
    Private _sectionViewSectionFilter As String

    Public Sub New()
        Schedules = New ObservableCollection(Of ClassSchedule)()
        TeacherViewSchedules = New ObservableCollection(Of ClassSchedule)()
        SectionViewSchedules = New ObservableCollection(Of ClassSchedule)()

        CourseOptions = New List(Of String) From {"All"}
        SectionOptions = New List(Of String) From {"All"}
        RoomOptions = New List(Of String) From {"All"}
        DayOptions = New List(Of String) From {"All", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
        SubjectOptions = New List(Of String) From {}
        TeacherOptions = New List(Of String) From {"All"}

        _selectedCourseFilter = "All"
        _selectedSectionFilter = "All"
        _selectedRoomFilter = "All"
        _selectedDayFilter = "All"
        _selectedTeacherFilter = "All"
        _sectionViewCourseFilter = "All"
        _sectionViewSectionFilter = "All"
        _selectedTabIndex = 0

        AddCommand = New RelayCommand(Sub(p)
                                          ' No-op stub — will be wired to data layer later.
                                      End Sub)
        UpdateCommand = New RelayCommand(Sub(p)
                                             ' No-op stub — will be wired to data layer later.
                                         End Sub)
        RemoveCommand = New RelayCommand(Sub(p)
                                             ' No-op stub — will be wired to data layer later.
                                         End Sub)
        ClearCommand = New RelayCommand(Sub(p) ClearForm())
        ClearFiltersCommand = New RelayCommand(Sub(p)
                                                   SelectedCourseFilter = "All"
                                                   SelectedSectionFilter = "All"
                                                   SelectedRoomFilter = "All"
                                                   SelectedDayFilter = "All"
                                               End Sub)
        RefreshCommand = New RelayCommand(Sub(p)
                                              ' No-op stub — will be wired to data layer later.
                                          End Sub)
    End Sub

    Public Property Schedules As ObservableCollection(Of ClassSchedule)
    Public Property TeacherViewSchedules As ObservableCollection(Of ClassSchedule)
    Public Property SectionViewSchedules As ObservableCollection(Of ClassSchedule)

    Public Property SelectedSchedule As ClassSchedule
        Get
            Return _selectedSchedule
        End Get
        Set(value As ClassSchedule)
            _selectedSchedule = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property SelectedTabIndex As Integer
        Get
            Return _selectedTabIndex
        End Get
        Set(value As Integer)
            _selectedTabIndex = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormSubject As String
        Get
            Return _formSubject
        End Get
        Set(value As String)
            _formSubject = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormTeacher As String
        Get
            Return _formTeacher
        End Get
        Set(value As String)
            _formTeacher = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormSection As String
        Get
            Return _formSection
        End Get
        Set(value As String)
            _formSection = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormRoom As String
        Get
            Return _formRoom
        End Get
        Set(value As String)
            _formRoom = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormDay As String
        Get
            Return _formDay
        End Get
        Set(value As String)
            _formDay = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormStartTime As String
        Get
            Return _formStartTime
        End Get
        Set(value As String)
            _formStartTime = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormEndTime As String
        Get
            Return _formEndTime
        End Get
        Set(value As String)
            _formEndTime = value
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

    Public Property SelectedDayFilter As String
        Get
            Return _selectedDayFilter
        End Get
        Set(value As String)
            _selectedDayFilter = value
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

    Public Property SelectedTeacherFilter As String
        Get
            Return _selectedTeacherFilter
        End Get
        Set(value As String)
            _selectedTeacherFilter = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property SectionViewCourseFilter As String
        Get
            Return _sectionViewCourseFilter
        End Get
        Set(value As String)
            _sectionViewCourseFilter = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property SectionViewSectionFilter As String
        Get
            Return _sectionViewSectionFilter
        End Get
        Set(value As String)
            _sectionViewSectionFilter = value
            OnPropertyChanged()
        End Set
    End Property

    Public ReadOnly Property CourseOptions As List(Of String)
    Public ReadOnly Property SectionOptions As List(Of String)
    Public ReadOnly Property RoomOptions As List(Of String)
    Public ReadOnly Property DayOptions As List(Of String)
    Public ReadOnly Property SubjectOptions As List(Of String)
    Public ReadOnly Property TeacherOptions As List(Of String)

    Public Property AddCommand As RelayCommand
    Public Property UpdateCommand As RelayCommand
    Public Property RemoveCommand As RelayCommand
    Public Property ClearCommand As RelayCommand
    Public Property ClearFiltersCommand As RelayCommand
    Public Property RefreshCommand As RelayCommand

    Private Sub ClearForm()
        FormSubject = String.Empty
        FormTeacher = String.Empty
        FormSection = String.Empty
        FormRoom = String.Empty
        FormDay = String.Empty
        FormStartTime = String.Empty
        FormEndTime = String.Empty
    End Sub
End Class
