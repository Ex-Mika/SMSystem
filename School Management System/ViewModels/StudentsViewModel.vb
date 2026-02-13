Imports System.Collections.ObjectModel

Public Class StudentsViewModel
    Inherits ViewModelBase

    Private _selectedStudent As Student
    Private _formStudentId As String
    Private _formFirstName As String
    Private _formMiddleName As String
    Private _formLastName As String
    Private _formSex As String
    Private _formBirthdate As String
    Private _formAddress As String
    Private _formContactNo As String
    Private _formGuardianName As String
    Private _formGuardianContact As String
    Private _formCourse As String
    Private _formYearLevel As String
    Private _formStatus As String
    Private _searchText As String
    Private _selectedStatusFilter As String
    Private _selectedCourseFilter As String

    Public Sub New()
        Students = New ObservableCollection(Of Student)()
        StatusOptions = New List(Of String) From {"All", "Active", "Graduated", "Blocked"}
        CourseOptions = New List(Of String) From {"All"}
        YearLevelOptions = New List(Of String) From {"1st Year", "2nd Year", "3rd Year", "4th Year"}
        SexOptions = New List(Of String) From {"Male", "Female"}
        _selectedStatusFilter = "All"
        _selectedCourseFilter = "All"
        _formStatus = "Active"

        NewCommand = New RelayCommand(Sub(p) ClearForm())
        EditCommand = New RelayCommand(Sub(p) CopySelectedToForm(), Function(p) _selectedStudent IsNot Nothing)
        SaveCommand = New RelayCommand(Sub(p)
                                           ' No-op stub — will be wired to data layer later.
                                       End Sub)
        CancelCommand = New RelayCommand(Sub(p) ClearForm())
        GraduateCommand = New RelayCommand(Sub(p)
                                               ' No-op stub — will be wired to data layer later.
                                           End Sub)
        BlockCommand = New RelayCommand(Sub(p)
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

    Public Property Students As ObservableCollection(Of Student)

    Public Property SelectedStudent As Student
        Get
            Return _selectedStudent
        End Get
        Set(value As Student)
            _selectedStudent = value
            OnPropertyChanged()
            CType(EditCommand, RelayCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Public Property FormStudentId As String
        Get
            Return _formStudentId
        End Get
        Set(value As String)
            _formStudentId = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormFirstName As String
        Get
            Return _formFirstName
        End Get
        Set(value As String)
            _formFirstName = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormMiddleName As String
        Get
            Return _formMiddleName
        End Get
        Set(value As String)
            _formMiddleName = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormLastName As String
        Get
            Return _formLastName
        End Get
        Set(value As String)
            _formLastName = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormSex As String
        Get
            Return _formSex
        End Get
        Set(value As String)
            _formSex = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormBirthdate As String
        Get
            Return _formBirthdate
        End Get
        Set(value As String)
            _formBirthdate = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormAddress As String
        Get
            Return _formAddress
        End Get
        Set(value As String)
            _formAddress = value
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

    Public Property FormGuardianName As String
        Get
            Return _formGuardianName
        End Get
        Set(value As String)
            _formGuardianName = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormGuardianContact As String
        Get
            Return _formGuardianContact
        End Get
        Set(value As String)
            _formGuardianContact = value
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

    Public ReadOnly Property StatusOptions As List(Of String)
    Public ReadOnly Property CourseOptions As List(Of String)
    Public ReadOnly Property YearLevelOptions As List(Of String)
    Public ReadOnly Property SexOptions As List(Of String)

    Public Property NewCommand As RelayCommand
    Public Property EditCommand As RelayCommand
    Public Property SaveCommand As RelayCommand
    Public Property CancelCommand As RelayCommand
    Public Property GraduateCommand As RelayCommand
    Public Property BlockCommand As RelayCommand
    Public Property ClearFiltersCommand As RelayCommand
    Public Property RefreshCommand As RelayCommand

    Private Sub ClearForm()
        FormStudentId = String.Empty
        FormFirstName = String.Empty
        FormMiddleName = String.Empty
        FormLastName = String.Empty
        FormSex = String.Empty
        FormBirthdate = String.Empty
        FormAddress = String.Empty
        FormContactNo = String.Empty
        FormGuardianName = String.Empty
        FormGuardianContact = String.Empty
        FormCourse = String.Empty
        FormYearLevel = String.Empty
        FormStatus = "Active"
    End Sub

    Private Sub CopySelectedToForm()
        If _selectedStudent Is Nothing Then Return
        FormStudentId = _selectedStudent.StudentId
        FormFirstName = _selectedStudent.FirstName
        FormMiddleName = _selectedStudent.MiddleName
        FormLastName = _selectedStudent.LastName
        FormSex = _selectedStudent.Sex
        FormBirthdate = _selectedStudent.Birthdate
        FormAddress = _selectedStudent.Address
        FormContactNo = _selectedStudent.ContactNo
        FormGuardianName = _selectedStudent.GuardianName
        FormGuardianContact = _selectedStudent.GuardianContact
        FormCourse = _selectedStudent.Course
        FormYearLevel = _selectedStudent.YearLevel
        FormStatus = _selectedStudent.Status
    End Sub
End Class
