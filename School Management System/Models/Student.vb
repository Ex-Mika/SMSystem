Public Class Student
    Inherits ViewModelBase

    Private _studentId As String
    Private _firstName As String
    Private _middleName As String
    Private _lastName As String
    Private _sex As String
    Private _birthdate As String
    Private _address As String
    Private _contactNo As String
    Private _guardianName As String
    Private _guardianContact As String
    Private _course As String
    Private _yearLevel As String
    Private _status As String

    Public Property StudentId As String
        Get
            Return _studentId
        End Get
        Set(value As String)
            _studentId = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FirstName As String
        Get
            Return _firstName
        End Get
        Set(value As String)
            _firstName = value
            OnPropertyChanged()
            OnPropertyChanged("FullName")
        End Set
    End Property

    Public Property MiddleName As String
        Get
            Return _middleName
        End Get
        Set(value As String)
            _middleName = value
            OnPropertyChanged()
            OnPropertyChanged("FullName")
        End Set
    End Property

    Public Property LastName As String
        Get
            Return _lastName
        End Get
        Set(value As String)
            _lastName = value
            OnPropertyChanged()
            OnPropertyChanged("FullName")
        End Set
    End Property

    Public ReadOnly Property FullName As String
        Get
            Return $"{_firstName} {_middleName} {_lastName}".Trim()
        End Get
    End Property

    Public Property Sex As String
        Get
            Return _sex
        End Get
        Set(value As String)
            _sex = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Birthdate As String
        Get
            Return _birthdate
        End Get
        Set(value As String)
            _birthdate = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Address As String
        Get
            Return _address
        End Get
        Set(value As String)
            _address = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property ContactNo As String
        Get
            Return _contactNo
        End Get
        Set(value As String)
            _contactNo = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property GuardianName As String
        Get
            Return _guardianName
        End Get
        Set(value As String)
            _guardianName = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property GuardianContact As String
        Get
            Return _guardianContact
        End Get
        Set(value As String)
            _guardianContact = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Course As String
        Get
            Return _course
        End Get
        Set(value As String)
            _course = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property YearLevel As String
        Get
            Return _yearLevel
        End Get
        Set(value As String)
            _yearLevel = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Status As String
        Get
            Return _status
        End Get
        Set(value As String)
            _status = value
            OnPropertyChanged()
        End Set
    End Property
End Class
