Public Class Teacher
    Inherits ViewModelBase

    Private _teacherId As String
    Private _fullName As String
    Private _email As String
    Private _contactNo As String
    Private _department As String
    Private _status As String

    Public Property TeacherId As String
        Get
            Return _teacherId
        End Get
        Set(value As String)
            _teacherId = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FullName As String
        Get
            Return _fullName
        End Get
        Set(value As String)
            _fullName = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Email As String
        Get
            Return _email
        End Get
        Set(value As String)
            _email = value
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

    Public Property Department As String
        Get
            Return _department
        End Get
        Set(value As String)
            _department = value
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
