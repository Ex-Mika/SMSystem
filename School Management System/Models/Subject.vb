Public Class Subject
    Inherits ViewModelBase

    Private _subjectCode As String
    Private _description As String
    Private _units As Integer
    Private _course As String
    Private _status As String

    Public Property SubjectCode As String
        Get
            Return _subjectCode
        End Get
        Set(value As String)
            _subjectCode = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Units As Integer
        Get
            Return _units
        End Get
        Set(value As Integer)
            _units = value
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
