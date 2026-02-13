Public Class ClassSchedule
    Inherits ViewModelBase

    Private _subject As String
    Private _teacher As String
    Private _section As String
    Private _room As String
    Private _day As String
    Private _startTime As String
    Private _endTime As String

    Public Property Subject As String
        Get
            Return _subject
        End Get
        Set(value As String)
            _subject = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Teacher As String
        Get
            Return _teacher
        End Get
        Set(value As String)
            _teacher = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Section As String
        Get
            Return _section
        End Get
        Set(value As String)
            _section = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Room As String
        Get
            Return _room
        End Get
        Set(value As String)
            _room = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Day As String
        Get
            Return _day
        End Get
        Set(value As String)
            _day = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property StartTime As String
        Get
            Return _startTime
        End Get
        Set(value As String)
            _startTime = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property EndTime As String
        Get
            Return _endTime
        End Get
        Set(value As String)
            _endTime = value
            OnPropertyChanged()
        End Set
    End Property
End Class
