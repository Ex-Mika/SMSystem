Public Class Room
    Inherits ViewModelBase

    Private _roomCode As String
    Private _roomType As String
    Private _capacity As Integer
    Private _status As String

    Public Property RoomCode As String
        Get
            Return _roomCode
        End Get
        Set(value As String)
            _roomCode = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property RoomType As String
        Get
            Return _roomType
        End Get
        Set(value As String)
            _roomType = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property Capacity As Integer
        Get
            Return _capacity
        End Get
        Set(value As Integer)
            _capacity = value
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
