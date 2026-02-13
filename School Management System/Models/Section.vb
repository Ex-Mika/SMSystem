Public Class Section
    Inherits ViewModelBase

    Private _sectionName As String
    Private _course As String
    Private _yearLevel As String
    Private _status As String

    Public Property SectionName As String
        Get
            Return _sectionName
        End Get
        Set(value As String)
            _sectionName = value
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
