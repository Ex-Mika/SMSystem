Public Class Course
    Inherits ViewModelBase

    Private _courseCode As String
    Private _courseName As String
    Private _status As String

    Public Property CourseCode As String
        Get
            Return _courseCode
        End Get
        Set(value As String)
            _courseCode = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property CourseName As String
        Get
            Return _courseName
        End Get
        Set(value As String)
            _courseName = value
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
