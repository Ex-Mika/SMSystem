Imports System.Collections.ObjectModel

Public Class RoomsViewModel
    Inherits ViewModelBase

    Private _selectedRoom As Room
    Private _formRoomCode As String
    Private _formRoomType As String
    Private _formCapacity As String
    Private _formStatus As String
    Private _searchText As String
    Private _selectedStatusFilter As String
    Private _selectedTypeFilter As String

    Public Sub New()
        Rooms = New ObservableCollection(Of Room)()
        StatusOptions = New List(Of String) From {"All", "Active", "Inactive"}
        TypeOptions = New List(Of String) From {"All", "Lecture", "Lab"}
        _selectedStatusFilter = "All"
        _selectedTypeFilter = "All"
        _formStatus = "Active"

        NewCommand = New RelayCommand(Sub(p) ClearForm())
        EditCommand = New RelayCommand(Sub(p) CopySelectedToForm(), Function(p) _selectedRoom IsNot Nothing)
        SaveCommand = New RelayCommand(Sub(p)
                                           ' No-op stub — will be wired to data layer later.
                                       End Sub)
        CancelCommand = New RelayCommand(Sub(p) ClearForm())
        DeactivateCommand = New RelayCommand(Sub(p)
                                                 ' No-op stub — will be wired to data layer later.
                                             End Sub)
        ClearFiltersCommand = New RelayCommand(Sub(p)
                                                   SearchText = String.Empty
                                                   SelectedStatusFilter = "All"
                                                   SelectedTypeFilter = "All"
                                               End Sub)
        RefreshCommand = New RelayCommand(Sub(p)
                                              ' No-op stub — will be wired to data layer later.
                                          End Sub)
    End Sub

    Public Property Rooms As ObservableCollection(Of Room)

    Public Property SelectedRoom As Room
        Get
            Return _selectedRoom
        End Get
        Set(value As Room)
            _selectedRoom = value
            OnPropertyChanged()
            CType(EditCommand, RelayCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Public Property FormRoomCode As String
        Get
            Return _formRoomCode
        End Get
        Set(value As String)
            _formRoomCode = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormRoomType As String
        Get
            Return _formRoomType
        End Get
        Set(value As String)
            _formRoomType = value
            OnPropertyChanged()
        End Set
    End Property

    Public Property FormCapacity As String
        Get
            Return _formCapacity
        End Get
        Set(value As String)
            _formCapacity = value
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

    Public Property SelectedTypeFilter As String
        Get
            Return _selectedTypeFilter
        End Get
        Set(value As String)
            _selectedTypeFilter = value
            OnPropertyChanged()
        End Set
    End Property

    Public ReadOnly Property StatusOptions As List(Of String)
    Public ReadOnly Property TypeOptions As List(Of String)

    Public Property NewCommand As RelayCommand
    Public Property EditCommand As RelayCommand
    Public Property SaveCommand As RelayCommand
    Public Property CancelCommand As RelayCommand
    Public Property DeactivateCommand As RelayCommand
    Public Property ClearFiltersCommand As RelayCommand
    Public Property RefreshCommand As RelayCommand

    Private Sub ClearForm()
        FormRoomCode = String.Empty
        FormRoomType = String.Empty
        FormCapacity = String.Empty
        FormStatus = "Active"
    End Sub

    Private Sub CopySelectedToForm()
        If _selectedRoom Is Nothing Then Return
        FormRoomCode = _selectedRoom.RoomCode
        FormRoomType = _selectedRoom.RoomType
        FormCapacity = _selectedRoom.Capacity.ToString()
        FormStatus = _selectedRoom.Status
    End Sub
End Class
