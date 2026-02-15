Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Text.Json

Class AdminAdministratorsView
    Private Class AdministratorStorageRecord
        Public Property AdministratorId As String
        Public Property FullName As String
        Public Property Role As String
        Public Property Email As String
    End Class

    Private _administratorsTable As DataTable
    Private _searchTerm As String = String.Empty
    Private ReadOnly _administratorsStoragePath As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SchoolManagementSystem", "administrators.json")
    Private ReadOnly _jsonOptions As New JsonSerializerOptions() With {
        .WriteIndented = True
    }

    Public Sub New()
        InitializeComponent()
        LoadAdministratorsTable()
    End Sub

    Public Sub ApplySearchFilter(searchTerm As String)
        _searchTerm = If(searchTerm, String.Empty).Trim()
        ApplyAdministratorsFilter()
    End Sub

    Private Sub LoadAdministratorsTable()
        _administratorsTable = ReadAdministratorsFromStorage()
        AdministratorsDataGrid.ItemsSource = _administratorsTable.DefaultView
        ApplyAdministratorsFilter()
    End Sub

    Private Function ReadAdministratorsFromStorage() As DataTable
        Dim table As DataTable = CreateEmptyAdministratorsTable()
        If Not File.Exists(_administratorsStoragePath) Then
            Return table
        End If

        Try
            Dim json As String = File.ReadAllText(_administratorsStoragePath)
            If String.IsNullOrWhiteSpace(json) Then
                Return table
            End If

            Dim records As List(Of AdministratorStorageRecord) =
                JsonSerializer.Deserialize(Of List(Of AdministratorStorageRecord))(json, _jsonOptions)
            If records Is Nothing Then
                Return table
            End If

            For Each record As AdministratorStorageRecord In records
                Dim row As DataRow = table.NewRow()
                row("Administrator ID") = If(record.AdministratorId, String.Empty).Trim()
                row("Full Name") = If(record.FullName, String.Empty).Trim()
                row("Role") = If(record.Role, String.Empty).Trim()
                row("Email") = If(record.Email, String.Empty).Trim()
                table.Rows.Add(row)
            Next
        Catch ex As Exception
            MessageBox.Show("Unable to load saved administrators data." & Environment.NewLine & ex.Message,
                            "Administrators",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning)
        End Try

        Return table
    End Function

    Private Function CreateEmptyAdministratorsTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("Administrator ID", GetType(String))
        table.Columns.Add("Full Name", GetType(String))
        table.Columns.Add("Role", GetType(String))
        table.Columns.Add("Email", GetType(String))
        Return table
    End Function

    Private Sub ApplyAdministratorsFilter()
        If _administratorsTable Is Nothing Then
            UpdateAdministratorsCount()
            Return
        End If

        If String.IsNullOrWhiteSpace(_searchTerm) Then
            _administratorsTable.DefaultView.RowFilter = String.Empty
        Else
            Dim escapedTerm As String = EscapeLikeValue(_searchTerm)
            _administratorsTable.DefaultView.RowFilter =
                "[Administrator ID] LIKE '*" & escapedTerm & "*' OR " &
                "[Full Name] LIKE '*" & escapedTerm & "*' OR " &
                "[Role] LIKE '*" & escapedTerm & "*' OR " &
                "[Email] LIKE '*" & escapedTerm & "*'"
        End If

        UpdateAdministratorsCount()
    End Sub

    Private Function EscapeLikeValue(value As String) As String
        Dim safeValue As String = If(value, String.Empty)
        Return safeValue.
            Replace("'", "''").
            Replace("[", "[[]").
            Replace("]", "[]]").
            Replace("*", "[*]").
            Replace("%", "[%]")
    End Function

    Private Sub UpdateAdministratorsCount()
        If AdministratorsCountTextBlock Is Nothing Then
            Return
        End If

        If _administratorsTable Is Nothing Then
            AdministratorsCountTextBlock.Text = "0 administrators"
            Return
        End If

        Dim visibleCount As Integer = _administratorsTable.DefaultView.Count
        Dim totalCount As Integer = _administratorsTable.Rows.Count
        Dim hasSearch As Boolean = Not String.IsNullOrWhiteSpace(_searchTerm)

        If hasSearch AndAlso visibleCount <> totalCount Then
            AdministratorsCountTextBlock.Text = visibleCount.ToString() & " of " & totalCount.ToString() & " administrators"
        Else
            AdministratorsCountTextBlock.Text = totalCount.ToString() & " administrators"
        End If
    End Sub
End Class
