Imports School_Management_System.Backend.Common
Imports School_Management_System.Backend.Models
Imports School_Management_System.Backend.Repositories

Namespace Backend.Services
    Public Class AuthenticationService
        Private ReadOnly _userRepository As UserRepository

        Public Sub New()
            Me.New(New UserRepository())
        End Sub

        Public Sub New(userRepository As UserRepository)
            _userRepository = userRepository
        End Sub

        Public Function Authenticate(request As LoginRequest) As ServiceResult(Of UserAccount)
            If request Is Nothing Then
                Return ServiceResult(Of UserAccount).Failure("Login request is required.")
            End If

            If String.IsNullOrWhiteSpace(request.Identifier) Then
                Return ServiceResult(Of UserAccount).Failure("A login identifier is required.")
            End If

            If String.IsNullOrWhiteSpace(request.Password) Then
                Return ServiceResult(Of UserAccount).Failure("A password is required.")
            End If

            Dim user As UserAccount = _userRepository.GetByLoginIdentifier(request.Role,
                                                                           request.Identifier)
            If user Is Nothing Then
                Return ServiceResult(Of UserAccount).Failure("User account was not found.")
            End If

            If Not user.IsActive Then
                Return ServiceResult(Of UserAccount).Failure("User account is inactive.")
            End If

            If Not Database.DatabaseModule.VerifyPassword(request.Password,
                                                          user.PasswordHash) Then
                Return ServiceResult(Of UserAccount).Failure("Invalid password.")
            End If

            Return ServiceResult(Of UserAccount).Success(user, "Login successful.")
        End Function
    End Class
End Namespace
