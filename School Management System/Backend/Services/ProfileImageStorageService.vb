Imports System.IO

Namespace Backend.Services
    Public Class ProfileImageStorageService
        Public Enum ProfileImageOwnerType
            Student
            Teacher
            Administrator
        End Enum

        Private ReadOnly _storageRootPath As String =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "SchoolManagementSystem",
                         "ProfileImages")

        Public Function StoreProfileImage(ownerType As ProfileImageOwnerType,
                                          ownerId As String,
                                          requestedPhotoPath As String,
                                          Optional existingPhotoPath As String = "") As String
            Dim normalizedRequestedPhotoPath As String =
                If(requestedPhotoPath, String.Empty).Trim()
            Dim normalizedExistingPhotoPath As String =
                If(existingPhotoPath, String.Empty).Trim()
            Dim normalizedOwnerId As String = If(ownerId, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedRequestedPhotoPath) Then
                Return String.Empty
            End If

            If String.IsNullOrWhiteSpace(normalizedOwnerId) Then
                Throw New IOException("A record ID is required before saving a profile image.")
            End If

            Dim targetPhotoPath As String =
                BuildManagedImagePath(ownerType,
                                      normalizedOwnerId,
                                      ResolveImageExtension(normalizedRequestedPhotoPath,
                                                            normalizedExistingPhotoPath))

            If String.Equals(normalizedRequestedPhotoPath,
                             targetPhotoPath,
                             StringComparison.OrdinalIgnoreCase) Then
                Return targetPhotoPath
            End If

            If Not File.Exists(normalizedRequestedPhotoPath) Then
                If String.Equals(normalizedRequestedPhotoPath,
                                 normalizedExistingPhotoPath,
                                 StringComparison.OrdinalIgnoreCase) Then
                    Return normalizedExistingPhotoPath
                End If

                Throw New FileNotFoundException("The selected profile image could not be found.",
                                                normalizedRequestedPhotoPath)
            End If

            Dim targetDirectoryPath As String = Path.GetDirectoryName(targetPhotoPath)

            If Not String.IsNullOrWhiteSpace(targetDirectoryPath) Then
                Directory.CreateDirectory(targetDirectoryPath)
            End If

            File.Copy(normalizedRequestedPhotoPath, targetPhotoPath, True)
            Return targetPhotoPath
        End Function

        Public Function IsManagedImagePath(photoPath As String) As Boolean
            Dim normalizedPhotoPath As String = If(photoPath, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedPhotoPath) Then
                Return False
            End If

            Try
                Dim fullStorageRootPath As String = EnsureTrailingSeparator(
                    Path.GetFullPath(_storageRootPath))
                Dim fullPhotoPath As String = Path.GetFullPath(normalizedPhotoPath)

                Return fullPhotoPath.StartsWith(fullStorageRootPath,
                                                StringComparison.OrdinalIgnoreCase)
            Catch
                Return False
            End Try
        End Function

        Public Sub DeleteManagedImage(photoPath As String)
            Dim normalizedPhotoPath As String = If(photoPath, String.Empty).Trim()

            If Not IsManagedImagePath(normalizedPhotoPath) Then
                Return
            End If

            Try
                If File.Exists(normalizedPhotoPath) Then
                    File.Delete(normalizedPhotoPath)
                End If
            Catch
            End Try
        End Sub

        Private Function BuildManagedImagePath(ownerType As ProfileImageOwnerType,
                                               ownerId As String,
                                               extension As String) As String
            Return Path.Combine(BuildOwnerDirectoryPath(ownerType),
                                SanitizeFileNameSegment(ownerId) &
                                NormalizeExtension(extension))
        End Function

        Private Function BuildOwnerDirectoryPath(ownerType As ProfileImageOwnerType) As String
            Select Case ownerType
                Case ProfileImageOwnerType.Student
                    Return Path.Combine(_storageRootPath, "Students")
                Case ProfileImageOwnerType.Teacher
                    Return Path.Combine(_storageRootPath, "Teachers")
                Case Else
                    Return Path.Combine(_storageRootPath, "Administrators")
            End Select
        End Function

        Private Function ResolveImageExtension(requestedPhotoPath As String,
                                               existingPhotoPath As String) As String
            Dim extension As String = Path.GetExtension(If(requestedPhotoPath, String.Empty).Trim())

            If String.IsNullOrWhiteSpace(extension) Then
                extension = Path.GetExtension(If(existingPhotoPath, String.Empty).Trim())
            End If

            Return NormalizeExtension(extension)
        End Function

        Private Function NormalizeExtension(extension As String) As String
            Dim normalizedExtension As String = If(extension, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedExtension) Then
                Return ".img"
            End If

            If Not normalizedExtension.StartsWith("."c) Then
                normalizedExtension = "." & normalizedExtension
            End If

            Return normalizedExtension.ToLowerInvariant()
        End Function

        Private Function SanitizeFileNameSegment(value As String) As String
            Dim normalizedValue As String = If(value, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedValue) Then
                Return "profile"
            End If

            For Each invalidCharacter As Char In Path.GetInvalidFileNameChars()
                normalizedValue = normalizedValue.Replace(invalidCharacter, "_"c)
            Next

            Return normalizedValue
        End Function

        Private Function EnsureTrailingSeparator(pathValue As String) As String
            Dim normalizedPathValue As String = If(pathValue, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedPathValue) Then
                Return String.Empty
            End If

            If normalizedPathValue.EndsWith(Path.DirectorySeparatorChar.ToString(),
                                            StringComparison.Ordinal) OrElse
               normalizedPathValue.EndsWith(Path.AltDirectorySeparatorChar.ToString(),
                                            StringComparison.Ordinal) Then
                Return normalizedPathValue
            End If

            Return normalizedPathValue & Path.DirectorySeparatorChar
        End Function
    End Class
End Namespace
