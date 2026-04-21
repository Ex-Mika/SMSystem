Imports System.IO
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Friend NotInheritable Class DashboardProfileImageHelper
    Private Sub New()
    End Sub

    Public Shared Sub ApplyProfilePhoto(avatarBorder As Border,
                                        fallbackTextBlock As TextBlock,
                                        photoPath As String,
                                        fallbackBackground As Brush)
        If avatarBorder Is Nothing OrElse fallbackTextBlock Is Nothing Then
            Return
        End If

        Dim resolvedFallbackBackground As Brush = If(fallbackBackground, Brushes.Transparent)
        Dim normalizedPhotoPath As String = If(photoPath, String.Empty).Trim()

        If String.IsNullOrWhiteSpace(normalizedPhotoPath) Then
            avatarBorder.Background = resolvedFallbackBackground
            fallbackTextBlock.Visibility = Visibility.Visible
            Return
        End If

        Try
            Dim resolvedPhotoPath As String = Path.GetFullPath(normalizedPhotoPath)
            If Not File.Exists(resolvedPhotoPath) Then
                avatarBorder.Background = resolvedFallbackBackground
                fallbackTextBlock.Visibility = Visibility.Visible
                Return
            End If

            avatarBorder.Background = CreatePhotoBrush(resolvedPhotoPath)
            fallbackTextBlock.Visibility = Visibility.Collapsed
        Catch
            avatarBorder.Background = resolvedFallbackBackground
            fallbackTextBlock.Visibility = Visibility.Visible
        End Try
    End Sub

    Private Shared Function CreatePhotoBrush(photoPath As String) As ImageBrush
        Dim bitmap As New BitmapImage()
        bitmap.BeginInit()
        bitmap.CacheOption = BitmapCacheOption.OnLoad
        bitmap.UriSource = New Uri(photoPath, UriKind.Absolute)
        bitmap.EndInit()
        bitmap.Freeze()

        Dim photoBrush As New ImageBrush(bitmap) With {
            .Stretch = Stretch.UniformToFill,
            .AlignmentX = AlignmentX.Center,
            .AlignmentY = AlignmentY.Center
        }
        photoBrush.Freeze()

        Return photoBrush
    End Function
End Class
