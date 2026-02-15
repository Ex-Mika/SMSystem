Imports System.Windows
Imports System.Windows.Media

Public NotInheritable Class RoundedClipHelper
    Private Sub New()
    End Sub

    Public Shared ReadOnly CornerRadiusProperty As DependencyProperty =
        DependencyProperty.RegisterAttached(
            "CornerRadius",
            GetType(Double),
            GetType(RoundedClipHelper),
            New PropertyMetadata(0.0, AddressOf OnCornerRadiusChanged))

    Public Shared Sub SetCornerRadius(element As DependencyObject, value As Double)
        element.SetValue(CornerRadiusProperty, value)
    End Sub

    Public Shared Function GetCornerRadius(element As DependencyObject) As Double
        Return CDbl(element.GetValue(CornerRadiusProperty))
    End Function

    Private Shared Sub OnCornerRadiusChanged(dependencyObject As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim element As FrameworkElement = TryCast(dependencyObject, FrameworkElement)
        If element Is Nothing Then
            Return
        End If

        RemoveHandler element.Loaded, AddressOf OnElementLoaded
        RemoveHandler element.SizeChanged, AddressOf OnElementSizeChanged

        Dim radius As Double = CDbl(e.NewValue)
        If radius <= 0 Then
            element.Clip = Nothing
            Return
        End If

        AddHandler element.Loaded, AddressOf OnElementLoaded
        AddHandler element.SizeChanged, AddressOf OnElementSizeChanged
        UpdateClip(element, radius)
    End Sub

    Private Shared Sub OnElementLoaded(sender As Object, e As RoutedEventArgs)
        Dim element As FrameworkElement = TryCast(sender, FrameworkElement)
        If element Is Nothing Then
            Return
        End If

        UpdateClip(element, GetCornerRadius(element))
    End Sub

    Private Shared Sub OnElementSizeChanged(sender As Object, e As SizeChangedEventArgs)
        Dim element As FrameworkElement = TryCast(sender, FrameworkElement)
        If element Is Nothing Then
            Return
        End If

        UpdateClip(element, GetCornerRadius(element))
    End Sub

    Private Shared Sub UpdateClip(element As FrameworkElement, radius As Double)
        If radius <= 0 OrElse element.ActualWidth <= 0 OrElse element.ActualHeight <= 0 Then
            element.Clip = Nothing
            Return
        End If

        Dim maxRadius As Double = Math.Min(element.ActualWidth, element.ActualHeight) / 2.0
        Dim safeRadius As Double = Math.Min(radius, maxRadius)

        element.Clip = New RectangleGeometry(
            New Rect(0, 0, element.ActualWidth, element.ActualHeight),
            safeRadius,
            safeRadius)
    End Sub
End Class
