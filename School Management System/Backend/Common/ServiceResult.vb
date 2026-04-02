Namespace Backend.Common
    Public Class ServiceResult(Of T)
        Public Property IsSuccess As Boolean
        Public Property Message As String = String.Empty
        Public Property Data As T

        Public Shared Function Success(data As T,
                                       Optional message As String = "") As ServiceResult(Of T)
            Return New ServiceResult(Of T) With {
                .IsSuccess = True,
                .Message = message,
                .Data = data
            }
        End Function

        Public Shared Function Failure(message As String) As ServiceResult(Of T)
            Return New ServiceResult(Of T) With {
                .IsSuccess = False,
                .Message = message
            }
        End Function
    End Class
End Namespace
