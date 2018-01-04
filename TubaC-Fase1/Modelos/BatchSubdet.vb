Public Class BatchSubDET
    Private Nombre As String
    Public Property _Nombre() As String
        Get
            Return Nombre
        End Get
        Set(ByVal value As String)
            Nombre = value
        End Set
    End Property

    Private Cantidad As Double
    Public Property _Cantidad() As Double
        Get
            Return Cantidad
        End Get
        Set(ByVal value As Double)
            Cantidad = value
        End Set
    End Property
End Class
