Public Class BatchsubENC
    Private ITEMCODE As String
    Public Property _ITEMCODE() As String
        Get
            Return ITEMCODE
        End Get
        Set(ByVal value As String)
            ITEMCODE = value
        End Set
    End Property

    Private CATNIDADTOTAL As Double
    Public Property _CATNIDADTOTAL() As Double
        Get
            Return CATNIDADTOTAL
        End Get
        Set(ByVal value As Double)
            CATNIDADTOTAL = value
        End Set
    End Property

    Public _batchsubdet As List(Of BatchSubDET)

End Class
