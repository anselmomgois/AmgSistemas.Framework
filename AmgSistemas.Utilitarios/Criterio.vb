''' <summary>
''' Classe de apoio - Criterio
''' </summary>
''' <remarks></remarks>
''' <history>
''' [pda] 13/02/2009 Criado
''' </history>
Public Class Criterio
    Dim _Condicional As String
    Dim _Clausula As String
    Public Property Condicional() As String
        Get
            Return _Condicional
        End Get
        Set(ByVal value As String)
            _Condicional = value
        End Set
    End Property

    Public Property Clausula() As String
        Get
            Return _Clausula
        End Get
        Set(ByVal value As String)
            _Clausula = value
        End Set
    End Property
End Class
