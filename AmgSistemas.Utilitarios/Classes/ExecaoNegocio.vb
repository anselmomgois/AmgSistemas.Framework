Imports Microsoft.VisualBasic

<Serializable()> _
Public Class ExecaoNegocio
    Inherits System.Exception

    Public Sub New(ByVal codigo As Integer, ByVal descricao As String)
        _Codigo = codigo
        _Descricao = descricao
    End Sub

#Region " Variáveis "

    Private _Codigo As Integer
    Private _Descricao As String

#End Region

#Region " Propriedades "

    Public ReadOnly Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
    End Property

    Public ReadOnly Property Descricao() As String
        Get
            Return _Descricao
        End Get
    End Property

#End Region

End Class
