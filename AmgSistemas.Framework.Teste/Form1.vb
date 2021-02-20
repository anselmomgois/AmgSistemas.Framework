Imports AmgSistemas.Framework

Public Class Form1

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim items As New List(Of Empresa)

        items.Add(New Empresa With {.Codigo = "1", .Nome = "Anselmo", .CodAcesso = "1", .Filiais = New List(Of Filial)})
        items.Add(New Empresa With {.Codigo = "2", .Nome = "Josi", .CodAcesso = "1", .Filiais = New List(Of Filial)})


        'Dim Teste As List(Of WindowsForms.Item) = WindowsForms.Util.ConverterItems(items, "Codigo", "Nome")

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        'Dim resultado As Boolean = Util.ValidarValorCampo("eu@eu", Enumeradores.TipoValidacao.EMAIL)

        'If resultado Then

        'End If

    End Sub
End Class
Public Class Item2

    Public Property Codigo As String
    Public Property Descricao As String
    Public Property Lista3 As List(Of Item3)

End Class

Public Class Item3
    Public Property Identificador As String
End Class

Public Class Empresa

    Public Property Identificador As String
    Public Property Nome As String
    Public Property Codigo As Integer
    Public Property CodAcesso As String
    Public Property Filiais As List(Of Filial)

End Class

Public Class Filial
    Public Property Identificador As String
End Class