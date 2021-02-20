''' <summary>
''' Classe de apoio (Util) - Define a lista de criterios da cláusula "where"
''' </summary>
''' <remarks></remarks>
''' <history>
''' [pda] 13/02/2009 Criado
''' </history>
Public Class CriterioColecion
    Inherits List(Of Criterio)

    'Retorna um novo objeto critério
    Public Function getNewCriterio() As Criterio
        Dim objCriterio As New Criterio
        Return objCriterio
    End Function

    ''' <summary>
    ''' Adiciona um novo critério na lista
    ''' </summary>
    ''' <param name="Condicial"></param>
    ''' <param name="Clausula"></param>
    ''' <remarks></remarks>
    Public Sub addCriterio(ByVal Condicial As String, ByVal Clausula As String)
        Dim objCriterio As New Criterio
        objCriterio.Condicional = Condicial
        objCriterio.Clausula = Clausula
        Me.Add(objCriterio)
    End Sub

End Class