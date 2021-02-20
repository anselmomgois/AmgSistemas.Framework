Public Class Util

    ''' <summary>
    ''' Monta a clausula in através de uma coleção de itens e o nome do campo
    ''' </summary>
    ''' <param name="Itens"></param>
    ''' <param name="Campo"></param>
    ''' <param name="frm"></param>
    ''' <param name="TipoClausula">Utilizado para informar se a clausula é WHERE, AND ou OR.</param>
    ''' <param name="Alias">Insere um alias para o campo de consulta.</param>
    ''' <param name="Dif">Utilizado para diferenciar o campo que será adicionado no parameter. Isto evita problemas com campos repetidos no parameter.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [anselmo.gois] 13/05/2009 Criado
    ''' </history>
    Public Shared Function MontarClausulaIn(ByVal Itens As Object, _
                                            ByVal Campo As String, _
                                            ByRef frm As Sql, _
                                            Optional ByVal IstanciarParametros As Boolean = False, _
                                            Optional ByVal TipoClausula As String = "", _
                                            Optional ByVal [Alias] As String = "", _
                                            Optional ByVal Dif As String = "", _
                                            Optional ByVal EsNotIn As Boolean = False) As String

        ' clausula in
        Dim clausulaIn As New System.Text.StringBuilder

        ' se alias não for vazio
        If Not [Alias].Equals(String.Empty) Then
            [Alias] &= "."
        End If

        ' se foram informados itens para pesquisa
        If Itens IsNot Nothing AndAlso Itens.Count > 0 Then

            ' criar flag
            Dim addIn As Boolean = True

            ' percorrer todos os itens
            For i As Integer = 0 To Itens.Count - 1

                If Itens(i) IsNot Nothing AndAlso Not Itens(i).Equals(String.Empty) AndAlso addIn Then

                    If EsNotIn Then

                        ' concatenar filtro na query
                        clausulaIn.Append(" " & TipoClausula & " " & [Alias] & Campo & " NOT IN (")

                    Else

                        ' concatenar filtro na query
                        clausulaIn.Append(" " & TipoClausula & " " & [Alias] & Campo & " IN (")

                    End If


                    ' alterar flag 
                    addIn = False

                ElseIf Itens(i) Is Nothing OrElse Itens(i).Equals(String.Empty) Then
                    Continue For
                End If

                ' concatenar parametro na query
                clausulaIn.Append("@" & Dif & Campo & i)

                ' se ainda existirem codigos
                If i <> Itens.Count - 1 Then
                    clausulaIn.Append(",")
                End If

                ' setar parameter
                frm.AdicionarParametro(Dif & Campo & i, Itens(i), IstanciarParametros)

                IstanciarParametros = False

            Next

            ' se adicionou IN deve fechar parenteses
            If Not addIn Then

                ' fechar comando do filtro
                clausulaIn.Append(")")

            End If

        End If

        Return clausulaIn.ToString

    End Function

    ''' <summary>
    ''' Utilizado para validar se o valor informado é igual a DBNull.Value. Se for, retorna nothing
    ''' </summary>
    ''' <param name="Valor"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [blcosta] 17/11/2009 criado
    ''' </history>
    Public Shared Function VerificarDbNull(ByVal Valor As Object) As Object

        If IsDBNull(Valor) Then
            Return Nothing
        Else
            Return Valor
        End If

    End Function

End Class
