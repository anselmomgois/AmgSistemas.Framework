Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO
Imports System.Text

Public Class Util

    ''' <summary>
    ''' Atribui o valor ao objeto passado, faz a conversão do tipo do banco para o tipo da propriedade.
    ''' </summary>
    ''' <param name="Campo"></param>
    ''' <param name="Valor"></param>
    ''' <param name="TipoCampo"></param>
    ''' <remarks></remarks>
    ''' <history>
    ''' </history>
    Public Shared Sub AtribuirValorObjeto(ByRef Campo As Object, _
                                          ByVal Valor As Object, _
                                          ByVal TipoCampo As System.Type)

        If Valor IsNot DBNull.Value AndAlso Not String.IsNullOrEmpty(Valor) Then
            If TipoCampo Is Nothing Then
                Campo = Valor
            ElseIf TipoCampo Is GetType(String) Then
                Campo = Convert.ToString(Valor)
            ElseIf (TipoCampo Is GetType(Int16) OrElse TipoCampo Is GetType(Nullable(Of Int16))) Then
                Campo = Convert.ToInt16(Valor)
            ElseIf (TipoCampo Is GetType(Int32) OrElse TipoCampo Is GetType(Nullable(Of Int32))) Then
                Campo = Convert.ToInt32(Valor)
            ElseIf (TipoCampo Is GetType(Int64) OrElse TipoCampo Is GetType(Nullable(Of Int64))) Then
                Campo = Convert.ToInt64(Valor)
            ElseIf (TipoCampo Is GetType(Decimal) OrElse TipoCampo Is GetType(Nullable(Of Decimal))) Then
                Campo = Convert.ToDecimal(Valor)
            ElseIf TipoCampo Is GetType(Double) Then
                Campo = Convert.ToDouble(Valor)
            ElseIf TipoCampo Is GetType(Boolean) Then
                Campo = Convert.ToBoolean(Valor.ToString.Trim)
            End If
        Else
            If TipoCampo Is GetType(Boolean) Then
                Campo = Convert.ToBoolean(False)
            ElseIf TipoCampo Is GetType(Double) Then
                Campo = Convert.ToDouble(0)
            ElseIf TipoCampo Is GetType(Decimal) Then
                Campo = Convert.ToDecimal(0)
            ElseIf TipoCampo Is GetType(Int16) Then
                Campo = Convert.ToInt16(0)
            ElseIf TipoCampo Is GetType(Int32) Then
                Campo = Convert.ToInt32(0)
            ElseIf TipoCampo Is GetType(Int64) Then
                Campo = Convert.ToInt64(0)
            Else
                Campo = Nothing
            End If

        End If

    End Sub

    Public Shared Function AtribuirValorObj(ByVal Valor As Object, _
                                            ByVal TipoCampo As System.Type) As Object

        Dim campo As New Object

        If Valor IsNot DBNull.Value AndAlso TipoCampo IsNot GetType(Byte()) AndAlso Not String.IsNullOrEmpty(Valor) Then
            If TipoCampo Is Nothing Then
                campo = Valor
            ElseIf TipoCampo Is GetType(String) Then
                campo = Convert.ToString(Valor)
            ElseIf (TipoCampo Is GetType(Int16) OrElse TipoCampo Is GetType(Nullable(Of Int16))) Then
                campo = Convert.ToInt16(Valor)
            ElseIf (TipoCampo Is GetType(Int32) OrElse TipoCampo Is GetType(Nullable(Of Int32))) Then
                campo = Convert.ToInt32(Valor)
            ElseIf (TipoCampo Is GetType(Int64) OrElse TipoCampo Is GetType(Nullable(Of Int64))) Then
                campo = Convert.ToInt64(Valor)
            ElseIf (TipoCampo Is GetType(Decimal) OrElse TipoCampo Is GetType(Nullable(Of Decimal))) Then
                campo = Convert.ToDecimal(Valor)
            ElseIf TipoCampo Is GetType(Double) Then
                campo = Convert.ToDouble(Valor)
            ElseIf TipoCampo Is GetType(Boolean) Then
                campo = Convert.ToBoolean(Valor.ToString.Trim)
            ElseIf TipoCampo Is GetType(DateTime) Then
                campo = Convert.ToDateTime(Valor)
            ElseIf TipoCampo Is GetType(Nullable(Of DateTime)) Then
                campo = Convert.ToDateTime(Valor)
            End If
        ElseIf Valor IsNot DBNull.Value AndAlso TipoCampo Is GetType(Byte()) Then
            campo = DirectCast(Valor, Byte())
        Else
            If TipoCampo Is GetType(Boolean) Then
                campo = Convert.ToBoolean(False)
            ElseIf TipoCampo Is GetType(Double) Then
                campo = Convert.ToDouble(0)
            ElseIf TipoCampo Is GetType(Decimal) Then
                campo = Convert.ToDecimal(0)
            ElseIf TipoCampo Is GetType(Int16) Then
                campo = Convert.ToInt16(0)
            ElseIf TipoCampo Is GetType(Int32) Then
                campo = Convert.ToInt32(0)
            ElseIf TipoCampo Is GetType(Int64) Then
                campo = Convert.ToInt64(0)
            ElseIf TipoCampo Is GetType(DateTime) Then
                campo = Nothing
            Else
                campo = Nothing
            End If

        End If

        Return campo
    End Function

    Public Shared Function AtribuirValorObj2(ByVal Valor As Object, _
                                             ByVal TipoCampo As System.Type,
                                             ByVal CampoIsNullable As Boolean) As Object

        Dim campo As New Object

        If Valor IsNot DBNull.Value AndAlso TipoCampo IsNot GetType(Byte()) AndAlso Not String.IsNullOrEmpty(Valor) Then
            If TipoCampo Is Nothing Then
                campo = Valor
            ElseIf TipoCampo Is GetType(String) Then
                campo = Convert.ToString(Valor)
            ElseIf (TipoCampo Is GetType(Int16) OrElse TipoCampo Is GetType(Nullable(Of Int16))) Then
                campo = Convert.ToInt16(Valor)
            ElseIf (TipoCampo Is GetType(Int32) OrElse TipoCampo Is GetType(Nullable(Of Int32))) Then
                campo = Convert.ToInt32(Valor)
            ElseIf (TipoCampo Is GetType(Int64) OrElse TipoCampo Is GetType(Nullable(Of Int64))) Then
                campo = Convert.ToInt64(Valor)
            ElseIf (TipoCampo Is GetType(Decimal) OrElse TipoCampo Is GetType(Nullable(Of Decimal))) Then
                campo = Convert.ToDecimal(Valor)
            ElseIf TipoCampo Is GetType(Double) Then
                campo = Convert.ToDouble(Valor)
            ElseIf TipoCampo Is GetType(Boolean) Then
                campo = Convert.ToBoolean(Valor.ToString.Trim)
            ElseIf TipoCampo Is GetType(DateTime) Then
                campo = Convert.ToDateTime(Valor)
            ElseIf TipoCampo Is GetType(Nullable(Of DateTime)) Then
                campo = Convert.ToDateTime(Valor)
            End If
        ElseIf Valor IsNot DBNull.Value AndAlso TipoCampo Is GetType(Byte()) Then
            campo = DirectCast(Valor, Byte())
        ElseIf CampoIsNullable Then
            campo = Nothing
        Else
            If TipoCampo Is GetType(Boolean) Then
                campo = Convert.ToBoolean(False)
            ElseIf TipoCampo Is GetType(Double) Then
                campo = Convert.ToDouble(0)
            ElseIf TipoCampo Is GetType(Decimal) Then
                campo = Convert.ToDecimal(0)
            ElseIf TipoCampo Is GetType(Int16) Then
                campo = Convert.ToInt16(0)
            ElseIf TipoCampo Is GetType(Int32) Then
                campo = Convert.ToInt32(0)
            ElseIf TipoCampo Is GetType(Int64) Then
                campo = Convert.ToInt64(0)
            ElseIf TipoCampo Is GetType(DateTime) Then
                campo = Nothing
            Else
                campo = Nothing
            End If

        End If

        Return campo
    End Function


    ''' <summary>
    ''' Retorna cópia do objeto passado como parâmetro
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ClonarObjeto(ByVal obj As Object) As Object

        If obj Is Nothing Then
            Return Nothing
        End If

        ' Create a memory stream and a formatter.
        Dim ms As New MemoryStream()
        Dim bf As New BinaryFormatter()

        ' Serialize the object into the stream.
        bf.Serialize(ms, obj)

        ' Position streem pointer back to first byte.
        ms.Seek(0, SeekOrigin.Begin)

        ' Deserialize into another object.
        ClonarObjeto = bf.Deserialize(ms)

        ' Release memory.
        ms.Close()

    End Function

    Public Shared Function RetornaDbNull(ByVal valor As String, Optional ByVal RetornarValorComUpper As Boolean = False) As String

        If String.IsNullOrEmpty(valor) Then
            Return DBNull.Value.ToString
        End If

        Return If(RetornarValorComUpper, valor.ToUpper, valor)
    End Function

    Public Shared Function RetornaDbNull(ByVal valor As Object) As Object

        If valor Is Nothing Then
            Return DBNull.Value
        End If

        Return valor
    End Function

    Public Shared Sub ValidarCampo(ByVal campo As Object, ByVal msgErro As String, ByVal type As System.Type, ByVal BolColecao As Boolean)

        If type Is GetType(String) Then

            If String.IsNullOrEmpty(campo) Then
                Throw New ExecaoNegocio(Erro.TipoErro.ERRO_PARAMETRO, msgErro)
            End If

        ElseIf type Is GetType(Nullable(Of Boolean)) Then

            If campo Is Nothing Then
                Throw New ExecaoNegocio(Erro.TipoErro.ERRO_PARAMETRO, msgErro)
            End If

        ElseIf type Is GetType(DateTime) Then

            If campo.Equals(Date.MinValue) Then
                Throw New ExecaoNegocio(Erro.TipoErro.ERRO_PARAMETRO, msgErro)
            End If

        ElseIf type Is GetType(Nullable(Of Decimal)) Then

            If campo Is Nothing Then
                Throw New ExecaoNegocio(Erro.TipoErro.ERRO_PARAMETRO, msgErro)
            End If

        ElseIf type Is GetType(Nullable(Of Integer)) Then

            If campo Is Nothing Then
                Throw New ExecaoNegocio(Erro.TipoErro.ERRO_PARAMETRO, msgErro)
            End If
        ElseIf type Is GetType(Integer) Then

            If campo = 0 Then
                Throw New ExecaoNegocio(Erro.TipoErro.ERRO_PARAMETRO, msgErro)
            End If

        Else

            If BolColecao Then

                If campo Is Nothing OrElse campo.count = 0 Then
                    Throw New ExecaoNegocio(Erro.TipoErro.ERRO_PARAMETRO, msgErro)
                End If

            Else

                If campo Is Nothing Then
                    Throw New ExecaoNegocio(Erro.TipoErro.ERRO_PARAMETRO, msgErro)
                End If

            End If
        End If
    End Sub

    Public Shared Sub ValidarCampo(ByVal campo As Object, ByVal msgErro As String, ByVal type As System.Type, ByVal BolColecao As Boolean,
                                   ByRef Erros As System.Text.StringBuilder)

        If type Is GetType(String) Then

            If String.IsNullOrEmpty(campo) Then
                Erros.AppendLine(msgErro)
            End If

        ElseIf type Is GetType(Nullable(Of Boolean)) Then

            If campo Is Nothing Then
                Erros.AppendLine(msgErro)
            End If

        ElseIf type Is GetType(DateTime) Then

            If campo.Equals(Date.MinValue) Then
                Erros.AppendLine(msgErro)
            End If

        ElseIf type Is GetType(Nullable(Of Decimal)) Then

            If campo Is Nothing Then
                Erros.AppendLine(msgErro)
            End If

        ElseIf type Is GetType(Nullable(Of Integer)) Then

            If campo Is Nothing Then
                Erros.AppendLine(msgErro)
            End If
        ElseIf type Is GetType(Integer) Then

            If campo = 0 Then
                Erros.AppendLine(msgErro)
            End If

        Else

            If BolColecao Then

                If campo Is Nothing OrElse campo.count = 0 Then
                    Erros.AppendLine(msgErro)
                End If

            Else

                If campo Is Nothing Then
                    Erros.AppendLine(msgErro)
                End If

            End If
        End If
    End Sub

    ''' <summary>
    ''' Garante a entrada apenas de valores númericos, virgula e ponto
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <history>
    ''' [octavio.piramo] 03/09/2008 Criado
    ''' </history>
    Public Shared Sub NumeroVirgulaPonto(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        Try

            If Not Char.IsNumber(e.KeyChar) And Not e.KeyChar = vbBack And Not e.KeyChar = "." And Not e.KeyChar = "," Then
                e.Handled = True
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' Permite a entreda apenas de numeros
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <history>
    ''' [octavio.piramo] 03/09/2008 Criado
    ''' </history>
    Public Shared Sub SomenteNumero(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        Try

            If Not Char.IsNumber(e.KeyChar) AndAlso Not e.KeyChar = vbBack Then
                e.Handled = True
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' Monta a clausula in através de uma coleção de itens e o nome do campo
    ''' </summary>
    ''' <param name="Itens"></param>
    ''' <param name="Campo"></param>
    ''' <param name="Parametros"></param>
    ''' <param name="TipoClausula">Utilizado para informar se a clausula é WHERE, AND ou OR.</param>
    ''' <param name="Alias">Insere um alias para o campo de consulta.</param>
    ''' <param name="Dif">Utilizado para diferenciar o campo que será adicionado no parameter. Isto evita problemas com campos repetidos no parameter.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [octavio.piramo] 03/02/2009 Criado
    ''' </history>
    Public Shared Function MontarClausulaIn(ByVal Itens As Object, _
                                            ByVal Campo As String, _
                                            ByRef Parametros As AcessoDados.ParametroColecao, _
                                            Optional ByVal TipoClausula As String = "", _
                                            Optional ByVal [Alias] As String = "", _
                                            Optional ByVal Dif As String = "") As String

        ' clausula in
        Dim clausulaIn As New StringBuilder

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

                    ' concatenar filtro na query
                    clausulaIn.Append(" " & TipoClausula & " " & [Alias] & Campo & " IN (")

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

                If Parametros Is Nothing Then Parametros = New AcessoDados.ParametroColecao

                Parametros.Add(New AcessoDados.Parametro With {.Campo = Campo & i, _
                                                               .Valor = Itens(i)})

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
    ''' Monta a clausula like através de uma coleção de itens e o nome do campo
    ''' </summary>
    ''' <param name="Itens"></param>
    ''' <param name="Campo"></param>
    ''' <param name="Parametros"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [anselmo.gois] 11/02/2009 Created
    ''' </history>
    Public Shared Function MontarClausulaLike(ByVal Itens As Object, _
                                              ByVal Campo As String, _
                                              ByRef Parametros As AcessoDados.ParametroColecao, _
                                              Optional ByVal TipoClausula As String = "", _
                                              Optional ByVal [Alias] As String = "") As String

        ' clausula in
        Dim clausulalike As New StringBuilder

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

                    ' concatenar filtro na query
                    clausulalike.Append(" " & TipoClausula & " (" & [Alias] & Campo & " LIKE (")

                    ' alterar flag 
                    addIn = False

                ElseIf Not addIn Then
                    clausulalike.Append(" OR " & [Alias] & Campo & " LIKE (")
                ElseIf Itens(i) Is Nothing OrElse Itens(i).Equals(String.Empty) Then
                    Continue For
                End If

                ' concatenar parametro na query
                clausulalike.Append("@" & Campo & i)

                ' fechar parenteses
                clausulalike.Append(")")

                If Parametros Is Nothing Then Parametros = New AcessoDados.ParametroColecao

                Parametros.Add(New AcessoDados.Parametro With {.Campo = Campo & i, _
                                                               .Valor = "%" & Itens(i) & "%"})

            Next

            ' se adicionou IN deve fechar parenteses
            If Not addIn Then

                ' fechar comando do filtro
                clausulalike.Append(")")

            End If

        End If

        Return clausulalike.ToString

    End Function

    ''' <summary>
    ''' Dada uma query genérica, monta a cláusula where com os pârametros informados
    ''' </summary>
    ''' <param name="Comandos">Lista de comandos - Objeto Criterio{Codicional e Clausula. Ex: Condicional-> AND e Clausula-> CampoExemplo like '%valor%'}</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <hitory>pda 12/02/09 - Created</hitory>
    Public Shared Function MontarClausulaWhere(ByVal Comandos As CriterioColecion) As String

        Dim clausula As New StringBuilder

        For Each item As Criterio In Comandos

            If Not item.Clausula.Equals(String.Empty) AndAlso clausula.Length = 0 Then
                clausula.Append(" WHERE " & item.Clausula)
            ElseIf Not item.Clausula.Equals(String.Empty) AndAlso clausula.Length > 0 Then
                clausula.Append(" " & item.Condicional & " " & item.Clausula)
            End If
        Next

        Return clausula.ToString

    End Function

    Public Shared Function RetonarLocalRaiz() As String

        Return Environment.CurrentDirectory.Split("\").FirstOrDefault
    End Function

End Class
