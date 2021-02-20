Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Configuration

Public Class Sql

#Region "[CONSTRUTOR]"

    Public Sub New()

    End Sub

#End Region

#Region "[VARIAVEIS]"

    Private Parametros As ParametroColecao
    Private _comando As SqlCommand
    Private _conexao As SqlConnection
    Private _transacao As SqlTransaction
    Private _PastaArquivoConexao As String = String.Empty

#End Region

#Region "[PROPRIEDADES]"

    Public Property PastaArquivoConexao As String
        Get
            Return _PastaArquivoConexao
        End Get
        Set(value As String)
            _PastaArquivoConexao = value
        End Set
    End Property

#End Region

#Region "[METODOS]"


    Public Sub AdicionarParametro(ByVal Campo As String, ByVal Valor As Object, Optional ByVal Istanciar As Boolean = False,
                                  Optional ByVal Tipo As Nullable(Of SqlDbType) = Nothing, Optional Saida As Boolean = False, Optional Tamanho As Integer = 0)

        If Parametros Is Nothing OrElse Istanciar Then Parametros = New ParametroColecao

        Parametros.Add(New Parametro With { _
                       .Campo = "@" & Campo, _
                       .Valor = Valor, _
                       .Tipo = Tipo,
                       .Saida = Saida,
                       .Tamanho = Tamanho})

    End Sub

    Public Sub FecharConexao()

        If _comando IsNot Nothing Then
            _comando.Dispose()
            _comando = Nothing
        End If

        If _transacao IsNot Nothing Then
            _transacao.Rollback()
            _transacao.Dispose()
            _transacao = Nothing
        End If

        If _conexao IsNot Nothing Then

            _conexao.Close()
            _conexao.Dispose()
            _conexao = Nothing
        End If


    End Sub

    ''' <summary>
    ''' Retorna o data table
    ''' </summary>
    ''' <param name="Query"></param>
    ''' <param name="NomeStringconexao"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecutarDataTable(ByVal Query As String, ByVal NomeStringconexao As String,
                                      Optional TipoComando As CommandType = CommandType.Text) As DataTable

        Try


            Dim ds As DataSet
            Dim cn As SqlConnection
            Dim cmd As SqlCommand
            Dim da As SqlDataAdapter

            Dim objConnectionStringSettings As ConnectionStringSettings = Conexao.GerarStringConexao(NomeStringconexao)

            da = New SqlDataAdapter

            da.TableMappings.Add("table", "Tabela")
            cn = New SqlConnection(objConnectionStringSettings.ConnectionString)
            cn.Open()

            cmd = New SqlCommand(Query.ToString, cn)
            cmd.CommandType = TipoComando

            If Parametros IsNot Nothing AndAlso Parametros.Count > 0 Then

                For Each NomeParametro In (From p In Parametros Select p.Campo Distinct)

                    If Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).Count > 1 Then

                        Dim param As Parametro = Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).FirstOrDefault
                        Dim tabela As DataTable = New DataTable()
                        tabela.Columns.Add(New DataColumn("VALOR", GetType(String)))

                        For Each Parametro In Parametros.FindAll(Function(pa) pa.Campo = NomeParametro)
                            tabela.Rows.Add(Parametro.Valor)
                        Next

                        Dim tvparam As SqlParameter = cmd.Parameters.AddWithValue(param.Campo, tabela)

                        tvparam.SqlDbType = SqlDbType.Structured
                        tvparam.TypeName = param.TipoColecao
                    Else

                        Dim Parametro As Parametro = Parametros.Find(Function(pa) pa.Campo = NomeParametro)

                        cmd.Parameters.Add(New SqlParameter(Parametro.Campo, Parametro.Valor))

                        If TipoComando = CommandType.StoredProcedure Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Size = Parametro.Tamanho
                        End If

                        If Parametro.Saida Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Direction = ParameterDirection.Output
                        End If

                        If Parametro.Tipo IsNot Nothing Then
                            cmd.Parameters(cmd.Parameters.Count - 1).SqlDbType = Parametro.Tipo
                        End If

                    End If

                Next
                Parametros = Nothing

            End If


            da.SelectCommand = cmd

            ds = New DataSet("Tabela")
            da.Fill(ds, "Tabela")

            cn.Close()

            Return ds.Tables("Tabela")
        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Retorna o data table
    ''' </summary>
    ''' <param name="Query"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecutarDataTableArquivo(ByVal Query As String, ByVal ArquivoConexao As String,
                                      Optional TipoComando As CommandType = CommandType.Text) As DataTable

        Try


            Dim ds As DataSet
            Dim cn As SqlConnection
            Dim cmd As SqlCommand
            Dim da As SqlDataAdapter

            Conexao.PastaConexao = PastaArquivoConexao

            Dim objConnectionString As ConnectionStringSettings = Conexao.GerarStringConexao(ArquivoConexao)

            da = New SqlDataAdapter

            da.TableMappings.Add("table", "Tabela")
            cn = New SqlConnection(objConnectionString.ConnectionString)
            cn.Open()

            cmd = New SqlCommand(Query.ToString, cn)
            cmd.CommandType = TipoComando

            If Parametros IsNot Nothing AndAlso Parametros.Count > 0 Then

                For Each NomeParametro In (From p In Parametros Select p.Campo Distinct)

                    If Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).Count > 1 Then

                        Dim param As Parametro = Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).FirstOrDefault
                        Dim tabela As DataTable = New DataTable()
                        tabela.Columns.Add(New DataColumn("VALOR", GetType(String)))

                        For Each Parametro In Parametros.FindAll(Function(pa) pa.Campo = NomeParametro)
                            tabela.Rows.Add(Parametro.Valor)
                        Next

                        Dim tvparam As SqlParameter = cmd.Parameters.AddWithValue(param.Campo, tabela)

                        tvparam.SqlDbType = SqlDbType.Structured
                        tvparam.TypeName = param.TipoColecao
                    Else

                        Dim Parametro As Parametro = Parametros.Find(Function(pa) pa.Campo = NomeParametro)

                        cmd.Parameters.Add(New SqlParameter(Parametro.Campo, Parametro.Valor))

                        If TipoComando = CommandType.StoredProcedure Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Size = Parametro.Tamanho
                        End If

                        If Parametro.Saida Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Direction = ParameterDirection.Output
                        End If

                        If Parametro.Tipo IsNot Nothing Then
                            cmd.Parameters(cmd.Parameters.Count - 1).SqlDbType = Parametro.Tipo
                        End If

                    End If

                Next
                Parametros = Nothing

            End If


            da.SelectCommand = cmd

            ds = New DataSet("Tabela")
            da.Fill(ds, "Tabela")

            cn.Close()

            Return ds.Tables("Tabela")
        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Retorna o data table
    ''' </summary>
    ''' <param name="Query"></param>
    ''' <param name="NomeStringconexao"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecutarDataTables(ByVal Query As String, ByVal NomeStringconexao As String,
                                      Optional TipoComando As CommandType = CommandType.Text,
                                      Optional NomesTabelas As List(Of String) = Nothing) As List(Of DataTable)

        Try


            Dim ds As DataSet
            Dim cn As SqlConnection
            Dim cmd As SqlCommand
            Dim da As SqlDataAdapter

            Dim objConnectionStringSettings As ConnectionStringSettings = Conexao.GerarStringConexao(NomeStringconexao)

            da = New SqlDataAdapter

            da.TableMappings.Add("table", "Tabela")
            cn = New SqlConnection(objConnectionStringSettings.ConnectionString)
            cn.Open()

            cmd = New SqlCommand(Query.ToString, cn)
            cmd.CommandType = TipoComando

            If Parametros IsNot Nothing AndAlso Parametros.Count > 0 Then

                For Each NomeParametro In (From p In Parametros Select p.Campo Distinct)

                    If Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).Count > 1 Then

                        Dim param As Parametro = Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).FirstOrDefault
                        Dim tabela As DataTable = New DataTable()
                        tabela.Columns.Add(New DataColumn("VALOR", GetType(String)))

                        For Each Parametro In Parametros.FindAll(Function(pa) pa.Campo = NomeParametro)
                            tabela.Rows.Add(Parametro.Valor)
                        Next

                        Dim tvparam As SqlParameter = cmd.Parameters.AddWithValue(param.Campo, tabela)

                        tvparam.SqlDbType = SqlDbType.Structured
                        tvparam.TypeName = param.TipoColecao
                    Else

                        Dim Parametro As Parametro = Parametros.Find(Function(pa) pa.Campo = NomeParametro)

                        cmd.Parameters.Add(New SqlParameter(Parametro.Campo, Parametro.Valor))

                        If TipoComando = CommandType.StoredProcedure Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Size = Parametro.Tamanho
                        End If

                        If Parametro.Saida Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Direction = ParameterDirection.Output
                        End If

                        If Parametro.Tipo IsNot Nothing Then
                            cmd.Parameters(cmd.Parameters.Count - 1).SqlDbType = Parametro.Tipo
                        End If

                    End If

                Next
                Parametros = Nothing

            End If


            da.SelectCommand = cmd

            ds = New DataSet("Tabela")
            da.Fill(ds, "Tabela")

            cn.Close()

            Dim dts As New List(Of DataTable)

            If ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 Then

                Dim indice As Integer = 0

                For Each dt As DataTable In ds.Tables

                    If NomesTabelas IsNot Nothing AndAlso NomesTabelas.Count > 0 Then
                        dt.TableName = NomesTabelas(indice)
                    End If

                    dts.Add(dt)
                    indice += 1

                Next

            End If

            Return dts
        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Retorna o data table
    ''' </summary>
    ''' <param name="Query"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecutarDataTablesArquivo(ByVal Query As String, ByVal ArquivoConexao As String,
                                      Optional TipoComando As CommandType = CommandType.Text,
                                      Optional NomesTabelas As List(Of String) = Nothing) As List(Of DataTable)

        Try


            Dim ds As DataSet
            Dim cn As SqlConnection
            Dim cmd As SqlCommand
            Dim da As SqlDataAdapter

            Conexao.PastaConexao = PastaArquivoConexao

            Dim objConnectionString As ConnectionStringSettings = Conexao.GerarStringConexao(ArquivoConexao)

            da = New SqlDataAdapter

            da.TableMappings.Add("table", "Tabela")
            cn = New SqlConnection(objConnectionString.ConnectionString)
            cn.Open()

            cmd = New SqlCommand(Query.ToString, cn)
            cmd.CommandType = TipoComando

            If Parametros IsNot Nothing AndAlso Parametros.Count > 0 Then

                For Each NomeParametro In (From p In Parametros Select p.Campo Distinct)

                    If Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).Count > 1 Then

                        Dim param As Parametro = Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).FirstOrDefault
                        Dim tabela As DataTable = New DataTable()
                        tabela.Columns.Add(New DataColumn("VALOR", GetType(String)))

                        For Each Parametro In Parametros.FindAll(Function(pa) pa.Campo = NomeParametro)
                            tabela.Rows.Add(Parametro.Valor)
                        Next

                        Dim tvparam As SqlParameter = cmd.Parameters.AddWithValue(param.Campo, tabela)

                        tvparam.SqlDbType = SqlDbType.Structured
                        tvparam.TypeName = param.TipoColecao
                    Else

                        Dim Parametro As Parametro = Parametros.Find(Function(pa) pa.Campo = NomeParametro)

                        cmd.Parameters.Add(New SqlParameter(Parametro.Campo, Parametro.Valor))

                        If TipoComando = CommandType.StoredProcedure Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Size = Parametro.Tamanho
                        End If

                        If Parametro.Saida Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Direction = ParameterDirection.Output
                        End If

                        If Parametro.Tipo IsNot Nothing Then
                            cmd.Parameters(cmd.Parameters.Count - 1).SqlDbType = Parametro.Tipo
                        End If

                    End If

                Next
                Parametros = Nothing

            End If


            da.SelectCommand = cmd

            ds = New DataSet("Tabela")
            da.Fill(ds, "Tabela")

            cn.Close()

            Dim dts As New List(Of DataTable)

            If ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 Then

                Dim indice As Integer = 0

                For Each dt As DataTable In ds.Tables

                    If NomesTabelas IsNot Nothing AndAlso NomesTabelas.Count > 0 Then
                        dt.TableName = NomesTabelas(indice)
                    End If

                    dts.Add(dt)
                    indice += 1

                Next

            End If

            Return dts
        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Executar scalar
    ''' </summary>
    ''' <param name="Query"></param>
    ''' <param name="NomeStringconexao"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecutarScalar(ByVal Query As String, ByVal NomeStringconexao As String) As Object


        Dim cmd As SqlCommand
        Dim cn As SqlConnection
        Dim obj As Object
        Try


            Dim objConnectionString As ConnectionStringSettings = Conexao.GerarStringConexao(NomeStringconexao)

            cn = New SqlConnection(objConnectionString.ConnectionString)
            cn.Open()

            cmd = New SqlCommand(Query, cn)
            cmd.CommandType = CommandType.Text

            If Parametros IsNot Nothing AndAlso Parametros.Count > 0 Then

                For Each Parametro In Parametros
                    cmd.Parameters.Add(New SqlParameter(Parametro.Campo, Parametro.Valor))

                    If Parametro.Tipo IsNot Nothing Then
                        cmd.Parameters(cmd.Parameters.Count - 1).SqlDbType = Parametro.Tipo
                    End If

                Next

                Parametros = Nothing

            End If

            obj = cmd.ExecuteScalar
            cn.Close()

        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try

        Return obj
    End Function

    ''' <summary>
    ''' Executar scalar
    ''' </summary>
    ''' <param name="Query"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecutarScalarArquivo(ByVal Query As String, ByVal ArquivoConexao As String) As Object

        Dim cmd As SqlCommand
        Dim cn As SqlConnection = Nothing
        Dim obj As Object = Nothing
        Try


            Conexao.PastaConexao = PastaArquivoConexao

            Dim objConnectionString As ConnectionStringSettings = Conexao.GerarStringConexao(ArquivoConexao)

            cn = New SqlConnection(objConnectionString.ConnectionString)
            cn.Open()

            cmd = New SqlCommand(Query, cn)
            cmd.CommandType = CommandType.Text

            If Parametros IsNot Nothing AndAlso Parametros.Count > 0 Then

                For Each Parametro In Parametros
                    cmd.Parameters.Add(New SqlParameter(Parametro.Campo, Parametro.Valor))

                    If Parametro.Tipo IsNot Nothing Then
                        cmd.Parameters(cmd.Parameters.Count - 1).SqlDbType = Parametro.Tipo
                    End If

                Next

                Parametros = Nothing

            End If

            obj = cmd.ExecuteScalar
            cn.Close()

        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try

        Return obj
    End Function

    Public Sub ExecutarNonQuery(ByVal Query As String, ByVal NomeStringconexao As String,
                                      Optional TipoComando As CommandType = CommandType.Text)

        Dim cmd As SqlCommand
        Dim cn As SqlConnection = Nothing
        Try


            Dim objConnectionString As ConnectionStringSettings = Conexao.GerarStringConexao(NomeStringconexao)

            cn = New SqlConnection(objConnectionString.ConnectionString)
            cn.Open()

            cmd = New SqlCommand(Query, cn)
            cmd.CommandType = TipoComando

            If Parametros IsNot Nothing AndAlso Parametros.Count > 0 Then

                For Each NomeParametro In (From p In Parametros Select p.Campo Distinct)

                    If Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).Count > 1 Then

                        Dim param As Parametro = Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).FirstOrDefault
                        Dim tabela As DataTable = New DataTable()
                        tabela.Columns.Add(New DataColumn("VALOR", GetType(String)))

                        For Each Parametro In Parametros.FindAll(Function(pa) pa.Campo = NomeParametro)
                            tabela.Rows.Add(Parametro.Valor)
                        Next

                        Dim tvparam As SqlParameter = cmd.Parameters.AddWithValue(param.Campo, tabela)

                        tvparam.SqlDbType = SqlDbType.Structured
                        tvparam.TypeName = param.TipoColecao
                    Else

                        Dim Parametro As Parametro = Parametros.Find(Function(pa) pa.Campo = NomeParametro)

                        cmd.Parameters.Add(New SqlParameter(Parametro.Campo, Parametro.Valor))

                        If TipoComando = CommandType.StoredProcedure Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Size = Parametro.Tamanho
                        End If

                        If Parametro.Saida Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Direction = ParameterDirection.Output
                        End If

                        If Parametro.Tipo IsNot Nothing Then
                            cmd.Parameters(cmd.Parameters.Count - 1).SqlDbType = Parametro.Tipo
                        End If

                    End If

                Next
                Parametros = Nothing

            End If

            cmd.ExecuteNonQuery()
            cn.Close()

        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try
    End Sub

    Public Sub ExecutarNonQueryArquivo(ByVal Query As String, ByVal ArquivoConexao As String,
                                      Optional TipoComando As CommandType = CommandType.Text)
        Dim cmd As SqlCommand
        Dim cn As SqlConnection = Nothing

        Try


            Conexao.PastaConexao = PastaArquivoConexao

            Dim objConnectionString As ConnectionStringSettings = Conexao.GerarStringConexao(ArquivoConexao)

            cn = New SqlConnection(objConnectionString.ConnectionString)
            cn.Open()

            cmd = New SqlCommand(Query, cn)
            cmd.CommandType = TipoComando

            If Parametros IsNot Nothing AndAlso Parametros.Count > 0 Then

                For Each NomeParametro In (From p In Parametros Select p.Campo Distinct)

                    If Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).Count > 1 Then

                        Dim param As Parametro = Parametros.FindAll(Function(pa) pa.Campo = NomeParametro).FirstOrDefault
                        Dim tabela As DataTable = New DataTable()
                        tabela.Columns.Add(New DataColumn("VALOR", GetType(String)))

                        For Each Parametro In Parametros.FindAll(Function(pa) pa.Campo = NomeParametro)
                            tabela.Rows.Add(Parametro.Valor)
                        Next

                        Dim tvparam As SqlParameter = cmd.Parameters.AddWithValue(param.Campo, tabela)

                        tvparam.SqlDbType = SqlDbType.Structured
                        tvparam.TypeName = param.TipoColecao
                    Else

                        Dim Parametro As Parametro = Parametros.Find(Function(pa) pa.Campo = NomeParametro)

                        cmd.Parameters.Add(New SqlParameter(Parametro.Campo, Parametro.Valor))

                        If TipoComando = CommandType.StoredProcedure Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Size = Parametro.Tamanho
                        End If

                        If Parametro.Saida Then
                            cmd.Parameters(cmd.Parameters.Count - 1).Direction = ParameterDirection.Output
                        End If

                        If Parametro.Tipo IsNot Nothing Then
                            cmd.Parameters(cmd.Parameters.Count - 1).SqlDbType = Parametro.Tipo
                        End If

                    End If

                Next
                Parametros = Nothing

            End If

            cmd.ExecuteNonQuery()
            cn.Close()

        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' Inicia a transação
    ''' </summary>
    ''' <param name="NomeConexaoString"></param>
    ''' <remarks></remarks>
    Public Sub IniciarTransacao(ByVal NomeConexaoString As String)

        Dim objConnectionStringSettings As ConnectionStringSettings = Conexao.GerarStringConexao(NomeConexaoString)

        Try


            _conexao = New SqlConnection(objConnectionStringSettings.ConnectionString)

            _conexao.Open()

            'Inicia a Transação 
            _transacao = _conexao.BeginTransaction

        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try

    End Sub

    ''' <summary>
    ''' Inicia a transação
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub IniciarTransacaoArquivo(ByVal ArquivoConexao As String)

        Conexao.PastaConexao = PastaArquivoConexao
        Try


            Dim objConnectionString As ConnectionStringSettings = Conexao.GerarStringConexao(ArquivoConexao)

            _conexao = New SqlConnection(objConnectionString.ConnectionString)

            _conexao.Open()

            'Inicia a Transação 
            _transacao = _conexao.BeginTransaction

         Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Adiciona Transação
    ''' </summary>
    ''' <param name="Query"></param>
    ''' <remarks></remarks>
    Public Sub AdicionarTransacao(ByVal Query As String)

        Try

            _comando = (New SqlCommand(Query, _conexao, _transacao))
            _comando.CommandType = CommandType.Text

            If Parametros IsNot Nothing AndAlso Parametros.Count > 0 Then

                For Each Parametro In Parametros
                    _comando.Parameters.Add(New SqlParameter(Parametro.Campo, Parametro.Valor))

                    If Parametro.Tipo IsNot Nothing Then
                        _comando.Parameters(_comando.Parameters.Count - 1).DbType = Parametro.Tipo
                    End If

                Next

                Parametros = Nothing

            End If


            _comando.ExecuteNonQuery()
        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Executa a transação
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ExecutarTransacao()

        Try
            _transacao.Commit()
            _conexao.Close()

        Catch ex As Exception
            FecharConexao()
            Throw ex
        End Try

    End Sub


    ''' <summary>
    ''' Executar scalar
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function TestarConexao(ByVal Tabela As String, ByVal objConnectionString As ConnectionStringSettings) As Object

        Dim cmd As SqlCommand
        Dim cn As SqlConnection
        Dim obj As Object = Nothing
        Try


            cn = New SqlConnection(objConnectionString.ConnectionString)
            cn.Open()

            cmd = New SqlCommand("SELECT 1 FROM " & Tabela, cn)
            cmd.CommandType = CommandType.Text

            obj = cmd.ExecuteScalar
            cn.Close()

        Catch ex As Exception
            If cmd IsNot Nothing Then cmd.Dispose()
            If cn IsNot Nothing Then cn.Close()
        End Try

        Return obj
    End Function
#End Region

End Class
