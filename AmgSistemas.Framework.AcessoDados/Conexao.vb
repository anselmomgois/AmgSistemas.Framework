Imports System.IO
Imports System.Xml.Serialization
Imports System.Configuration
Imports System.Security

Public Class Conexao

    Public Const ChaveCriptografia As String = "AMGSISTEMAS"
    Private Shared _PastaConexao As String = String.Empty
    Private Shared _LocalArquivo As String = String.Empty

    Public Shared Property LocalArquivo As String
        Get
            Return _LocalArquivo
        End Get
        Set(ByVal value As String)
            _LocalArquivo = value
        End Set
    End Property
    Public Shared Property PastaConexao As String
        Get
            Return _PastaConexao
        End Get
        Set(ByVal value As String)
            _PastaConexao = value
        End Set
    End Property

    Public Shared Function RetornarDadosConexao(ByVal Arquivo As String) As DadosConexao

        Dim ArquivoCompleto As String = Arquivo & ".xml"
        Dim objDadosConexao As DadosConexao = Nothing

        If Not File.Exists(ArquivoCompleto) Then
            ArquivoCompleto = RetornarDiretorio() & Arquivo & ".xml"
        End If

        If File.Exists(ArquivoCompleto) Then

            Dim w As StreamReader = Nothing

            Dim obj As New XmlSerializer(GetType(DadosConexao))
            w = New StreamReader(ArquivoCompleto)
            objDadosConexao = CType(obj.Deserialize(w), DadosConexao)
            w.Close()
            w = Nothing

            If objDadosConexao IsNot Nothing Then

                If Not String.IsNullOrEmpty(objDadosConexao.DataBase) Then
                    objDadosConexao.DataBase = AmgSistemas.Framework.Criptografia.Criptografar.DecryptString128Bit(objDadosConexao.DataBase, ChaveCriptografia)
                End If

                If Not String.IsNullOrEmpty(objDadosConexao.DataSource) Then
                    objDadosConexao.DataSource = AmgSistemas.Framework.Criptografia.Criptografar.DecryptString128Bit(objDadosConexao.DataSource, ChaveCriptografia)
                End If

                If Not String.IsNullOrEmpty(objDadosConexao.pwd) Then
                    objDadosConexao.pwd = AmgSistemas.Framework.Criptografia.Criptografar.DecryptString128Bit(objDadosConexao.pwd, ChaveCriptografia)
                End If

                If Not String.IsNullOrEmpty(objDadosConexao.Server) Then
                    objDadosConexao.Server = AmgSistemas.Framework.Criptografia.Criptografar.DecryptString128Bit(objDadosConexao.Server, ChaveCriptografia)
                End If

                If Not String.IsNullOrEmpty(objDadosConexao.UID) Then
                    objDadosConexao.UID = AmgSistemas.Framework.Criptografia.Criptografar.DecryptString128Bit(objDadosConexao.UID, ChaveCriptografia)
                End If


            End If

        End If

        Return objDadosConexao
    End Function


    Public Shared Function GerarStringConexaoDefault(ByVal NomeStringConexao As String, ByRef ConexaoOk As Boolean) As ConnectionStringSettings

        Dim objStringConexao As New ConnectionStringSettings

        ConexaoOk = False

        objStringConexao.ConnectionString = ConfigurationManager.AppSettings(NomeStringConexao)

        If String.IsNullOrEmpty(objStringConexao.ConnectionString) Then
            objStringConexao.ConnectionString = ConfigurationManager.AppSettings("STRINGDEFAULT")
        End If

        If Not String.IsNullOrEmpty(objStringConexao.ConnectionString) Then
            ConexaoOk = True
            objStringConexao.ProviderName = "System.Data.SqlClient"
        End If

        Return objStringConexao
    End Function

    Public Shared Function GerarStringConexao(ByVal strDatosSring As String) As ConnectionStringSettings

        Dim ConexaoOk As Boolean = False

        Dim StringConexao As ConnectionStringSettings = GerarStringConexaoDefault(strDatosSring, ConexaoOk)

        If Not ConexaoOk Then

            Dim objDadosConexao As DadosConexao = Nothing

            objDadosConexao = RetornarDadosConexao(strDatosSring)

            If objDadosConexao IsNot Nothing Then

                StringConexao = New ConnectionStringSettings

                If objDadosConexao.AutenticacaoIntegrada Then

                    If String.IsNullOrEmpty(objDadosConexao.DataBase) Then
                        StringConexao.ConnectionString = "Data Source=" & objDadosConexao.DataSource & ";Integrated Security=True"
                    Else
                        StringConexao.ConnectionString = "Data Source=" & objDadosConexao.DataSource & ";Initial Catalog=" & objDadosConexao.DataBase & ";Integrated Security=True"
                    End If

                Else
                    StringConexao.ConnectionString = "SERVER=" & objDadosConexao.Server & ";DATABASE=" & objDadosConexao.DataBase & ";UID=" & objDadosConexao.UID & ";pwd=" & objDadosConexao.pwd & ";"
                End If

                StringConexao.ProviderName = "System.Data.SqlClient"
            End If

        End If

        Return StringConexao
    End Function

    Public Shared Sub GerarArquivoConexao(ByVal objDadosConexao As DadosConexao, _
                                          ByVal Arquivo As String)

        If objDadosConexao IsNot Nothing Then

            If Not String.IsNullOrEmpty(objDadosConexao.DataBase) Then
                objDadosConexao.DataBase = AmgSistemas.Framework.Criptografia.Criptografar.EncryptString128Bit(objDadosConexao.DataBase, ChaveCriptografia)
            End If

            If Not String.IsNullOrEmpty(objDadosConexao.DataSource) Then
                objDadosConexao.DataSource = AmgSistemas.Framework.Criptografia.Criptografar.EncryptString128Bit(objDadosConexao.DataSource, ChaveCriptografia)
            End If

            If Not String.IsNullOrEmpty(objDadosConexao.pwd) Then
                objDadosConexao.pwd = AmgSistemas.Framework.Criptografia.Criptografar.EncryptString128Bit(objDadosConexao.pwd, ChaveCriptografia)
            End If

            If Not String.IsNullOrEmpty(objDadosConexao.Server) Then
                objDadosConexao.Server = AmgSistemas.Framework.Criptografia.Criptografar.EncryptString128Bit(objDadosConexao.Server, ChaveCriptografia)
            End If

            If Not String.IsNullOrEmpty(objDadosConexao.UID) Then
                objDadosConexao.UID = AmgSistemas.Framework.Criptografia.Criptografar.EncryptString128Bit(objDadosConexao.UID, ChaveCriptografia)
            End If

            Dim w As StreamWriter = Nothing
            Dim s As XmlSerializer

            'Serializa objeto e salva o XML
            Try

                w = New StreamWriter(RetornarDiretorio() & Arquivo & ".xml")
                s = New XmlSerializer(GetType(DadosConexao))
                s.Serialize(w, objDadosConexao)
            Catch ex As Exception
                Throw ex
            Finally
                If w IsNot Nothing Then
                    w.Close()
                    w = Nothing
                End If
            End Try

        End If

    End Sub

    Public Shared Function RetornarDiretorio() As String

        Dim LocalRaiz As String = RetonarLocalRaiz()

        If String.IsNullOrEmpty(_PastaConexao) AndAlso String.IsNullOrEmpty(_LocalArquivo) Then

            If Not System.IO.Directory.Exists(LocalRaiz & "\CE") Then
                System.IO.Directory.CreateDirectory(LocalRaiz & "\CE")
                System.IO.Directory.CreateDirectory(LocalRaiz & "\CE\DC")
            End If

            LocalRaiz &= "\CE\DC\"

        ElseIf Not String.IsNullOrEmpty(_LocalArquivo) Then
            LocalRaiz = _LocalArquivo
        Else

            If Not System.IO.Directory.Exists(LocalRaiz & "\" & _PastaConexao) Then
                System.IO.Directory.CreateDirectory(LocalRaiz & "\" & _PastaConexao)
                System.IO.Directory.CreateDirectory(LocalRaiz & "\" & _PastaConexao & "\DC")
            End If

            LocalRaiz &= "\" & _PastaConexao & "\DC\"
        End If

        Return LocalRaiz
    End Function

    Private Shared Function RetonarLocalRaiz() As String

        Return Environment.CurrentDirectory.Split("\").FirstOrDefault
    End Function

    Public Shared Function VerificarArquivoConexaoExiste(ByVal Arquivo As String) As Boolean

        Return File.Exists(RetornarDiretorio() & Arquivo & ".xml")
    End Function

    Public Shared Function TestarConexao(ByVal Arquivo As String, ByVal Tabela As String, _
                                         Optional ByRef objStringConexao As String = "") As Boolean

        Dim StringConexao As ConnectionStringSettings = GerarStringConexao(Arquivo)
        objStringConexao = StringConexao.ConnectionString

        Return Sql.TestarConexao(Tabela, StringConexao)
    End Function
End Class
