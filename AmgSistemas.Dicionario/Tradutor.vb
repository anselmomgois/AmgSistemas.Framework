Imports System.Configuration
Imports System.Globalization
Imports System.Web
Imports System.IO
Imports System.Text
Imports System.Collections.Specialized
Imports System.Text.RegularExpressions

''' -----------------------------------------------------------------------------
''' Project	 : Prosegur.Framework.Dicionario
''' Class	 : Tradutor
''' -----------------------------------------------------------------------------
''' <summary>
'''     Classe utilizada para efetuar traduções de texto
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[rafael.gans]		03/01/2008	Created
''' 	[carlos.bomtempo]	09/03/2009	Modificado para suportar múltiplos dicionarios
''' </history>
''' -----------------------------------------------------------------------------
Public Class Tradutor

#Region "  Constantes  "
    Private Const MSG_CHAVE_DUPLICADA As String = "A chave {0} já existe no dicionário. As chaves no dicionário devem ser únicas."
    Private Const MSG_DIR_DICIONARIOS As String = "É necessário informar o diretório de dicionários no AppSettings do arquivo de configuração. Ex: <add key=""DirDicionarios"" value=""Dicionarios"">"
    Private Const MSG_DIR_INEXISTENTE As String = "O diretório de dicionários informado não existe!"
#End Region

#Region "  Variáveis  "
    Private Shared _dirDicionario As String
    Private Shared _culturaPadrao As CultureInfo
    Private Shared _culturaAtiva As CultureInfo
    Private Shared _separadorChave As Char
    Private Shared _dicionarios As New Dictionary(Of String, StringDictionary)
#End Region

#Region "  Propriedades  "
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Diretório onde se encontram os dicionários
    ''' </summary>
    ''' <history>
    ''' 	[rafael.gans]		03/01/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Property DirDicionarios() As String
        Get
            If (_dirDicionario = String.Empty) Then
                _dirDicionario = ConfigurationManager.AppSettings("DirDicionarios")

                'Verifica se foi informado e se existe o diretório onde se encontram os dicionários
                If (_dirDicionario Is Nothing) Then
                    Throw New System.IO.DirectoryNotFoundException(MSG_DIR_DICIONARIOS)
                ElseIf Not (Directory.Exists(_dirDicionario)) Then
                    Dim appPath As String = String.Empty

                    'Recupera o diretório da aplicação para Web ou WindowsForm.
                    If (AplicacaoWeb) Then
                        appPath = HttpContext.Current.Request.MapPath(HttpContext.Current.Request.ApplicationPath)
                    Else
                        appPath = My.Application.Info.DirectoryPath
                    End If

                    If Not (appPath.EndsWith("\")) Then
                        appPath = appPath.Insert(appPath.Length, "\")
                    End If
                    If (_dirDicionario.StartsWith("/") OrElse _dirDicionario.StartsWith("\")) Then
                        _dirDicionario = _dirDicionario.Remove(0, 1)
                    End If
                    _dirDicionario = _dirDicionario.Insert(0, appPath)

                    'Se o diretório não existe levanta um erro
                    If Not (Directory.Exists(_dirDicionario)) Then
                        Throw New System.IO.DirectoryNotFoundException(MSG_DIR_INEXISTENTE)
                    End If
                End If
            End If

            Return _dirDicionario
        End Get
        Set(ByVal value As String)
            'Limpa o dicionário para poder recarregar com os novos parametros
            _dicionarios.Clear()
            _dirDicionario = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Caractere utilizado para separar a chave de seu valor no arquivo texto de dicionário
    ''' </summary>
    ''' <history>
    ''' 	[rafael.gans]	03/01/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Property SeparadorChave() As Char
        Get
            If (_separadorChave = Nothing) Then
                _separadorChave = "="
            End If

            Return _separadorChave
        End Get
        Set(ByVal value As Char)
            'Limpa o dicionário para poder recarregar com os novos parametros
            _dicionarios.Clear()
            _separadorChave = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cultura utilizada para a tradução caso não exista o dicionário para a cultura ativa
    ''' </summary>
    ''' <history>
    ''' 	[rafael.gans]	03/01/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Property CulturaPadrao() As CultureInfo
        Get
            If (_culturaPadrao Is Nothing) Then
                Dim cultura As String = ConfigurationManager.AppSettings("culturaDicionario")
                If cultura <> String.Empty Then
                    _culturaPadrao = New CultureInfo(cultura)
                Else
                    _culturaPadrao = New CultureInfo("pt-BR")
                End If
            End If
            Return _culturaPadrao
        End Get
        Set(ByVal value As CultureInfo)
            _culturaPadrao = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cultura utilizada para a tradução dos textos do sistema
    ''' </summary>
    ''' <history>
    ''' 	[rafael.gans]		03/01/2008	Created
    ''' 	[carlos.bomtempo]	09/03/2009	Corrigido o funcionamento da função que setava a cultura na variavel errada.
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <Obsolete("Esta propriedade não é mais utilizada e foi mantida por questões de compatibilidade.", False)> _
    Public Shared ReadOnly Property CulturaAtiva() As CultureInfo
        Get
            If (_culturaAtiva Is Nothing) Then
                Dim cultura As String = ConfigurationManager.AppSettings("culturaDicionario")
                If cultura <> String.Empty Then
                    _culturaAtiva = New CultureInfo(cultura)
                Else
                    _culturaAtiva = New CultureInfo("pt-BR")
                End If
            End If

            Return _culturaAtiva
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retorna true se a aplicação que está usando o dicionário é uma aplicação web, caso contrário retorna false
    ''' </summary>
    ''' <history>
    ''' 	[rafael.gans]	03/01/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Shared ReadOnly Property AplicacaoWeb() As Boolean
        Get
            Return (HttpContext.Current IsNot Nothing)
        End Get
    End Property
#End Region

#Region "  Métodos  "

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recupera os idiomas configurados do usuário
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[rafael.gans]	03/01/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Shared Function ObterIdiomas() As String()
        Dim culturas As String()

        If (AplicacaoWeb) Then
            culturas = HttpContext.Current.Request.UserLanguages
        Else
            culturas = New String() {CultureInfo.CurrentCulture.Name}
        End If

        If (culturas Is Nothing) Then
            culturas = New String() {}
        End If

        Return culturas
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Lê os arquivos textos de dicionários e cria uma coleção na memória para guarda-los.
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[rafael.gans]		03/01/2008	Created
    ''' 	[carlos.bomtempo]	09/03/2009	Modificado para tratar corretamente concorrencia na WEB e multiplos dicionarios
    ''' 	[carlos.bomtempo]	17/03/2009	Modificado para voltar a suportar culturas genéricas e multiplos arquivos para o mesmo dicionário
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Shared Sub CarregarDicionario()

        SyncLock _dicionarios

            If _dicionarios.Count = 0 Then

                Dim chave As String
                Dim valor As String
                Dim dicionario As StringDictionary
                Dim AplicarTrimValor As Boolean = True
                Dim cultura As CultureInfo = Nothing
                Dim validaCultura As New Regex("^((.+)(\.|\\))?([a-z]{2}(-[a-z]{2})?)\.txt$", RegexOptions.Compiled Or RegexOptions.IgnoreCase) 'valida as culturas para os seguintes exemplos de máscara: *.pt.txt pt.txt *.pt-br.txt pt-br.txt
                Dim resultadoCultura As Match = Nothing

                ' Obtém a parametrização para utilização do trim.
                If (Not (ConfigurationManager.AppSettings("DicionarioAplicaTrimValor") Is Nothing) _
                    AndAlso (ConfigurationManager.AppSettings("DicionarioAplicaTrimValor").ToString.Trim = "0")) Then
                    AplicarTrimValor = False
                End If

                ' Carrega em memória todos os dicionarios que obedeçam a máscara.
                For Each arquivo As String In Directory.GetFiles(DirDicionarios, "*.txt")
                    Try
                        'verifica a mascara do nome do arquivo.
                        resultadoCultura = validaCultura.Match(arquivo)
                        If resultadoCultura.Success Then
                            'verifica se é uma cultura válida.
                            cultura = New CultureInfo(resultadoCultura.Groups(4).Value) 'o grupo 4 representa a cultura informada
                        Else
                            Continue For
                        End If
                    Catch ex As Exception
                        Continue For
                    End Try

                    'Verifica se o dicionário da cultura já existe para continuar a popular o mesmo.
                    If _dicionarios.ContainsKey(cultura.Name) Then
                        dicionario = _dicionarios(cultura.Name)
                    Else
                        dicionario = New StringDictionary
                        _dicionarios.Add(cultura.Name, dicionario)
                    End If

                    'Lê cada linha do dicionário txt e cria uma coleção na memória
                    For Each s As String In File.ReadAllLines(arquivo, Encoding.Default)
                        If (Not s.StartsWith("//") AndAlso Not s.StartsWith("#") AndAlso s <> String.Empty AndAlso s.Contains(SeparadorChave)) Then
                            Dim indice As Integer = s.IndexOf(SeparadorChave)
                            chave = s.Substring(0, indice).Trim
                            valor = s.Substring(indice + 1)
                            ' Se for configurado para utilizar trim...
                            If AplicarTrimValor Then
                                valor = valor.Trim
                            End If
                            Try
                                dicionario.Add(chave, valor)
                            Catch ex As Exception
                                Throw New ArgumentException(String.Format(MSG_CHAVE_DUPLICADA, chave))
                            End Try
                        End If
                    Next
                Next

            End If

        End SyncLock

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recupera o texto referente a chave informada de acordo com a linguagem.
    ''' </summary>
    ''' <param name="chave">Chave do texto a ser recuperado</param>
    ''' <returns>Texto referente a chave informada</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[rafael.gans]		03/01/2008	Created
    ''' 	[carlos.bomtempo]	09/03/2009	Modificado para a criação das sobrecargas
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function Traduzir(ByVal chave As String) As String

        If (chave = String.Empty) Then
            Return String.Empty
        End If

        Dim traducao As String
        Dim encontrou As Boolean

        'Busca a tradução da chave en todos os dicionarios informados pelo cliente
        Dim idiomas As String() = ObterIdiomas()
        For Each idioma As String In idiomas
            traducao = Traduzir(chave, New CultureInfo(idioma.Split(";")(0)), True, encontrou)
            If encontrou Then
                Return traducao
            End If
        Next

        'Não foi encontrada nenhuma tradução nas culturas informadas pelo cliente. Então buscaremos na cultura padrão do sistema
        Return Traduzir(chave, Nothing, False, encontrou)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recupera o texto referente a chave informada de acordo com a linguagem.
    ''' </summary>
    ''' <param name="chave">Chave do texto a ser recuperado</param>
    ''' <param name="cultura">Cultura correspondente ao dicionário a ser utilizado</param>
    ''' <param name="chamadaInterna">Indica se a chamada foi ocasionada por uma sobrecarga</param>
    ''' <param name="encontrou">Indica se a chave foi encontrada no dicionario</param>
    ''' <returns>Texto referente a chave informada</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[rafael.gans]		03/01/2008	Created
    ''' 	[carlos.bomtempo]	09/03/2009	Modificado para a criação das sobrecargas
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Shared Function Traduzir(ByVal chave As String, ByVal cultura As CultureInfo, ByVal chamadaInterna As Boolean, ByRef encontrou As Boolean) As String

        If (chave = String.Empty) Then
            Return String.Empty
        End If

        CarregarDicionario()

        Dim retorno As String = String.Empty
        Dim dicionario As StringDictionary

        encontrou = False

        'Verifica se foi especificada alguma cultura
        If cultura IsNot Nothing Then
            'Verifica se o sistema possui o dicionario da cultura específica informada pelo cliente
            If _dicionarios.ContainsKey(cultura.Name) Then
                dicionario = _dicionarios(cultura.Name)
                'Busca a tradução da chave no dicionário
                If dicionario.ContainsKey(chave) Then
                    encontrou = True
                    retorno = dicionario(chave)
                Else
                    'Verifica na cultura genérica informada pelo cliente
                    If cultura.Name.Length > 2 Then
                        retorno = Traduzir(chave, New CultureInfo(cultura.TwoLetterISOLanguageName), True, encontrou)
                    End If
                End If
            Else 'Verifica se o sistema possui o dicionario da cultura genérica informada pelo cliente
                If cultura.Name.Length > 2 Then
                    retorno = Traduzir(chave, New CultureInfo(cultura.TwoLetterISOLanguageName), True, encontrou)
                End If
            End If
        End If

        'Somente deverá retornar a própria chave se quem iniciou a chamada foi a aplicação cliente e nenhuma tradução foi encontrada.
        If Not chamadaInterna And Not encontrou Then
            'Não foi encontrada a tradução em nenhum dicionário. Verificar no dicionario padrão.
            retorno = Traduzir(chave, CulturaPadrao, True, encontrou)

            'Nenhuma tradução encontrada. Retorna a própria chave entre colchetes.
            If Not encontrou Then
                retorno = "[" & chave & "]"
            End If
        End If

        Return retorno

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recupera o texto referente a chave informada de acordo com a linguagem.
    ''' </summary>
    ''' <param name="chave">Chave do texto a ser recuperado</param>
    ''' <param name="cultura">Cultura correspondente ao dicionário a ser utilizado</param>
    ''' <returns>Texto referente a chave informada</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[carlos.bomtempo]	03/01/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function Traduzir(ByVal chave As String, ByVal cultura As CultureInfo) As String

        Return Traduzir(chave, cultura, False, New Boolean)

    End Function

#End Region

End Class