Imports System.Text.RegularExpressions

Public Class Validacao


    Public Shared Function ValidarValorCampo(ByVal Texto As String, ByVal TipoValidacao As Enumeradores.TipoValidacao) As Boolean

        Dim strRegex As String = String.Empty


        Select Case TipoValidacao

            Case Enumeradores.TipoValidacao.EMAIL

                strRegex = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" + "\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" + ".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"

                Dim re As New Regex(strRegex)


                If (re.IsMatch(Texto)) Then
                    Return True
                End If

            Case Enumeradores.TipoValidacao.CPF

                Dim cpf As New CPF_CNPJ

                cpf.cpf = Texto

                Return cpf.isCpfValido()

            Case Enumeradores.TipoValidacao.CNPJ

                Dim cnpj As New CPF_CNPJ

                cnpj.cnpj = Texto

                Return cnpj.isCnpjValido()

            Case Enumeradores.TipoValidacao.EAN13

                Return ValidarCodigoBarrarEAN13(Texto)

            Case Enumeradores.TipoValidacao.CONEXAOINTERNET

                Return ValidaConexaoInternet(Texto)
        End Select


        Return False
    End Function

    Private Shared Function ValidarCodigoBarrarEAN13(ByVal Valor As String) As Boolean

        If String.IsNullOrEmpty(Valor.Trim) OrElse Valor.Trim.Length <> 13 OrElse Not IsNumeric(Valor) Then
            Return False
        End If

        Dim Digitos As Array = Valor.ToCharArray

        Dim SomaDigitosImpares As Integer = 0
        Dim SomaDigitosPares As Integer = 0
        Dim DigitoAtual As Integer = 0
        Dim SomaTotal As Integer = 0
        Dim UltimoDigito As Integer = Convert.ToInt32(Digitos(12).ToString)

        Dim Posicao As Integer = 0
        Dim DigitoControle As Integer = 0


        For i As Integer = 0 To 11

            Posicao = i + 1

            DigitoAtual = Convert.ToInt32(Digitos(i).ToString)

            ValidarNumeroParImpar(DigitoAtual, SomaDigitosImpares, SomaDigitosPares, Posicao)

        Next

        SomaDigitosPares *= 3

        SomaTotal = SomaDigitosImpares + SomaDigitosPares


        Dim indice As Integer = SomaTotal.ToString.Length

        DigitoControle = 10 - Convert.ToInt32(SomaTotal.ToString.Substring(indice - 1))


        If DigitoControle = 10 Then DigitoControle = 0



        If (UltimoDigito <> DigitoControle) Then
            Return False
        End If

        Return True
    End Function

    Private Shared Sub ValidarNumeroParImpar(ByVal Valor As Integer, ByRef SomaDigitosImpares As Integer, ByRef SomaDigitosPares As Integer, ByVal Indice As Integer)

        If Indice Mod 2 = 0 Then

            SomaDigitosPares += Valor
        Else

            SomaDigitosImpares += Valor
        End If
    End Sub

    Public Shared Function ValidaConexaoInternet(ByVal url As String) As Boolean

        'Define uma URL valida para consultar
        Dim HomePage As New System.Uri(url)

        'Monta a requisisão HTTP
        Dim req As System.Net.WebRequest
        req = System.Net.WebRequest.Create(HomePage)


        'Tenta fazer a requisisão
        Try

            Dim resp As System.Net.WebResponse
            resp = req.GetResponse()
            resp.Close()
            req = Nothing

            'Tudo certo... Temos conexão com a Internet
            Return True

        Catch
            req = Nothing
            'Não há conexão
            Return False
        End Try
    End Function

End Class
