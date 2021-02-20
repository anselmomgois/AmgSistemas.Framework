Public Class Consultar

    Public Shared Function BuscaCep(ByVal cep As String) As DadosCep

        Dim ds As DataSet
        Dim _resultado As String
        Dim DadosCep As DadosCep = Nothing
        Try
            ds = New DataSet()
            ds.ReadXml("http://cep.republicavirtual.com.br/web_cep.php?cep=" + cep.Replace("-", "").Trim() + "&formato=xml")
            If Not IsNothing(ds) Then
                If (ds.Tables(0).Rows.Count > 0) Then
                    _resultado = ds.Tables(0).Rows(0).Item("resultado").ToString()
                    DadosCep = New DadosCep
                    Select Case _resultado
                        Case "1"
                            DadosCep.UF = ds.Tables(0).Rows(0).Item("uf").ToString().Trim()
                            DadosCep.Cidade = ds.Tables(0).Rows(0).Item("cidade").ToString().Trim()
                            DadosCep.Bairro = ds.Tables(0).Rows(0).Item("bairro").ToString().Trim()
                            DadosCep.TipoLogradouro = ds.Tables(0).Rows(0).Item("tipo_logradouro").ToString().Trim()
                            DadosCep.Logradouro = ds.Tables(0).Rows(0).Item("logradouro").ToString().Trim()
                            DadosCep.Tipo = 1

                        Case "2"
                            DadosCep.UF = ds.Tables(0).Rows(0).Item("uf").ToString().Trim()
                            DadosCep.Cidade = ds.Tables(0).Rows(0).Item("cidade").ToString().Trim()
                            DadosCep.Tipo = 2
                        Case Else
                            DadosCep.Tipo = 0
                    End Select
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Falha ao Buscar o Cep" & vbCrLf & ex.ToString)
            Return Nothing
        End Try

        Return DadosCep
    End Function


End Class
