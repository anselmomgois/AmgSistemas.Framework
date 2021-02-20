Imports System.Net.Mail

Public Class Email

    Public Shared Sub enviaMensagemEmail(ByVal para As List(Of String), ByVal de As String, ByVal bcc As List(Of String),
                                         ByVal cc As List(Of String), ByVal subject As String, ByVal body As String, ByVal EnderecoSmtp As String,
                                         ByVal Prioridade As MailPriority, ByVal Porta As Nullable(Of Integer), ByVal HabiltarSsl As Boolean, ByVal Senha As String, Optional ByVal UseDefaultCredentials As Boolean = False)


        Dim mMailMessage As New MailMessage()

        mMailMessage.From = New MailAddress(de)

        For Each item In para
            mMailMessage.To.Add(New MailAddress(item))
        Next


        If bcc IsNot Nothing AndAlso bcc.Count > 0 Then

            For Each item In bcc
                ' Define o endereço bcc
                mMailMessage.Bcc.Add(New MailAddress(item))
            Next

        End If

        ' verifica se o valor para cc é nulo ou uma string vazia
        If cc IsNot Nothing AndAlso cc.Count > 0 Then

            For Each item In cc
                ' Define o endereço cc 
                mMailMessage.CC.Add(New MailAddress(item))
            Next

        End If

        mMailMessage.Subject = subject
        ' Define o corpo da mensagem
        mMailMessage.Body = body

        ' Define o formato do email como HTML
        mMailMessage.IsBodyHtml = True
        ' Define a prioridade da mensagem como normal
        mMailMessage.Priority = Prioridade

        Dim mSmtpClient As SmtpClient = Nothing

        If Porta IsNot Nothing Then
            mSmtpClient = New SmtpClient(EnderecoSmtp, Porta)
        Else
            mSmtpClient = New SmtpClient(EnderecoSmtp)
        End If

        mSmtpClient.UseDefaultCredentials = UseDefaultCredentials
        mSmtpClient.EnableSsl = HabiltarSsl
        mSmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network
        mSmtpClient.Credentials = New Net.NetworkCredential(de, Senha)

        ' Envia o email
        mSmtpClient.Send(mMailMessage)


    End Sub

    Public Shared Sub enviaMensagemEmail(ByVal para As List(Of String), ByVal de As String, ByVal bcc As List(Of String),
                                       ByVal cc As List(Of String), ByVal subject As String, ByVal body As String, ByVal EnderecoSmtp As String,
                                       ByVal Prioridade As MailPriority, ByVal Porta As Nullable(Of Integer), ByVal HabiltarSsl As Boolean, ByVal Senha As String,
                                       ByVal Anexo As String, Optional ByVal UseDefaultCredentials As Boolean = False)


        Dim mMailMessage As New MailMessage()

        mMailMessage.From = New MailAddress(de)

        For Each item In para
            mMailMessage.To.Add(New MailAddress(item))
        Next


        If bcc IsNot Nothing AndAlso bcc.Count > 0 Then

            For Each item In bcc
                ' Define o endereço bcc
                mMailMessage.Bcc.Add(New MailAddress(item))
            Next

        End If

        ' verifica se o valor para cc é nulo ou uma string vazia
        If cc IsNot Nothing AndAlso cc.Count > 0 Then

            For Each item In cc
                ' Define o endereço cc 
                mMailMessage.CC.Add(New MailAddress(item))
            Next

        End If

        mMailMessage.Subject = subject
        ' Define o corpo da mensagem
        mMailMessage.Body = body

        ' Define o formato do email como HTML
        mMailMessage.IsBodyHtml = True
        ' Define a prioridade da mensagem como normal
        mMailMessage.Priority = Prioridade

        If Not String.IsNullOrEmpty(Anexo) Then

            mMailMessage.Attachments.Add(New Attachment(Anexo))

        End If

        Dim mSmtpClient As SmtpClient = Nothing

        If Porta IsNot Nothing Then
            mSmtpClient = New SmtpClient(EnderecoSmtp, Porta)
        Else
            mSmtpClient = New SmtpClient(EnderecoSmtp)
        End If

        mSmtpClient.EnableSsl = HabiltarSsl
        mSmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network
        mSmtpClient.Credentials = New Net.NetworkCredential(de, Senha)
        mSmtpClient.UseDefaultCredentials = UseDefaultCredentials
        ' Envia o email
        mSmtpClient.Send(mMailMessage)


    End Sub

    Public Shared Sub enviaMensagemEmail(ByVal para As List(Of String), ByVal de As String, ByVal bcc As List(Of String),
                                         ByVal cc As List(Of String), ByVal subject As String, ByVal body As String, ByVal EnderecoSmtp As String,
                                         ByVal Prioridade As MailPriority, ByVal Porta As Nullable(Of Integer), ByVal HabiltarSsl As Boolean, ByVal Senha As String,
                                         ByVal Anexo As System.IO.Stream, ByVal TipoArquivoAnexo As Net.Mime.ContentType, Optional ByVal UseDefaultCredentials As Boolean = False)


        Dim mMailMessage As New MailMessage()

        mMailMessage.From = New MailAddress(de)

        For Each item In para
            mMailMessage.To.Add(New MailAddress(item))
        Next


        If bcc IsNot Nothing AndAlso bcc.Count > 0 Then

            For Each item In bcc
                ' Define o endereço bcc
                mMailMessage.Bcc.Add(New MailAddress(item))
            Next

        End If

        ' verifica se o valor para cc é nulo ou uma string vazia
        If cc IsNot Nothing AndAlso cc.Count > 0 Then

            For Each item In cc
                ' Define o endereço cc 
                mMailMessage.CC.Add(New MailAddress(item))
            Next

        End If

        mMailMessage.Subject = subject
        ' Define o corpo da mensagem
        mMailMessage.Body = body

        ' Define o formato do email como HTML
        mMailMessage.IsBodyHtml = True
        ' Define a prioridade da mensagem como normal
        mMailMessage.Priority = Prioridade
        mMailMessage.BodyEncoding = System.Text.Encoding.UTF8

        If Anexo IsNot Nothing AndAlso TipoArquivoAnexo IsNot Nothing Then

            mMailMessage.Attachments.Add(New Attachment(Anexo, TipoArquivoAnexo))
            mMailMessage.Attachments(0).ContentDisposition.FileName = "ReportFechamentoCaixa.Pdf"

        End If

        Dim mSmtpClient As SmtpClient = Nothing

        If Porta IsNot Nothing Then
            mSmtpClient = New SmtpClient(EnderecoSmtp, Porta)
        Else
            mSmtpClient = New SmtpClient(EnderecoSmtp)
        End If

        mSmtpClient.EnableSsl = HabiltarSsl

        mSmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network
        mSmtpClient.Credentials = New Net.NetworkCredential(de, Senha)
        mSmtpClient.UseDefaultCredentials = UseDefaultCredentials
        ' Envia o email
        mSmtpClient.Send(mMailMessage)


    End Sub
End Class
