Imports System.IO
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports System.Web.Script.Serialization

Public Class ServiceMail
    Public Shared ReadOnly smtpLogin As String = ini.ReadValue("MAIL", "userMail")
    Private Shared ReadOnly smtpPassword As String = ini.ReadValue("MAIL", "userMailPassword_encrypted")
    Private Shared ReadOnly smtpServer As String = ini.ReadValue("MAIL", "SMTPServer")
    Private Shared ReadOnly QueuePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MailQueue.json")
    Private Shared ReadOnly QueueLock As New Object()
    Private Shared ReadOnly QueueEncoding As New System.Text.UTF8Encoding(True)

    Private Class MailEnAttente
        Public Property Sender As String
        Public Property Recipient As String
        Public Property Subject As String
        Public Property Body As String
        Public Property AttachmentPath As String
        Public Property NombreEchecs As Integer
        Public Property DernierEchec As DateTime
    End Class

    Public Shared Sub SendEmail(ByVal sender As String, ByVal recipient As String, ByVal subject As String, ByVal body As String, Optional ByVal attachmentString As String = "")
        Dim mail As New MailEnAttente With {
            .Sender = sender,
            .Recipient = recipient,
            .Subject = subject,
            .Body = body,
            .AttachmentPath = attachmentString
        }

        Dim erreur As String = Nothing
        If EssayerEnvoyerMail(mail, erreur) Then
            Exit Sub
        End If

        mail.NombreEchecs = 1
        mail.DernierEchec = Now
        AjouterMailEnAttente(mail)
        Commun.Journal("ERREUR : echec de l'envoi de mail sur le port 587 avec TLS : " & erreur & ". Mail place en attente.")
    End Sub

    Public Shared Sub TraiterMailsEnAttente()
        SyncLock QueueLock
            Dim mails As List(Of MailEnAttente) = ChargerMailsEnAttente()
            If mails Is Nothing OrElse mails.Count = 0 Then
                Exit Sub
            End If

            Commun.Journal("Reprise des mails en attente : " & mails.Count.ToString())
            Dim mailsRestants As New List(Of MailEnAttente)()

            For Each mail As MailEnAttente In mails
                Dim erreur As String = Nothing
                If EssayerEnvoyerMail(mail, erreur) Then
                    Commun.Journal("Mail en attente envoye : " & mail.Subject)
                Else
                    mail.NombreEchecs += 1
                    mail.DernierEchec = Now
                    mailsRestants.Add(mail)
                    Commun.Journal("ERREUR : nouvel echec d'un mail en attente : " & erreur & " : " & mail.Subject)
                End If
            Next

            EcrireMailsEnAttente(mailsRestants)
        End SyncLock
    End Sub

    Private Shared Function EssayerEnvoyerMail(ByVal mail As MailEnAttente, ByRef erreur As String) As Boolean
        Try
            Using message As New MailMessage()
                Dim senderAddress As String = Trim(mail.Sender)
                Dim replyToAddress As String = ""

                If InStr(senderAddress, ";") > 0 Then
                    Dim tabSender As String() = Split(senderAddress, ";")
                    senderAddress = Trim(tabSender(0))
                    For i = 1 To UBound(tabSender)
                        Dim senderPart As String = Trim(tabSender(i))
                        If LCase(Left(senderPart, 8)) = "replyto:" Then
                            replyToAddress = Trim(Mid(senderPart, 9))
                        End If
                    Next
                End If

                message.From = New MailAddress(senderAddress)

                If replyToAddress <> "" Then
                    message.ReplyToList.Add(New MailAddress(replyToAddress))
                End If

                Dim tabDest As String() = Split(mail.Recipient, ";")
                For i = 0 To UBound(tabDest)
                    Dim dest As String = Trim(tabDest(i))
                    If dest <> "" Then
                        If LCase(Left(dest, 3)) = "cc:" Then
                            message.CC.Add(Trim(Mid(dest, 4)))
                        ElseIf LCase(Left(dest, 4)) = "bcc:" Then
                            message.Bcc.Add(Trim(Mid(dest, 5)))
                        Else
                            message.To.Add(dest)
                        End If
                    End If
                Next

                message.Subject = mail.Subject

                If Left(mail.Body, 21) = "<!DOCTYPE HTML PUBLIC" OrElse InStr(LCase(mail.Body), "</") > 0 Then
                    message.IsBodyHtml = True
                    message.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(RemoveHtml(mail.Body), Nothing, "text/plain"))
                    message.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(mail.Body, Nothing, "text/html"))
                Else
                    message.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(mail.Body, Nothing, "text/plain"))
                End If

                If mail.AttachmentPath <> "" Then
                    message.Attachments.Add(New Attachment(mail.AttachmentPath))
                End If


                Using mailSender As New SmtpClient(smtpServer, 587)
                    With mailSender
                        .UseDefaultCredentials = False
                        .Credentials = New Net.NetworkCredential(smtpLogin, smtpPassword, "IGBMC")
                        .EnableSsl = True
                    End With

                    mailSender.Send(message)
                End Using
            End Using

            Return True
        Catch ex As Exception
            erreur = ex.Message
            Return False
        End Try
    End Function

    Private Shared Sub AjouterMailEnAttente(ByVal mail As MailEnAttente)
        SyncLock QueueLock
            Dim mails As List(Of MailEnAttente) = ChargerMailsEnAttente()
            If mails Is Nothing Then
                Exit Sub
            End If

            mails.Add(mail)
            EcrireMailsEnAttente(mails)
        End SyncLock
    End Sub

    Private Shared Function ChargerMailsEnAttente() As List(Of MailEnAttente)
        If Not File.Exists(QueuePath) Then
            Return New List(Of MailEnAttente)()
        End If

        Try
            Dim mails As List(Of MailEnAttente) = New JavaScriptSerializer().Deserialize(Of List(Of MailEnAttente))(File.ReadAllText(QueuePath, QueueEncoding))
            Return If(mails, New List(Of MailEnAttente)())
        Catch ex As Exception
            Commun.Journal("ERREUR : lecture de la file de mails en attente : " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Shared Sub EcrireMailsEnAttente(ByVal mails As List(Of MailEnAttente))
        Try
            If mails.Count = 0 Then
                If File.Exists(QueuePath) Then
                    File.Delete(QueuePath)
                End If
                Exit Sub
            End If

            File.WriteAllText(QueuePath, New JavaScriptSerializer().Serialize(mails), QueueEncoding)
        Catch ex As Exception
            Commun.Journal("ERREUR : ecriture de la file de mails en attente : " & ex.Message)
        End Try
    End Sub

    Private Shared Function RemoveHtml(ByVal html As String) As String
        Return Regex.Replace(html, "<.*?>", "")
    End Function
End Class