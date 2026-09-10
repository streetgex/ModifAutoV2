Imports System.IO
Imports System.Net.Mail
Imports System.DirectoryServices
Imports System.Text.RegularExpressions

Public Class Commun
    Public Shared tabExcepUser As String()
    Public Shared journalECHECMail As String = ""
    Public Shared controlSendMail As Boolean = False
    Public Shared fichierLog As String = "c:\temp\Log" & application.productname & ".log"
    Public Shared tabLoginAliasID As String()
    Public Shared tabLoginAlias As String()
    Public Shared iniFilePath As String = "\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\ScriptStephV2.ini"
    Public Shared ini As New IniFile(iniFilePath)
    Public Shared smtpLogin As String = ini.ReadValue("MAIL", "userMail")
    Public Shared smtpPassword As String = ini.ReadValue("MAIL", "userMailPassword_encrypted")
    'Public Shared mailSenderAddress As String = ini.ReadValue("GLOBAL", "MailSenderAddress")
    Public Shared auth As AuthenticationTypes = AuthenticationTypes.Secure 'AuthenticationTypes.Secure Or AuthenticationTypes.SecureSocketsLayer

    Private Shared _barreEnCours As Boolean = False
    Private Shared _longueurDerniereBarre As Integer = 0
    Private Shared ReadOnly JournalEncoding As New System.Text.UTF8Encoding(True)

    Shared Sub New()
        Console.OutputEncoding = System.Text.Encoding.UTF8
    End Sub

    Private Shared Sub AjouterAuJournal(texte As String)
        If File.Exists(fichierLog) Then
            File.AppendAllText(fichierLog, texte, JournalEncoding)
        Else
            File.WriteAllText(fichierLog, texte, JournalEncoding)
        End If
    End Sub


    Friend Shared DCName As String = ""
    ''' <summary>
    ''' Retourne le préfixe du serveur LDAP basé sur la valeur de <c>DCName</c>.
    ''' </summary>
    ''' <returns>
    ''' Une chaîne au format <c>NomServeur/</c> lorsque <c>DCName</c> est renseigné,
    ''' sinon une chaîne vide.
    ''' </returns>
    ''' <remarks>
    ''' Cette méthode est utilisée pour construire des chemins LDAP relatifs au contrôleur
    ''' de domaine ou serveur d'annuaire courant.
    ''' </remarks>
    ''' <example>
    ''' Exemple avec <c>DCName = "dc01.contoso.local"</c> :
    ''' <code>
    ''' Dim prefix As String = LdapServerPrefix()
    ''' ' Résultat : "dc01.contoso.local/"
    ''' </code>
    ''' </example>
    Public Shared Function LdapServerPrefix() As String
        If String.IsNullOrWhiteSpace(DCName) Then
            Return ""
        Else
            Return DCName & "/"
        End If
    End Function

    ''' <summary>
    ''' Construit un chemin LDAP relatif à partir d'un distinguished name (DN).
    ''' </summary>
    ''' <param name="dn">
    ''' Le distinguished name à concaténer au préfixe serveur LDAP.
    ''' </param>
    ''' <returns>
    ''' Le chemin LDAP relatif au format <c>serveur/dn</c>, ou uniquement <c>dn</c>
    ''' si aucun serveur n'est défini.
    ''' </returns>
    ''' <remarks>
    ''' Cette méthode ne préfixe pas le résultat par <c>LDAP://</c>.
    ''' Pour obtenir un chemin LDAP complet, utiliser <see cref="LdapPrefix"/>.
    ''' </remarks>
    ''' <example>
    ''' Exemple :
    ''' <code>
    ''' Dim path As String = LdapPath("OU=Users,DC=contoso,DC=local")
    ''' ' Résultat : "dc01.contoso.local/OU=Users,DC=contoso,DC=local"
    ''' </code>
    ''' </example>
    Public Shared Function LdapPath(ByVal dn As String) As String
        Return LdapServerPrefix() & dn
    End Function

    ''' <summary>
    ''' Retourne le préfixe LDAP complet.
    ''' </summary>
    ''' <returns>
    ''' Une chaîne au format <c>LDAP://serveur/</c> lorsque <c>DCName</c> est renseigné,
    ''' sinon <c>LDAP://</c>.
    ''' </returns>
    ''' <remarks>
    ''' Cette méthode sert de base pour construire des chemins LDAP complets vers des objets
    ''' de l'annuaire, comme <c>rootDSE</c> ou un DN spécifique.
    ''' </remarks>
    ''' <example>
    ''' Exemple avec <c>DCName = "dc01.contoso.local"</c> :
    ''' <code>
    ''' Dim prefix As String = LdapPrefix()
    ''' ' Résultat : "LDAP://dc01.contoso.local/"
    ''' </code>
    ''' </example>
    Public Shared Function LdapPrefix() As String
        Return "LDAP://" & LdapServerPrefix()
    End Function

    ''' <summary>
    ''' Retourne le chemin LDAP complet vers l'entrée <c>rootDSE</c>.n ldappath
    ''' </summary>
    ''' <returns>
    ''' Une chaîne au format <c>LDAP://serveur/rootDSE</c>, ou <c>LDAP://rootDSE</c>
    ''' si aucun serveur n'est défini.
    ''' </returns>
    ''' <remarks>
    ''' <c>rootDSE</c> est une entrée spéciale de l'annuaire permettant de récupérer
    ''' des informations sur le serveur LDAP et ses capacités.
    ''' </remarks>
    ''' <example>
    ''' Exemple avec <c>DCName = "dc01.contoso.local"</c> :
    ''' <code>
    ''' Dim rootDse As String = LdapRootDsePath()
    ''' ' Résultat : "LDAP://dc01.contoso.local/rootDSE"
    ''' </code>
    ''' </example>
    Public Shared Function LdapRootDsePath() As String
        Return LdapPrefix() & "rootDSE"
    End Function

    ''' <summary>
    ''' Instription dans les fichiers Log de l'appli
    ''' </summary>
    ''' <param name="texte">Texte à inscrire dans le fichier log.</param>
    ''' <param name="envoyerMail">Optional: Envoyer un mail sur cette entrée du fichier journal.</param>
    ''' <remarks>Si 'texte' commence par ERREUR, le texte est décalé dans le fichier journal</remarks>
    Shared Sub Journal(ByVal texte As String, Optional ByVal envoyerMail As Boolean = False)

        If envoyerMail = True Then
            controlSendMail = True
            journalECHECMail = journalECHECMail & vbCrLf & Now & " : " & texte
        End If

        Dim ctrlError As Boolean = False
        If texte.TrimStart().StartsWith("ERREUR", StringComparison.Ordinal) Then
            If Not texte.StartsWith(vbTab) Then texte = vbTab & texte
            ctrlError = True
        End If

        Dim separateur As String = "__________________________________________________________________________________________________"

        If texte = "Debut de traitement" Then
            AjouterAuJournal(separateur & vbCrLf & vbCrLf)
        End If

        Dim color As ConsoleColor = ConsoleColor.Gray
        If ctrlError = True Then color = ConsoleColor.Red

        ConsoleWrite(texte, color)

        AjouterAuJournal(Now & " : " & texte & vbCrLf)

        If texte Like "Fin de traitement*" Then
            AjouterAuJournal(separateur & vbCrLf & vbCrLf)
        End If

    End Sub
    Shared Sub ConsoleWrite(ByVal texte As String, ByVal couleur As ConsoleColor)
        'voir les couleurs ici : https://msdn.microsoft.com/fr-fr/library/system.consolecolor(v=vs.110).aspx
        Console.ForegroundColor = couleur
        Console.WriteLine(texte)
        Console.ResetColor()
    End Sub

    Shared Sub EcrireMembreDe(ByVal group As DirectoryEntry, ByVal membresActuels As Object(), ByVal membresAEcrire As Object())
        Try
            Dim membresExplicites As Object() = MembresDuGroupe(group.Properties("sAMAccountName").Value, False)
            If membresExplicites Is Nothing Then membresExplicites = New Object() {}

            Dim membresCibles As Object()
            If membresAEcrire Is Nothing Then
                membresCibles = New Object() {}
            Else
                membresCibles = membresAEcrire
            End If

            Dim membresPrimaires As Object() = New Object() {}
            If membresActuels IsNot Nothing Then
                membresPrimaires = membresActuels.Except(membresExplicites).ToArray()
            End If

            ' Les membres primaires ne sont pas stockes dans l'attribut LDAP "member".
            If membresPrimaires.Length > 0 Then
                Dim membresPrimairesHorsFiltre As Object() = membresPrimaires.Except(membresCibles).ToArray()
                If membresPrimairesHorsFiltre.Length > 0 Then
                    Commun.Journal("AVERTISSEMENT : membres primaires conserves dans le groupe """ &
                                   group.Properties("cn").Value.ToString & """ :" & vbCrLf &
                                   Join(membresPrimairesHorsFiltre, vbCrLf))
                End If

                membresCibles = membresCibles.Except(membresPrimaires).ToArray()
            End If

            Dim membresAAjouter As Object() = membresCibles.Except(membresExplicites).ToArray()
            Dim membresARetirer As Object() = membresExplicites.Except(membresCibles).ToArray()

            For Each membre As Object In membresAAjouter
                Try
                    group.Properties("member").Add(membre)
                    group.CommitChanges()
                Catch ex As DirectoryServices.DirectoryServicesCOMException
                    Commun.Journal("ERREUR LDAP lors de l'ajout de """ & membre.ToString & """ au groupe """ &
                                   group.Properties("cn").Value.ToString & """ : " & ex.Message &
                                   " - HResult=0x" & ex.ErrorCode.ToString("X8"))
                    group.RefreshCache()
                End Try
            Next

            For Each membre As Object In membresARetirer
                Try
                    group.Properties("member").Remove(membre)
                    group.CommitChanges()
                Catch ex As DirectoryServices.DirectoryServicesCOMException
                    Commun.Journal("ERREUR LDAP lors du retrait de """ & membre.ToString & """ du groupe """ &
                                   group.Properties("cn").Value.ToString & """ : " & ex.Message &
                                   " - HResult=0x" & ex.ErrorCode.ToString("X8"))
                    group.RefreshCache()
                End Try
            Next
        Catch ex As DirectoryServices.DirectoryServicesCOMException
            Commun.Journal(
          "ERREUR LDAP lors de la mise à jour du groupe """ &
          group.Properties("cn").Value.ToString &
          """ (" & group.Path & ") : " &
          ex.Message & " - HResult=0x" & ex.ErrorCode.ToString("X8"))

        Catch ex As Exception
            Commun.Journal(
          "ERREUR lors de la mise à jour du groupe """ &
          group.Properties("cn").Value.ToString &
          """ (" & group.Path & ") : " &
          ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' Creation d'un nouveau groupe
    ''' </summary>
    ''' <param name="OUCreation">OU ou le nouveau groupe doit etre créé.</param>
    ''' <param name="samaNewGroup">Nom du nouveau groupe.</param>
    ''' <param name="description">Description du nouveau groupe.</param>
    ''' <param name="SAMmanager">login du manager du groupe ou de l'equipe (optional).</param>
    ''' <remarks></remarks>
    Shared Sub NouveauGroupe(ByVal OUCreation As DirectoryEntry, ByVal samaNewGroup As String, ByVal description As String, Optional ByVal SAMmanager As String = "?")
        Dim nouveauGroupe As DirectoryEntries = OUCreation.Children
        Using newgroup As DirectoryEntry = nouveauGroupe.Add("CN=" & samaNewGroup, "group")
            newgroup.Properties("sAMAccountName").Add(samaNewGroup)
            If description <> "" Then
                newgroup.Properties("description").Add(description)
            End If
            If SAMmanager <> "?" Then
                SetADLDAPProperty(newgroup, "managedBy", TransformeSAMACCOUNTenCN(SAMmanager))
            End If

            AppliquerChangement(newgroup)
            Commun.Journal("Nouvelle Equipe créée : " & description & " (" & samaNewGroup & ")")
        End Using
    End Sub
    ''' <summary>
    ''' Envoie un e-mail.
    ''' </summary>
    ''' <param name="sender">
    ''' Adresse e-mail de l'expéditeur.
    ''' Format possible : expediteur@x.fr ou expediteur@x.fr;ReplyTo:reponse@x.fr.
    ''' </param>
    ''' <param name="recipient">
    ''' Adresse(s) e-mail du ou des destinataires.
    ''' Plusieurs adresses peuvent être séparées par un point-virgule (;).
    ''' Préfixes pris en charge :
    ''' Cc: pour une copie carbone.
    ''' Bcc: pour une copie cachée.
    ''' </param>
    ''' <param name="subject">Sujet du message.</param>
    ''' <param name="body">Corps du message.</param>
    ''' <param name="attachmentString">
    ''' Optionnel. Chemin complet du fichier à joindre.
    ''' </param>
    ''' <remarks>
    ''' Exemples :
    ''' sender = "expediteur@x.fr"
    ''' sender = "expediteur@x.fr;ReplyTo:reponse@x.fr"
    ''' recipient = "dest1@x.fr;Cc:dest2@x.fr;Bcc:dest3@x.fr"
    ''' </remarks>
    Shared Sub SendEmail(ByVal sender As String, ByVal recipient As String, ByVal subject As String, ByVal body As String, Optional ByVal attachmentString As String = "")

        Try
            Using message As New MailMessage()
                Dim senderAddress As String = Trim(sender)
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

                Dim smtpServer As String = ini.ReadValue("GLOBAL", "SMTPServer")
                Dim sent As Boolean = False
                Dim lastError As Exception = Nothing

                For Each smtpTry In New Object() {New With {.Port = 587, .Ssl = True}, New With {.Port = 25, .Ssl = False}}
                    Try
                        Using mailSender As New SmtpClient(smtpServer, smtpTry.Port)
                            With mailSender
                                .UseDefaultCredentials = False
                                .Credentials = New Net.NetworkCredential(smtpLogin, smtpPassword, "IGBMC")
                                .EnableSsl = smtpTry.Ssl
                            End With

                            'Ajouter les destinataires
                            Dim tabDest As String() = Split(recipient, ";")
                            For i = 0 To UBound(tabDest)
                                Dim dest As String = Trim(tabDest(i))

                                If LCase(Left(dest, 3)) = "cc:" Then
                                    message.CC.Add(Trim(Mid(dest, 4)))
                                ElseIf LCase(Left(dest, 4)) = "bcc:" Then
                                    message.Bcc.Add(Trim(Mid(dest, 5)))
                                Else
                                    message.To.Add(dest)
                                End If
                            Next

                            message.Subject = subject

                            'detecter si le mail est en HTML
                            If Left(body, 21) = "<!DOCTYPE HTML PUBLIC" Or InStr(LCase(body), "</") > 0 Then
                                message.IsBodyHtml = True
                                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(body, Nothing, "text/html")
                                Dim bodyPlain As String = RemoveHtml(body)
                                Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString(bodyPlain, Nothing, "text/plain")
                                message.AlternateViews.Add(htmlView)
                            Else
                                Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString(body, Nothing, "text/plain")
                                message.AlternateViews.Add(plainView)
                            End If

                            If Not attachmentString = "" Then
                                Using msgAttach As New Attachment(attachmentString)
                                    message.Attachments.Add(msgAttach)
                                    mailSender.Send(message)
                                End Using
                            Else
                                mailSender.Send(message)
                            End If
                        End Using

                        sent = True
                        Exit For

                    Catch ex As Exception
                        lastError = ex
                        If smtpTry.Port = 587 Then
                            Commun.Journal("ERREUR : echec de l'envoi de mail sur le port 587 avec TLS : " & ex.Message & ". Nouvelle tentative sur le port 25 sans TLS.")
                        Else
                            Commun.Journal("ERREUR : echec de l'envoi de mail sur le port 25 sans TLS : " & ex.Message)
                        End If
                    End Try
                Next

                If Not sent AndAlso lastError IsNot Nothing Then
                    Throw lastError
                End If
            End Using
        Catch ex As Exception
            Commun.Journal("ERREUR : echec de l'envoi de mail : " & ex.Message)
        End Try
    End Sub
    Shared Function RemoveHtml(ByVal html As String) As String
        ' Remove HTML tags.
        Return Regex.Replace(html, "<.*?>", "")
    End Function

    Shared Function SansAccent(ByVal Chaine As String) As String
        Dim aOctets() As Byte = System.Text.Encoding.GetEncoding(1251).GetBytes(Chaine) 'converti en byte la chaine avec accents
        Dim sEnleverAccents As String = System.Text.Encoding.ASCII.GetString(aOctets) 'converti en string la chaine sans accents
        Return sEnleverAccents
    End Function
    Shared Function GIDMiniGroup() As Integer

        Dim resultat As Integer
        Dim ResultFields() As String = {"gidNumber"}
        Dim tabUIDN() As Integer
        Dim CtrlGIDNumberLibre As Boolean = False
        Dim i As Integer = -1

        Try

            Using resultLdapU As SearchResultCollection = SearchFilterAll("OU=Equipes,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr", "(&(objectClass=group))", SearchScope.Subtree, "gidNumber")
                For Each result As SearchResult In resultLdapU
                    If result.Properties.Contains("gidNumber") Then
                        i += 1
                        ReDim Preserve tabUIDN(i)
                        tabUIDN(i) = CInt(result.Properties("gidNumber")(0))
                    End If
                Next result
            End Using

            Dim k = 3500
            Do Until CtrlGIDNumberLibre = True
                k += 1
                CtrlGIDNumberLibre = True
                For j = 0 To UBound(tabUIDN)
                    If tabUIDN(j) = k Then
                        CtrlGIDNumberLibre = False
                    End If
                Next j
            Loop
            resultat = k
        Catch ex As Exception
            Commun.Journal("ERREUR : GIDMini : Récuération d'un GID libre : " & ex.Message, True)
        End Try
        Return resultat
    End Function
    Shared Function PrimaryGroupId(ByVal objectSID) As Integer
        Dim GroupIDBytes() As Byte = DirectCast(objectSID, Byte())
        Dim GroupSID As New System.Security.Principal.SecurityIdentifier(GroupIDBytes, 0)
        Dim SplitSID() As String = GroupSID.Value.Split("-"c)
        Dim RID As Integer = CInt(SplitSID(SplitSID.Length - 1))
        Return RID
    End Function

    ''' <summary>
    ''' Définit la valeur d'un attribut LDAP mono-valué sur un <see cref="DirectoryEntry"/>.
    ''' </summary>
    ''' <param name="de">Objet LDAP cible à modifier.</param>
    ''' <param name="pName">Nom de l'attribut LDAP à modifier.</param>
    ''' <param name="pValue">
    ''' Valeur à affecter à l'attribut.
    ''' Si la valeur est vide, l'attribut est effacé.
    ''' </param>
    ''' <param name="addValueto">
    ''' Si <c>False</c>, la valeur remplace la valeur existante.
    ''' Si <c>True</c>, la valeur est concaténée à la valeur existante.
    ''' </param>
    ''' <remarks>
    ''' Cette méthode n'appelle pas <c>CommitChanges()</c>.
    ''' Il faut appeler <c>CommitChanges()</c> ou la méthode d'application des changements après son utilisation.
    ''' </remarks>
    Public Shared Sub SetADLDAPProperty(ByVal de As DirectoryEntry, ByVal pName As String, ByVal pValue As String, Optional ByVal addValueto As Boolean = False)
        If pValue = "" Then pValue = Nothing

        If pValue IsNot Nothing Then
            If de.Properties.Contains(pName) Then
                If addValueto = False Then
                    de.Properties(pName)(0) = pValue
                Else
                    de.Properties(pName)(0) += pValue
                End If
            Else
                de.Properties(pName).Add(pValue)
            End If
        Else
            de.Properties(pName).Clear()
        End If
    End Sub
    Public Shared Sub SetADLDAPPropertyByte(ByVal de As DirectoryEntry, ByVal pName As String, ByVal pValue As Byte())
        If pValue IsNot Nothing Then
            If de.Properties.Contains(pName) Then
                de.Properties(pName).Clear()
                de.Properties(pName).Add(pValue)
            Else
                de.Properties(pName).Add(pValue)
            End If
        Else
            de.Properties(pName).Clear()
        End If
    End Sub

    '''' <summary>
    ''' Définit la valeur d'un attribut LDAP multi-valué sur un <see cref="DirectoryEntry"/>.
    ''' </summary>
    ''' <param name="de">Objet LDAP cible à modifier.</param>
    ''' <param name="pName">Nom de l'attribut LDAP à modifier.</param>
    ''' <param name="pValues">
    ''' Ensemble des valeurs à affecter à l'attribut.
    ''' Les valeurs nulles, vides ou composées uniquement d'espaces sont ignorées.
    ''' </param>
    ''' <param name="addValueto">
    ''' Si <c>False</c>, les valeurs existantes sont remplacées.
    ''' Si <c>True</c>, les nouvelles valeurs sont ajoutées sans supprimer les anciennes,
    ''' en évitant les doublons.
    ''' </param>
    ''' <remarks>
    ''' Cette méthode n'appelle pas <c>CommitChanges()</c>.
    ''' Il faut appeler <c>CommitChanges()</c> ou la méthode d'application des changements après son utilisation.
    ''' </remarks>
    Public Shared Sub SetADLDAPPropertyMulti(ByVal de As DirectoryEntry, ByVal pName As String, ByVal pValues As String(), Optional ByVal addValueto As Boolean = False)
        If pValues Is Nothing Then pValues = New String() {}

        Dim valeursNettoyees As New List(Of String)

        For Each value As String In pValues
            If value IsNot Nothing Then
                Dim v As String = value.Trim()
                If v <> "" AndAlso Not valeursNettoyees.Contains(v) Then
                    valeursNettoyees.Add(v)
                End If
            End If
        Next

        If addValueto = False Then
            de.Properties(pName).Clear()

            For Each v As String In valeursNettoyees
                de.Properties(pName).Add(v)
            Next

            Return
        End If

        For Each v As String In valeursNettoyees
            Dim dejaPresent As Boolean = False

            For Each existingValue As Object In de.Properties(pName)
                If existingValue IsNot Nothing AndAlso String.Equals(CStr(existingValue).Trim(), v, StringComparison.OrdinalIgnoreCase) Then
                    dejaPresent = True
                    Exit For
                End If
            Next

            If Not dejaPresent Then
                de.Properties(pName).Add(v)
            End If
        Next
    End Sub
    ''' <summary>
    ''' <para>Transforme un SAMAccount en CN (sans "LDAP://" ou l'inverse.</para>
    ''' </summary>
    ''' <param name="loginOuCN">login ou CN (sans "LDAP://").</param>
    ''' <remarks>Retourne "" si le SAMaccountName n'existe pas. Peut servir a tester l'existance d'un groupe ou d'un utilisateur</remarks>
    Shared Function TransformeSAMACCOUNTenCN(ByVal loginOuCN As String) As String

        Dim resultat As String = ""
        If loginOuCN <> "" Then

            If Left(loginOuCN, 3) = "CN=" Then
                If DirectoryEntry.Exists("LDAP://" & Commun.LdapPath(loginOuCN)) Then
                    Using objectAD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(loginOuCN), Nothing, Nothing, auth)
                        resultat = objectAD.Properties("sAMAccountName").Value
                    End Using
                End If

            Else
                Using Ldap As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Nothing, Nothing, auth)

                    Using searcher As DirectorySearcher = New DirectorySearcher(Ldap)
                        searcher.Filter = "(&(sAMAccountName=" & loginOuCN & "))"
                        Dim result As SearchResult = searcher.FindOne()
                        If Not result Is Nothing Then
                            Dim pathCN As String = Replace(result.Path, "LDAP://" & Commun.LdapServerPrefix(), "")
                            resultat = pathCN
                        End If
                        result = Nothing
                    End Using

                End Using
            End If
        End If
        Return resultat
        'GC.Collect()
    End Function
    ''' <summary>
    ''' <para>Renvoi True or False si un utilisateur appartient à un groupe ou non.</para>
    ''' <para>Gestion de l'appartenance au groupe principal.</para>
    ''' <para>La fonction gere le login ou le DN pour le user ou le groupe.</para>
    ''' </summary>
    ''' <param name="user">login ou CN de l'utilisateur.</param>
    ''' <param name="groupe">login ou CN du groupe.</param>
    ''' <remarks></remarks>
    Shared Function AppartientGroup(ByVal user As String, ByVal groupe As String) As Boolean
        'La fonction gere le login ou le DN pour le user ou le groupe
        'Gestion de l'appartenance au groupe principal

        user = Replace(user, "LDAP://" & Commun.LdapServerPrefix(), "")
        If Strings.Left(groupe, 3) <> "CN=" Then
            groupe = Commun.TransformeSAMACCOUNTenCN(groupe)
        End If
        If Strings.Left(user, 3) <> "CN=" Then
            user = Commun.TransformeSAMACCOUNTenCN(user)
        End If
        Dim result As Boolean = False


        If DirectoryEntry.Exists("LDAP://" & Commun.LdapPath(user)) Then
            Using userEntry As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(user), Nothing, Nothing, auth)

                Dim ridGroup As Integer
                If DirectoryEntry.Exists("LDAP://" & Commun.LdapPath(groupe)) Then
                    Using groupEntry As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(groupe), Nothing, Nothing, auth)
                        ridGroup = Commun.PrimaryGroupId(groupEntry.Properties("objectSid").Value)
                    End Using

                    If userEntry.Properties("primaryGroupID").Value = ridGroup Then
                        result = True
                        GoTo sortie
                    End If

                End If
                If userEntry.Properties("memberOf").IndexOf(groupe) <> -1 Then
                    result = True
                End If
sortie:

            End Using
        End If
        Return result
    End Function
    ''' <summary>
    ''' <para>Ajoute ou retire un objet d'un groupe.</para>
    ''' <para>Controle l'appartenance au groupe avant l'action.</para>
    ''' <para>La fonction gere le login ou le DN pour le user ou le groupe.</para>
    ''' </summary>
    ''' <param name="login">login ou CN de l'objet à ajouter/retirer.</param>
    ''' <param name="group">login ou CN du groupe auquel l'objet doit etre ajouté/retiré.</param>
    ''' <param name="action">"Add" pour ajouter, "Remove" pour retirer.</param>
    ''' <remarks>
    ''' La fonction gere le login ou le DN pour le login ou le group
    ''' Controle l'appartenance au groupe avant l'action. Si l'utilisateur doit etre ajouté, verification si l'utilisateur est deja membre du groupe
    ''' </remarks>
    Shared Function AddRemoveADGroup(ByVal login As String, ByVal group As String, ByVal action As String) As Boolean
        Dim result As Boolean = False
        login = Replace(login, "LDAP://" & Commun.LdapServerPrefix(), "")
        Dim CNuser As String = login
        Dim CNgroup As String = group
        Try
            If Strings.Left(group, 3) <> "CN=" Then
                CNgroup = Commun.TransformeSAMACCOUNTenCN(group)
            End If
            If Strings.Left(login, 3) <> "CN=" Then
                CNuser = Commun.TransformeSAMACCOUNTenCN(login)
            End If

            'Control avant l'action
            Dim ctrlAction As Boolean = True
            If action = "Add" Then
                ctrlAction = Not Commun.AppartientGroup(CNuser, CNgroup)
            Else
                ctrlAction = Commun.AppartientGroup(CNuser, CNgroup)
            End If

            If CNgroup = "" Or CNuser = "" Then
                Commun.Journal("ERREUR : AddRemoveADGroup : login ou group inconnu : " & login & " - " & group)
            Else
                If ctrlAction = True Then
                    Using GroupPath As New DirectoryEntry("LDAP://" & Commun.LdapPath(CNgroup), Nothing, Nothing, auth)
                        GroupPath.Invoke(action, New Object() {"LDAP://" & Commun.LdapPath(CNuser)})
                        result = False
                    End Using
                    'GC.Collect()
                End If
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : AddRemoveADGroup : " & action & " : " & login & " - " & group & " : " & ex.Message, True)
        End Try
        Return result
    End Function

    ''' <summary>
    ''' <para>Renvoi un tableau contenant les groupes auquel un autre groupe appartient.</para>
    ''' </summary>
    ''' <param name="SamAccount">sAMAccountName du groupe.</param>
    ''' <remarks></remarks>
    Shared Function GroupeMembresDe(ByVal SamAccount As String) As Array
        Dim tabMembres As String() = Nothing
        Dim i As Integer = -1
        Try
            'Recupération de l'attribut Member pour le mettre dans le tableau des resultats
            Using AD As DirectoryEntry = New DirectoryEntry("LDAP://" & LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Nothing, Nothing, auth)
                Using searcherGroup As DirectorySearcher = New DirectorySearcher(AD)
                    searcherGroup.Filter = "(&(objectClass=group) (sAMAccountName=" & SamAccount & "))"
                    searcherGroup.PropertiesToLoad.Add("Member")
                    searcherGroup.PropertiesToLoad.Add("objectSid")
                    Dim resultGroup As SearchResult = searcherGroup.FindOne()
                    If Not resultGroup Is Nothing Then
                        Dim Group As DirectoryEntry = resultGroup.GetDirectoryEntry
                        For Each unMembre In Group.Properties("memberof")
                            i += 1
                            ReDim Preserve tabMembres(i)
                            tabMembres(i) = unMembre.ToString
                        Next unMembre

                    End If

                End Using
            End Using
            Return tabMembres

        Catch ex As Exception

        End Try
    End Function
    Shared Function IsMembreEquipeAdministratif(ByVal user As DirectoryEntry) As Boolean
        Dim tabEquipeADM As String() = Split(ini.ReadValue("MODIFAUTO", "EquipesAdministratives"), ",")
        Try
            Dim control As Boolean = False
            For i As Integer = 0 To user.Properties("MemberOf").Count - 1
                'Récupère la chaine LDAP.
                Dim sProp As String = user.Properties("MemberOf")(i)
                Dim nomEquipe As String = sProp.Substring(3, sProp.IndexOf(",") - 3)
                If Array.IndexOf(tabEquipeADM, nomEquipe) <> -1 Then
                    Return True
                    Exit Function
                End If
            Next i
        Catch ex As Exception

        End Try

        Return False

    End Function
    ''' <summary>
    ''' <para>Renvoi le nom court de l'equipeinfo en fonction de la destination.</para>
    ''' </summary>
    ''' <param name="nomCourtDest">Nom court de la destination.</param>
    ''' <remarks></remarks>
    Shared Function RecupEquipeinfo(ByVal nomCourtDest As String) As String

        Dim equipeInfo As String = ""
        Try
            Dim samaccountDest As String = nomCourtDest & " grp"
            Dim tabMembresGroupDest As String() = Commun.GroupeMembresDe(samaccountDest)
            Dim chaineARechercher As String = "_dest,OU=Gestion_Destinations,OU=Equipes,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr"
            If Not tabMembresGroupDest Is Nothing Then
                For i = 0 To UBound(tabMembresGroupDest)
                    If InStr(tabMembresGroupDest(i), chaineARechercher) > 0 Then
                        equipeInfo = Replace(Commun.TransformeSAMACCOUNTenCN(tabMembresGroupDest(i)), "_dest", "")
                        Exit For
                    End If
                Next
            End If
        Catch
            Return equipeInfo
        End Try
        Return equipeInfo

    End Function

    ''' <summary>
    ''' <para>Renvoi un attribut AD d'un obbjet.</para>
    ''' <para>Renvoi Nothing si l'attribut n'est pas defini.</para>
    ''' </summary>
    ''' <param name="loginOuCN">sAMAccountName de l'objet.</param>
    ''' <param name="attribut">attribut a récupérer.</param>
    ''' <remarks>La fonction gere le login ou le CN pour "loginOuCN"</remarks>
    Shared Function FindAttribut(ByVal loginOuCN As String, ByVal attribut As String) As Object
        Dim resultat = Nothing
        Dim CNObject As String = loginOuCN

        If Left(loginOuCN, 3) <> "CN=" Then
            CNObject = TransformeSAMACCOUNTenCN(loginOuCN)
        End If


        If CNObject <> "" Then
            If DirectoryEntry.Exists("LDAP://" & Commun.LdapPath(CNObject)) Then
                Using objectAD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(CNObject), Nothing, Nothing, auth)
                    If objectAD.Properties.Contains(attribut) Then
                        If objectAD.Properties.Contains(attribut) Then
                            resultat = objectAD.Properties(attribut).Value
                        End If
                    End If
                End Using
            End If
        End If

        Return resultat

    End Function
    Shared Function IsException(ByVal user As String, ByVal repZone As String) As Boolean
        Dim resultat As Boolean = False
        If tabExcepUser Is Nothing Then CreateTabExceptionUser()
        Dim ligne As String = LCase(user & "," & repZone)
        If repZone <> "" Then
            If Array.IndexOf(tabExcepUser, ligne) > -1 Then
                resultat = True
                Exit Function
            End If
        Else
            For i = 0 To UBound(tabExcepUser)
                If Split(tabExcepUser(i), ",")(0) = user Then
                    resultat = True
                    Exit For
                End If
            Next
        End If
        Return resultat
    End Function
    Shared Function UserExists(ByVal login As String) As Boolean
        Dim resultat As Boolean = False
        If TransformeSAMACCOUNTenCN(login) <> "" Then resultat = True
        Return resultat
    End Function
    Shared Sub CreateTabExceptionUser()
        Dim ligne As String
        Dim h As Integer = -1
        Try
            Using monStreamReader As StreamReader = New StreamReader("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\UsersExceptions.txt")
                Do
                    ligne = monStreamReader.ReadLine()
                    If ligne Is Nothing Then Exit Do
                    If String.IsNullOrWhiteSpace(ligne) Then Continue Do

                    Dim tabLigneExceptUser As String() = Split(ligne, ","c)
                    If tabLigneExceptUser.Length < 3 Then
                        Commun.Journal("ERREUR : CreateTabExceptionUser : ligne invalide dans UsersExceptions.txt : " & ligne, True)
                        Continue Do
                    End If

                    Dim dateExpire As DateTime
                    If Not DateTime.TryParseExact(
                      tabLigneExceptUser(1).Trim(),
                      "dd/MM/yyyy",
                      System.Globalization.CultureInfo.InvariantCulture,
                      System.Globalization.DateTimeStyles.None,
                      dateExpire
                  ) Then
                        Commun.Journal("ERREUR : CreateTabExceptionUser : date invalide dans UsersExceptions.txt : " & ligne, True)
                        Continue Do
                    End If

                    If dateExpire >= Now.Date Then
                        If UserExists(tabLigneExceptUser(0).Trim()) = True Then
                            h += 1
                            ReDim Preserve tabExcepUser(h)
                            tabExcepUser(h) = LCase(tabLigneExceptUser(0).Trim() & "," & tabLigneExceptUser(2).Trim())
                        Else
                            Commun.Journal("ERREUR : CreateTabExceptionUser : utilisateur inexistant dans UsersExceptions.txt : " & ligne, True)
                        End If
                    End If
                Loop
            End Using

        Catch ex As Exception
            Commun.Journal("ERREUR : CreateTabExceptionUser : " & ex.Message, True)
        End Try
    End Sub

    ''' <summary>
    ''' <para>Verifie l'existance d'un login ou alias dans l'historique.</para>
    ''' <para>Verifie si le compte adm existe.</para>
    ''' </summary>
    ''' <param name="loginAlias">Login ou alias a controler (pour eviter l'ambiguité dans le cas d'un alias, il faut le mettre entre "'").</param>
    ''' <param name="idUser">employeeID de l'utilisateur.</param>
    ''' <remarks></remarks>
    Shared Function ctrlAliasDispo(ByVal loginAlias As String) As Boolean

        Dim result As Boolean = False
        loginAlias = "'" & loginAlias & "'"
        'Dim loginAliasID As String = loginAlias & "," & idUser

        Dim fichierComplet As String
        Using srAll As New StreamReader("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\Alias.txt")
            fichierComplet = srAll.ReadToEnd
        End Using

        If tabLoginAliasID Is Nothing Then
            Using sr As New StreamReader("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\Alias.txt")
                Dim ligne As String
                Dim i = -1

                Do
                    i += 1
                    'creation d'un tableau issu du fichier d'historique, comprenant une ligne alias/login et l'employeeID
                    ReDim Preserve tabLoginAliasID(i)

                    'creation d'un tableau issu du fichier d'historique, comprenant une ligne juste avec l'alias/login
                    ReDim Preserve tabLoginAlias(i)
                    ligne = sr.ReadLine()
                    If ligne Is Nothing Then Exit Do
                    tabLoginAliasID(i) = LCase(ligne)
                    tabLoginAlias(i) = Split(Split(ligne, ",")(0), "§")(1)
                Loop Until ligne Is Nothing

            End Using

        End If

        'Sinon on cherche dans le tableau des login/alias, si celui proposé est dispo
        If result = False Then
            If InStr(fichierComplet, loginAlias) = 0 Then
                result = True
                Commun.Journal("ctrlAliasDispo : Creation d'un nouveau login/alias : " & loginAlias, True)
            End If
        End If
        Return result

    End Function
    Shared Function DetermineAliasLibre(ByVal prenom As String, ByVal nom As String, ByVal userID As String) As String() 'le tableau retourné ne contient pas de "'"

        Dim aliasMailTab(0) As String
        Dim aliasAattribuer As String
        Dim tablogin As String(,) = Commun.CreateTabHistoAliasLogin("alias")
        Dim index As Integer = Commun.MultiIndexOf(tablogin, userID, 0)
        Dim aliasRecupere As String
        Dim aliasRecupereTab As String() = Nothing

        'on construit un alias avec le prenom et le nom
        Dim aliasConstruit As String = prenom & "." & nom
        aliasConstruit = LCase(Replace(aliasConstruit, " ", "-"))
        aliasConstruit = LCase(Replace(aliasConstruit, "'", ""))


        'si le numero d'emloyé existe deja dans l'historique, on lui réatribue les memes alias
        If index <> -1 And userID <> "0" Then
            aliasAattribuer = tablogin(1, index)
            aliasRecupere = Replace(aliasAattribuer, "'", "")
            aliasRecupereTab = Split(aliasRecupere, ",")

            Dim index2 As Integer = Array.IndexOf(aliasRecupereTab, aliasConstruit)
            'Si on trouve l'alias construit dans les alias récupérés, on place l'alias construit, qui sera le principal (GDPI), en premier
            If index2 > -1 Then
                Dim echange As String
                echange = aliasRecupereTab(0)
                aliasRecupereTab(0) = aliasConstruit
                aliasRecupereTab(index2) = echange
                'le tableau retourné ne contient pas de "'"
                Return aliasRecupereTab
                Exit Function
            End If
        End If

        Dim ctrlLoginLibre = False


        'ne pas mettre 'integer' pour n, comme ca, la premiere boucle se fait sans chiffre à la fin
        Dim i = Nothing
        Do
            aliasConstruit = aliasConstruit & i
            If Commun.ctrlAliasDispo(aliasConstruit) = True Then
                aliasMailTab(0) = aliasConstruit
                If Not aliasRecupereTab Is Nothing Then
                    ReDim Preserve aliasMailTab(UBound(aliasRecupereTab) + 1)
                    aliasRecupereTab.CopyTo(aliasMailTab, 1)
                End If
                'le tableau retourné ne contient pas de "'"
                Return aliasMailTab
                Exit Function
            End If
            i += 1
        Loop

    End Function
    Shared Sub ajoutAliasFichierHisto(ByVal aliasLoginMail As String, ByVal userID As String)

        Dim type As String = "login"
        If InStr(aliasLoginMail, ".") > 0 Then type = "alias"

        If userID = "" Then userID = "0"

        If type = "login" Then
            Dim tabLogin(,) As String = Commun.CreateTabHistoAliasLogin("login")
            Dim index3 As Integer = Commun.MultiIndexOf(tabLogin, userID, 0)
            If index3 > -1 Then
                tabLogin(1, index3) = aliasLoginMail
                Using sw1 As New StreamWriter("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\Login.txt", False)
                    'on ecrit le tableau alias dans le fichier
                    For j = 0 To UBound(tabLogin, 2)
                        sw1.WriteLine(tabLogin(0, j) & "§" & tabLogin(1, j))
                    Next
                End Using
            Else
                Using sw2 As New StreamWriter("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\Login.txt", True)
                    sw2.WriteLine(userID & "§" & aliasLoginMail)
                End Using
            End If
        End If

        Dim tabAlias(,) As String = Commun.CreateTabHistoAliasLogin("alias")
        Dim index4 As Integer = Commun.MultiIndexOf(tabAlias, userID, 0)
        If index4 > -1 Then
            If InStr(tabAlias(1, index4), "'" & aliasLoginMail & "'") = 0 Then
                tabAlias(1, index4) = tabAlias(1, index4) & ",'" & aliasLoginMail & "'"
                Using sw1 As New StreamWriter("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\Alias.txt", False)
                    'on ecrit le tableau alias dans le fichier
                    For j = 0 To UBound(tabAlias, 2)
                        sw1.WriteLine(tabAlias(0, j) & "§" & tabAlias(1, j))
                    Next
                End Using
            End If
        Else
            Using sw2 As New StreamWriter("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\Alias.txt", True)

                sw2.WriteLine(userID & "§'" & aliasLoginMail & "'")
            End Using
        End If



    End Sub
    ''' <summary>
    ''' <para>Active ou Desactive un compte.</para>
    ''' <para>Verifie si le compte existe.</para>
    ''' </summary>
    ''' <param name="login">sAMAccountName de l'utilisateur.</param>
    ''' <param name="action">"Active" ou "Desactive".</param>
    ''' <remarks></remarks>
    Shared Sub ReactiveDesactiveCompte(ByVal login As String, ByVal action As String)
        Dim cheminLdapCompte As String = Commun.TransformeSAMACCOUNTenCN(login)
        action = LCase(action)
        If cheminLdapCompte <> "" Then
            Using user As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(cheminLdapCompte), Nothing, Nothing, auth)

                If action = "active" Then
                    user.NativeObject.accountdisabled = False
                    AppliquerChangement(user)
                End If

                If action = "desactive" Then
                    user.NativeObject.accountdisabled = True
                    AppliquerChangement(user)
                End If
            End Using
        End If
    End Sub
    ''' <summary>
    ''' <para>Active ou Desactive un compte.</para>
    ''' <para>Verifie si le compte existe.</para>
    ''' </summary>
    ''' <param name="objUsr">Objet LDAP de l'utilisateur adm.</param>
    ''' <param name="action">"Active" ou "Desactive".</param>
    ''' <remarks></remarks>
    Shared Sub ReactiveDesactiveCompte(ByVal objUsr As DirectoryEntry, ByVal action As String)

        action = LCase(action)

        If LCase(action) = "active" Then
            objUsr.NativeObject.accountdisabled = False
            AppliquerChangement(objUsr)
        End If

        If LCase(action) = "desactive" Then
            objUsr.NativeObject.accountdisabled = True
            AppliquerChangement(objUsr)
        End If

    End Sub
    ''' <summary>
    ''' <para>Recherche la ligne correspondante à une entree dans un tableau multidimensionnel.</para>
    ''' </summary>
    ''' <param name="tab">Nom du tableau.</param>
    ''' <param name="valeurCherchee">Valeur recherchée dans le tableau.</param>
    ''' <param name="colRecherche">indique la collone dans laquelle la recherche doit se faire (a partir de 0).</param>
    ''' <param name="exact">si "true la valeur rechercher doit correspondre exactement a la valeur de la cellule ("False"= recherche instr)</param>
    ''' <remarks></remarks>
    Shared Function MultiIndexOf(ByVal tab As String(,), ByVal valeurCherchee As String, ByVal colRecherche As Integer, Optional ByVal exact As Boolean = True) As Integer

        Dim resultat As Integer = -1

        For i = 0 To UBound(tab, 2)
            If exact = True Then
                If tab(colRecherche, i) = valeurCherchee Then
                    resultat = i
                    Exit For
                End If
            Else
                If InStr(tab(colRecherche, i), valeurCherchee) > 0 Then
                    resultat = i
                    Exit For
                End If
            End If
        Next

        Return resultat

    End Function
    ''' <summary>
    ''' <para>Creer un tableau avec l'historique des comptes pour les alias mail ou les logins.</para>
    ''' </summary>
    ''' <param name="type">"Alias" ou "login".</param>
    ''' <remarks></remarks>
    Shared Function CreateTabHistoAliasLogin(ByVal type As String) As String(,)

        Dim i As Integer = -1
        Dim tabAliasLogin(1, -1) As String
        Dim cheminFichier As String
        If type = "alias" Then
            cheminFichier = "\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\Alias.txt"
        ElseIf type = "login" Then
            cheminFichier = "\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\Login.txt"
        End If

        Using monStreamReader As StreamReader = New StreamReader(cheminFichier)

            Dim ligneA As String

            Do
                ligneA = monStreamReader.ReadLine()
                If Not ligneA Is Nothing Then
                    i += 1
                    ReDim Preserve tabAliasLogin(1, i)
                    tabAliasLogin(0, i) = Split(ligneA, "§")(0)
                    tabAliasLogin(1, i) = Split(ligneA, "§")(1)
                End If
            Loop Until ligneA Is Nothing

        End Using
        Return tabAliasLogin
    End Function
    Shared Sub AppliquerChangement(ByVal objet As DirectoryEntry)
        objet.CommitChanges()
    End Sub
    Shared Function SearchFilterAll(
    ByVal ou As String,
    ByVal filter As String,
    ByVal scope As SearchScope,
    Optional ByVal propertyToLoad As String = "") As SearchResultCollection

        Dim objAD As New DirectoryEntry("LDAP://" & Commun.LdapPath(ou), Nothing, Nothing, auth)
        Dim searcher As New DirectorySearcher(objAD)

        searcher.Filter = filter
        searcher.PageSize = 5000
        searcher.SearchScope = scope

        If propertyToLoad <> "" Then
            For Each propertyName As String In Split(propertyToLoad, ",")
                propertyName = propertyName.Trim()
                If propertyName <> "" Then
                    searcher.PropertiesToLoad.Add(propertyName)
                End If
            Next
        End If

        Return searcher.FindAll()
    End Function
    Shared Function SearchFilterOne(
    ByVal ou As String,
    ByVal filter As String,
    ByVal scope As SearchScope,
    Optional ByVal propertyToLoad As String = "") As SearchResult

        Using objAD As New DirectoryEntry("LDAP://" & Commun.LdapPath(ou), Nothing, Nothing, auth)
            Using searcher As New DirectorySearcher(objAD)
                searcher.Filter = filter
                searcher.PageSize = 5000
                searcher.SearchScope = scope

                If propertyToLoad <> "" Then
                    For Each propertyName As String In Split(propertyToLoad, ",")
                        propertyName = propertyName.Trim()
                        If propertyName <> "" Then
                            searcher.PropertiesToLoad.Add(propertyName)
                        End If
                    Next
                End If

                Return searcher.FindOne()
            End Using
        End Using
    End Function
    Shared Function SearchFilterAll(ByVal objAD As DirectoryEntry, ByVal filter As String, ByVal scope As SearchScope, Optional ByVal propertyToLoad As String = "") As SearchResultCollection

        Dim searcher As New DirectorySearcher(objAD)

        searcher.Filter = filter
        searcher.PageSize = 5000
        searcher.SearchScope = scope

        If propertyToLoad <> "" Then
            For Each propertyName As String In Split(propertyToLoad, ",")
                propertyName = propertyName.Trim()
                If propertyName <> "" Then
                    searcher.PropertiesToLoad.Add(propertyName)
                End If
            Next
        End If

        Return searcher.FindAll()
    End Function

    Shared Function UIDNumberMini() As String

        Dim k = 1000
        Using LdapU As DirectoryEntry = New DirectoryEntry("LDAP://" & LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Nothing, Nothing, auth)
            Dim ResultFields() As String = {"uidNumber"}
            Dim searcherLdapU As DirectorySearcher = New DirectorySearcher(LdapU)
            Dim CtrlUIDNumberLibre As Boolean = False
            Dim i As Integer = -1
            Dim nouveauUIDNumber As Integer

            searcherLdapU.Filter = "(&(objectClass=posixAccount))"
            searcherLdapU.SearchScope = SearchScope.Subtree
            searcherLdapU.PropertiesToLoad.AddRange(ResultFields)
            searcherLdapU.PageSize = 5000
            Dim tabUIDN() As Integer
            Dim resultLdapU As SearchResultCollection = searcherLdapU.FindAll()
            For Each result As SearchResult In resultLdapU
                If result.Properties.Contains("uidNumber") Then
                    i += 1
                    ReDim Preserve tabUIDN(i)

                    tabUIDN(i) = result.Properties("uidNumber")(0).ToString
                End If
            Next
            Do Until CtrlUIDNumberLibre = True
                k += 1
                CtrlUIDNumberLibre = True
                For j = 0 To UBound(tabUIDN)
                    If tabUIDN(j) = k Then
                        CtrlUIDNumberLibre = False
                    End If
                Next
            Loop
            nouveauUIDNumber = k
        End Using
        Return k

    End Function
    Shared Function AccountIsDisabled(ByVal objUser As DirectoryEntry) As Boolean
        'Just posting the relevant parts of the code
        Dim result As Boolean = True
        Const ADS_UF_ACCOUNTDISABLE As Integer = 2

        Dim Flags As Integer = objUser.Properties("userAccountControl").Value
        If CBool(Flags And ADS_UF_ACCOUNTDISABLE) Then
            result = True
        Else
            result = False
        End If
        Return result
    End Function
    ''' <summary>
    ''' <para>Renvoi un tableau contenant le CN des membres d'un groupe.</para>
    ''' </summary>
    ''' <param name="GroupSamAccount">sAMAccountName du groupe.</param>
    ''' <param name="avecMembresGroupePrincipal">True ou False pour ajouter les membres qui ont ce groupe comme groupe principal.</param>
    ''' <remarks></remarks>
    Shared Function MembresDuGroupe(ByVal GroupSamAccount As String, Optional ByVal avecMembresGroupePrincipal As Boolean = False) As Array
        Dim tabMembres As String()
        Dim i As Integer = -1
        Try
            'Recupération de l'attribut Member pour le mettre dans le tableau des resultats
            Using AD As DirectoryEntry = New DirectoryEntry("LDAP://" & LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Nothing, Nothing, auth)
                Using searcherGroup As DirectorySearcher = New DirectorySearcher(AD)
                    searcherGroup.Filter = "(&(objectClass=group) (sAMAccountName=" & GroupSamAccount & "))"
                    searcherGroup.PropertiesToLoad.Add("Member")
                    searcherGroup.PropertiesToLoad.Add("objectSid")
                    Dim resultGroup As SearchResult = searcherGroup.FindOne()
                    Using Group As DirectoryEntry = New DirectoryEntry(resultGroup.Path, Nothing, Nothing, auth)
                        For Each unMembre In Group.Properties("member")
                            i += 1
                            ReDim Preserve tabMembres(i)
                            tabMembres(i) = unMembre.ToString
                        Next unMembre


                        If avecMembresGroupePrincipal = True Then

                            Dim RID As Integer = PrimaryGroupId(Group.Properties("objectSid")(0))

                            Using searcher As DirectorySearcher = New DirectorySearcher(AD)
                                searcher.Filter = "(&(objectClass=user) (primaryGroupID=" & RID & "))"
                                searcher.PageSize = 5000
                                Using resultCollection As SearchResultCollection = searcher.FindAll()
                                    For Each result As SearchResult In resultCollection
                                        i += 1
                                        ReDim Preserve tabMembres(i)
                                        tabMembres(i) = Replace(result.Path, "LDAP://" & Commun.LdapServerPrefix(), "")
                                    Next result
                                End Using
                            End Using
                        End If
                    End Using
                End Using
            End Using
            Return tabMembres
        Catch ex As Exception

        End Try
    End Function


    Shared Sub AfficherBarre(texte As String, valeur As Integer, max As Integer, Optional afficherPourcentage As Boolean = True)
        If Not Environment.UserInteractive Then Exit Sub

        Const largeur As Integer = 50

        If max <= 0 Then max = 1

        Dim progression As Double = CDbl(valeur) / CDbl(max)
        If progression < 0 Then progression = 0
        If progression > 1 Then progression = 1


        Dim remplis As Integer = CInt(Math.Floor(progression * largeur))

        Dim info As String
        If afficherPourcentage Then
            info = CInt(progression * 100) & "%"
        Else
            info = valeur & "/" & max
        End If

        Dim barre As String = "[" &
            New String("◼"c, remplis) &
            New String("◻"c, largeur - remplis) &
            "]"

        Console.Title = texte & " " & barre & " " & info
    End Sub



End Class


Public Class application
    'Shared Function exit()
    'End Function
    Shared Function productname() As String
        Return My.Application.Info.AssemblyName
    End Function
    Shared Function path() As String
        Return My.Application.Info.DirectoryPath.ToString()
    End Function
End Class
Public Module MyExtensions
    <Runtime.CompilerServices.Extension()>
    Public Sub Add(Of T)(ByRef arr As T(), item As T)
        If arr IsNot Nothing Then
            If Array.IndexOf(arr, item) = -1 Then
                Array.Resize(arr, arr.Length + 1)
                arr(arr.Length - 1) = item
            End If
        Else
            ReDim arr(0)
            arr(0) = item
        End If
    End Sub
    <Runtime.CompilerServices.Extension()>
    Public Function Egal(Of T)(ByRef a As T(), b As T()) As Boolean
        Dim result As Boolean = True

        If Not a Is Nothing Xor Not b Is Nothing Then
            Return False
            Exit Function
        End If
        If a Is Nothing And b Is Nothing Then
            Return True
            Exit Function
        End If
        If a.Length <> b.Length Then result = False

        For Each value In a
            If Array.IndexOf(b, value) = -1 Then
                Return False
                Exit Function
            End If
        Next
        Return result
    End Function
End Module
