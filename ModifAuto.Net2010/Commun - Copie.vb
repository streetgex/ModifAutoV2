Imports System.IO
Imports System.Net.Mail
Imports System.DirectoryServices


Public Class Commun
    Public Shared tabExcepUser As String()
    Public Shared tabVar(1, 0) As String
    'Public Shared ftp = New Collection
    Public Shared journalECHECMail As String = ""
    Public Shared controlSendMail As Boolean = False
    Public Shared fichierLog As String

    Shared Sub recupDataConnect()
        Dim monStreamReader As StreamReader = New StreamReader("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\ScriptSteph.ini")
        Dim ligne As String
        Dim reponse As String = ""
        Dim i As Integer = 0
        'Dim var As String

        Try
            Do
                ligne = monStreamReader.ReadLine()
                If Left(ligne, 1) = "[" And Right(ligne, 1) = "]" Then
                    'Dim var As String = Mid(ligne, 2, Len(ligne) - 2)
                    Dim var = ligne


                    If InStr(var, "Password") = 0 Then
                        reponse = monStreamReader.ReadLine()
                    Else
                        Dim psswdtmp As String = ""
                        Do
                            psswdtmp = monStreamReader.ReadLine()
                            Dim rep As String = Chr(98) & Chr(72) & Chr(113) 'bHq

                            Dim pos As Integer = Strings.InStr(psswdtmp, rep)
                            If pos <> 0 Then
                                reponse = reponse & Strings.Mid(psswdtmp, pos + 3, 1)
                            End If

                        Loop Until psswdtmp = "-----END RSA PRIVATE KEY-----"
                    End If
                    'test de valeur des variables récupérées
                    If reponse = "" Then
                        Journal("ERREUR : FTP : Recuperation de la donnée ini : " & var, True)
                        Application.Exit()
                    End If
                    ReDim Preserve tabVar(1, i)
                    tabVar(0, i) = var
                    tabVar(1, i) = reponse
                    i += 1
                    reponse = ""

                End If

            Loop Until ligne Is Nothing

            monStreamReader.Close()
            monStreamReader.Dispose()
        Catch ex As Exception
            Journal("ERREUR : FTP : Recuperation des données ini : " & ex.Message, True)
            Application.Exit()
        End Try
    End Sub
    Shared Function RecupVar(ByVal var As String) As String
        If tabVar(0, 0) Is Nothing Then
            recupDataConnect()
        End If
        Dim i As Integer
        For i = 0 To UBound(tabVar, 2)
            If tabVar(0, i) = var Then
                Return tabVar(1, i)
                End
            End If
        Next i
        Journal("ERREUR : RecupVar : Cette donnée n'existe pas : " & var & " : ", True)
        Application.Exit()
    End Function
    Shared Sub Journal(ByVal erreur As String, Optional ByVal envoyerMail As Boolean = False)

        If envoyerMail = True Then
            controlSendMail = True
        End If

        If Left(erreur, 6) = "ERREUR" Then
            journalECHECMail = journalECHECMail & vbCrLf & Now & " : " & erreur
            erreur = vbTab & erreur
        End If

        If Not File.Exists(fichierLog) Then
            IO.File.Create(fichierLog)
        End If
        File.AppendAllText(fichierLog, Now & " : " & erreur & vbCrLf)

    End Sub
    Shared Sub SendEmail(ByVal sender As String, _
      ByVal recipient As String, ByVal subject As String, _
      ByVal body As String, Optional ByVal attachmentString As String = "")

        Try

            Dim fromAddress As New MailAddress(sender)
            Dim toAddress As New MailAddress(recipient)
            Dim message As New MailMessage(fromAddress, toAddress)

            Dim mailSender As SmtpClient = New SmtpClient(Commun.RecupVar("[SMTPServer]"), 25)
            With mailSender

                .Credentials = New Net.NetworkCredential(Commun.RecupVar("[AdminScriptLogin]"), Commun.RecupVar("[AdminScriptPassword]"), "igbmc")
                '.UseDefaultCredentials = false
                .EnableSsl = False

            End With

            'Ajouter un destinataire caché
            'message.Bcc.Add(fromAddress)

            message.Subject = subject
            message.IsBodyHtml = False

            'detecter si le mail est en HTML
            If Left(body, 21) = "<!DOCTYPE HTML PUBLIC" Then
                message.IsBodyHtml = True
            End If
            message.Body = body

            If Not attachmentString = "" Then
                Dim msgAttach As New Attachment(attachmentString)
                message.Attachments.Add(msgAttach)
            End If

            mailSender.Send(message)

        Catch ex As Exception
            Commun.Journal("ERREUR : echec de l'envoi de mail")
        End Try
    End Sub
    Shared Sub getFileFTP(ByVal cheminServeur As String, ByVal cheminLocal As String)
        Dim persoFTP As FTPP = New FTPP()
        persoFTP.m_FTPSite = "ftp://" & Commun.RecupVar("[FCCserver]")
        persoFTP.UserName = Commun.RecupVar("[FCCAccount]")
        persoFTP.Password = Commun.RecupVar("[FCCPassword]")

        Try
            persoFTP.GetFile(cheminServeur, cheminLocal)
        Catch ex As Exception
            Commun.Journal("ERREUR : FTP : Recuperation du fichier : " & cheminLocal & " : " & ex.Message, True)
            Application.Exit()
        End Try

    End Sub
    Shared Sub UploadFileFTP(ByVal cheminLocal As String, ByVal cheminServeur As String)
        Dim persoFTP As FTPP = New FTPP()
        persoFTP.m_FTPSite = "ftp://" & Commun.RecupVar("[FCCserver]")
        persoFTP.UserName = Commun.RecupVar("[FCCAccount]")
        persoFTP.Password = Commun.RecupVar("[FCCPassword]")

        Try
            persoFTP.UploadFile(cheminLocal, cheminServeur)
        Catch ex As Exception
            Commun.Journal("ERREUR : FTP : Envoi du fichier : " & cheminLocal & " : " & ex.Message, True)
            Application.Exit()
        End Try
    End Sub
    ''' <summary>
    ''' Creation des groupes d'equipeinfo (_eq, _eq_ADMIN, _dest)
    ''' </summary>
    ''' <param name="groupeSAMAccountSansEQ">Nom court de l'equipeinfo</param>
    ''' <param name="repzone">Chemin du repertoire de la zone.</param>
    ''' <param name="nomComplet">Nom complet de la zone.</param>
    ''' <remarks></remarks>
    Shared Sub CreationGroupesEquipeInfo(ByVal groupeSAMAccountSansEQ As String, ByVal repzone As String, ByVal nomComplet As String)
        'creation des groupes de l'equipe
        Dim ou As String = "OU=Equipes"
        Dim nomGroup As String = groupeSAMAccountSansEQ & "_eq"
        nomComplet = SansAccent(nomComplet)


        For i = 0 To 2

            If i = 1 Then
                ou = "OU=Admin_Zones,OU=Equipes"
                nomGroup = groupeSAMAccountSansEQ & "_eq_ADMIN"

            End If

            If i = 2 Then
                ou = "OU=Gestion_Destinations,OU=Equipes"
                nomGroup = groupeSAMAccountSansEQ & "_dest"
            End If

            If Not DirectoryEntry.Exists("LDAP://cn=" & nomGroup & "," & ou & ",OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr") Then

                Try
                    Dim groupe As DirectoryEntries = New DirectoryEntry("LDAP://" & ou & ",OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr").Children
                    Dim newgroup As DirectoryEntry = groupe.Add("CN=" & nomGroup, "group")
                    newgroup.Properties("SAMAccountName").Add(nomGroup)
                    If i = 0 Then
                        newgroup.Properties("DisplayName").Add(nomComplet)
                    End If
                    newgroup.Properties("groupType").Add(-2147483640) 'Groupe universel
                    If i <> 2 Then
                        Dim gid As Integer = GIDMiniGroup()
                        newgroup.Properties("gidnumber").Value = gid
                    End If
                    newgroup.CommitChanges()
                    newgroup.Close()
                    newgroup.Dispose()
                    newgroup = Nothing
                    groupe = Nothing

                    Commun.Journal("Succes : GroupExists : Creation du groupe : " & nomGroup)
                Catch ex As Exception
                    Commun.Journal("ERREUR : GroupExists : Creation du groupe : " & nomGroup & " : " & ex.Message, True)
                    End
                End Try
            End If

        Next i
    End Sub
    Shared Function SansAccent(ByVal Chaine As String) As String
        Dim aOctets() As Byte = System.Text.Encoding.GetEncoding(1251).GetBytes(Chaine) 'converti en byte la chaine avec accents
        Dim sEnleverAccents As String = System.Text.Encoding.ASCII.GetString(aOctets) 'converti en string la chaine sans accents
        Return sEnleverAccents
    End Function
    Shared Function GIDMiniGroup() As Integer
        Dim resultat As Integer
        Try
            Dim AD As DirectoryEntry = New DirectoryEntry("LDAP://OU=Equipes,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr")
            Dim ResultFields() As String = {"gidNumber"}
            Dim searcherAD As DirectorySearcher = New DirectorySearcher(AD)
            Dim CtrlUIDNumberLibre As Boolean = False
            Dim i As Integer = -1
            Dim nouveauUIDNumber As Integer

            searcherAD.Filter = "(&(objectClass=group))"
            searcherAD.SearchScope = SearchScope.Subtree
            searcherAD.PropertiesToLoad.AddRange(ResultFields)
            Dim tabUIDN() As Integer
            Dim resultLdapU As SearchResultCollection = searcherAD.FindAll()
            For Each result As SearchResult In resultLdapU
                If result.Properties.Contains("gidNumber") Then
                    i += 1
                    ReDim Preserve tabUIDN(i)
                    tabUIDN(i) = result.Properties("gidNumber")(0).ToString
                End If
            Next result
            resultLdapU.Dispose()
            resultLdapU = Nothing
            searcherAD.Dispose()
            searcherAD = Nothing
            AD.Close()
            AD.Dispose()
            AD = Nothing

            Dim k = 3500
            Do Until CtrlUIDNumberLibre = True
                k += 1
                CtrlUIDNumberLibre = True
                For j = 0 To UBound(tabUIDN)
                    If tabUIDN(j) = k Then
                        CtrlUIDNumberLibre = False
                    End If
                Next j
            Loop
            resultat = k
        Catch ex As Exception
            Commun.Journal("ERREUR : GIDMini : Récuération d'un GID libre : " & ex.Message, True)
        End Try
        Return resultat
    End Function
    Shared Function PrimaryGroupId(ByVal objectSID)
        Dim GroupIDBytes() As Byte = DirectCast(objectSID, Byte())
        Dim GroupSID As New Security.Principal.SecurityIdentifier(GroupIDBytes, 0)
        Dim SplitSID() As String = GroupSID.Value.Split("-"c)
        Dim RID As Integer = CInt(SplitSID(SplitSID.Length - 1))
        Return RID
    End Function
    ''' <summary>
    ''' <para>Definition des Attributs d'un directoryentry dans l'AD ou le LDAP.</para>
    ''' <para>ATTENTION: Le Commitchange doit etre fait appres l'appel de la function.</para>
    ''' </summary>
    ''' <param name="de">Directoryentry a modifier</param>
    ''' <param name="pName">Nom de l'attribut.</param>
    ''' <param name="pValue">Valeur de l'attribut.</param>
    ''' <param name="addValueto">Ajouter la nouvelle valeur a l'ancienne.</param>
    ''' <remarks></remarks>
    Public Shared Sub SetADLDAPProperty(ByVal de As DirectoryEntry, ByVal pName As String, ByVal pValue As String, Optional ByVal addValueto As Boolean = False)
        'ATTENTION LE COMMITCHANGES() DOIT ETRE FAIT APRES L'APPEL DE LA FONCTION

        If Not pValue Is Nothing Then
            'Check to see if the DirectoryEntry contains this property already 
            If de.Properties.Contains(pName) Then 'The DE contains this property 
                'Update the properties value 
                If addValueto = False Then
                    de.Properties(pName)(0) = pValue
                Else
                    de.Properties(pName)(0) += pValue
                End If
            Else    'Property doesnt exist 
                'Add the property and set it's value 
                de.Properties(pName).Add(pValue)
            End If
        End If
    End Sub
    ''' <summary>
    ''' <para>Transforme un SAMAccount ou l'inverse.</para>
    ''' </summary>
    ''' <param name="loginOuCN">login ou CN (sans "LDAP://").</param>
    ''' <remarks></remarks>
    Shared Function TransformeSAMACCOUNTenCN(ByVal loginOuCN As String) As String

        Dim resultat As String = ""
        If loginOuCN <> "" Then


            If Left(loginOuCN, 3) = "CN=" Then
                If DirectoryEntry.Exists("LDAP://" & loginOuCN) Then
                    Dim objectAD As DirectoryEntry = New DirectoryEntry("LDAP://" & loginOuCN)
                    resultat = objectAD.Properties("SAMAccountName").Value
                    objectAD.Close()
                    objectAD.Dispose()
                    objectAD = Nothing
                End If

            Else
                Dim Ldap As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")
                Dim searcher As DirectorySearcher = New DirectorySearcher(Ldap)
                searcher.Filter = "(& (SAMAccountName=" & loginOuCN & "))"
                Dim result As SearchResult = searcher.FindOne()
                If Not result Is Nothing Then
                    Dim pathCN As String = Replace(result.Path, "LDAP://", "")
                    resultat = pathCN
                End If
                searcher.Dispose()
                searcher = Nothing
                result = Nothing
                Ldap.Close()
                Ldap.Dispose()
                Ldap = Nothing
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

        user = Replace(user, "LDAP://", "")
        If Strings.Left(groupe, 3) <> "CN=" Then
            groupe = Commun.TransformeSAMACCOUNTenCN(groupe)
        End If
        If Strings.Left(user, 3) <> "CN=" Then
            user = Commun.TransformeSAMACCOUNTenCN(user)
        End If
        Dim result As Boolean = False


        If DirectoryEntry.Exists("LDAP://" & user) Then
            Dim userEntry As DirectoryEntry = New DirectoryEntry("LDAP://" & user)

            If DirectoryEntry.Exists("LDAP://" & groupe) Then
                Dim groupEntry As DirectoryEntry = New DirectoryEntry("LDAP://" & groupe)
                Dim ridGroup As Integer = Commun.PrimaryGroupId(groupEntry.Properties("objectSid").Value)
                groupEntry.Close()
                groupEntry.Dispose()
                groupEntry = Nothing

                If userEntry.Properties("primaryGroupID").Value = ridGroup Then
                    result = True
                    GoTo sortie
                End If

            End If
            For i = 0 To userEntry.Properties("memberOf").Count() - 1
                Dim temp As String = userEntry.Properties("memberOf")(i).ToString
                If temp = groupe Then
                    result = True
                    Exit For
                End If
            Next i
sortie:

            userEntry.Close()
            userEntry.Dispose()
            userEntry = Nothing
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
    ''' <remarks></remarks>
    Shared Sub AddRemoveADGroup(ByVal login As String, ByVal group As String, ByVal action As String) 'Action : "Add" ou "Remove"
        'La fonction gere le login ou le DN pour le login ou le group
        'controle l'appartenance au groupe avant l'action. Si l'utilisateur doit etre ajouté, verification si l'utilisateur est deja membre du groupe

        login = Replace(login, "LDAP://", "")
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
                    Dim GroupPath As New DirectoryEntry("LDAP://" & CNgroup)
                    GroupPath.Invoke(action, New Object() {"LDAP://" & CNuser})
                    GroupPath.Close()
                    GroupPath.Dispose()
                    GroupPath = Nothing
                    'GC.Collect()
                End If
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : AddRemoveADGroup : " & action & " : " & login & "-" & group & " : " & ex.Message, True)
        End Try
    End Sub
    'Shared Sub CreateEquipeAD(ByVal nomDuGroupe As String, ByVal CNParent As String)
    '    Dim objDomGroup As DirectoryEntry = New DirectoryEntry("LDAP://" & CNParent)
    '    Dim dirSearcher As DirectorySearcher = New DirectorySearcher(objDomGroup)
    '    Dim nouveauGroupe As DirectoryEntries = objDomGroup.Children
    '    Dim newgroup As DirectoryEntry = nouveauGroupe.Add("CN=" & nomDuGroupe, "group")
    '    newgroup.Properties("SAMAccountName").Add(nomDuGroupe)
    '    newgroup.CommitChanges()
    '    newgroup.Close()
    '    newgroup.Dispose()
    '    newgroup = Nothing
    '    dirSearcher.Dispose()
    '    dirSearcher = Nothing
    '    objDomGroup.Close()
    '    objDomGroup.Dispose()
    '    objDomGroup = Nothing
    'End Sub

    ''' <summary>
    ''' <para>Renvoi un tableau contenant les membres d'un groupe.</para>
    ''' </summary>
    ''' <param name="GroupSamAccount">SAMAccountName du groupe.</param>
    ''' <param name="avecMembresGroupePrincipal">True ou False pour ajouter les membres qui ont ce groupe comme groupe principal.</param>
    ''' <remarks></remarks>
    Shared Function MembresDuGroupe(ByVal GroupSamAccount As String, Optional ByVal avecMembresGroupePrincipal As Boolean = False) As Array
        Dim tabMembres As String()
        Dim i As Integer = -1
        Try
            'Recupération de l'attribut Member pour le mettre dans le tableau des resultats
            Dim AD As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")
            Dim searcherGroup As DirectorySearcher = New DirectorySearcher(AD)
            searcherGroup.Filter = "(&(objectClass=group) (SAMAccountName=" & GroupSamAccount & "))"
            searcherGroup.PropertiesToLoad.Add("Member")
            searcherGroup.PropertiesToLoad.Add("objectSid")
            Dim resultGroup As SearchResult = searcherGroup.FindOne()
            Dim Group As DirectoryEntry = resultGroup.GetDirectoryEntry
            For Each unMembre In Group.Properties("member")
                i += 1
                ReDim Preserve tabMembres(i)
                tabMembres(i) = unMembre.ToString
            Next unMembre
            searcherGroup.Dispose()
            searcherGroup = Nothing

            If avecMembresGroupePrincipal = True Then

                Dim RID As Integer = PrimaryGroupId(Group.Properties("objectSid")(0))

                Dim searcher As DirectorySearcher = New DirectorySearcher(AD)
                searcher.Filter = "(&(objectClass=user) (primaryGroupID=" & RID & "))"
                Dim resultCollection As SearchResultCollection = searcher.FindAll()
                For Each result As SearchResult In resultCollection
                    i += 1
                    ReDim Preserve tabMembres(i)
                    tabMembres(i) = Replace(result.Path, "LDAP://", "")
                Next result
                resultCollection.Dispose()
                resultCollection = Nothing
                searcher.Dispose()
                searcher = Nothing
            End If
            AD.Close()
            AD.Dispose()
            AD = Nothing

            Return tabMembres

        Catch ex As Exception

        End Try
    End Function
    ''' <summary>
    ''' <para>Renvoi un tableau contenant les groupes auquel un autre groupe appartient.</para>
    ''' </summary>
    ''' <param name="SamAccount">SAMAccountName du groupe.</param>
    ''' <remarks></remarks>
    Shared Function GroupeMembresDe(ByVal SamAccount As String) As Array
        Dim tabMembres As String() = Nothing
        Dim i As Integer = -1
        Try
            'Recupération de l'attribut Member pour le mettre dans le tableau des resultats
            Dim AD As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")
            Dim searcherGroup As DirectorySearcher = New DirectorySearcher(AD)
            searcherGroup.Filter = "(&(objectClass=group) (SAMAccountName=" & SamAccount & "))"
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
                searcherGroup.Dispose()
                searcherGroup = Nothing

                AD.Close()
                AD.Dispose()
                AD = Nothing
            End If

            Return tabMembres

        Catch ex As Exception

        End Try
    End Function
    Shared Function IsMembreEquipeAdministratif(ByVal user As DirectoryEntry) As Boolean
        Dim tabEquipeADM As String() = Split(Commun.RecupVar("[EquipesAdministratives]"), ",")
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
    Shared Sub GestionGroupeUserActive(ByVal dirEntry1 As DirectoryEntry)

        Dim CNUser As String = Replace(dirEntry1.Path, "LDAP://", "")
        Try
            'ajout au groupe utilisa. du domaine 
            If Commun.AppartientGroup(CNUser, "Utilisa. du domaine") = False Then
                Commun.AddRemoveADGroup(CNUser, "Utilisa. du domaine", "Add")
                'Definition du groupe utilisa. du domaine comme groupe principal
                dirEntry1.Properties("primaryGroupID").Value = Commun.PrimaryGroupId(Commun.FindAttribut("Utilisa. du domaine", "objectSid"))
                dirEntry1.CommitChanges()

            End If



            'Retrait du groupe G_Domain_DisableOpenSession
            If Commun.AppartientGroup(CNUser, "G_Domain_DisableOpenSession") = True Then
                Commun.AddRemoveADGroup(CNUser, "G_Domain_DisableOpenSession", "Remove")
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : réactivation du compte(modification des groupes) : " & CNUser, True)
        End Try
    End Sub
    Shared Sub GestionGroupeUserDesactive(ByVal dirEntry1 As DirectoryEntry)

        Dim CNUser As String = Replace(dirEntry1.Path, "LDAP://", "")

        Try

            'Ajout au groupe G_Domaine_DisableOpenSession
            If Commun.AppartientGroup(CNUser, "G_Domain_DisableOpenSession") = False Then
                Commun.AddRemoveADGroup(CNUser, "G_Domain_DisableOpenSession", "Add")
                'definition du groupe G_Domain_DisableOpenSession comme groupe principal
                dirEntry1.Properties("primaryGroupID").Value = Commun.PrimaryGroupId(Commun.FindAttribut("G_Domain_DisableOpenSession", "objectSid"))
                dirEntry1.CommitChanges()
            End If

            'Retrait du groupe Utilisa. du domaine
            If Commun.AppartientGroup(CNUser, "Utilisa. du domaine") = True Then
                Commun.AddRemoveADGroup(CNUser, "Utilisa. du domaine", "Remove")
            End If

        Catch ex As Exception
            Commun.Journal("ERREUR : désactivation du compte(modification des groupes) : " & Replace(CNUser, "LDAP://", ""), True)
        End Try
    End Sub
    ''' <summary>
    ''' <para>Renvoi un attribut AD d'un obbjet.</para>
    ''' <para>Renvoi Nothing si l'attribut n'est pas defini.</para>
    ''' </summary>
    ''' <param name="samaccountName">SAMAccountName de l'objet.</param>
    ''' <param name="attribut">attribut a récupérer.</param>
    ''' <remarks></remarks>
    Shared Function FindAttribut(ByVal samaccountName As String, ByVal attribut As String)
        Dim resultat = Nothing

        Dim CNObject As String = TransformeSAMACCOUNTenCN(samaccountName)
        If CNObject <> "" Then
            If DirectoryEntry.Exists("LDAP://" & CNObject) Then
                Dim objectAD As DirectoryEntry = New DirectoryEntry("LDAP://" & CNObject)
                If objectAD.Properties.Contains(attribut) Then
                    resultat = objectAD.Properties(attribut).Value
                End If
                objectAD.Close()
                objectAD.Dispose()
                objectAD = Nothing
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
    Shared Sub CreateTabExceptionUser()
        Try
            Dim monStreamReader As StreamReader = New StreamReader("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\UsersExceptions.txt")
            Dim ligne As String
            Dim h As Integer = -1
            Do
                ligne = monStreamReader.ReadLine()
                If Not ligne Is Nothing Then
                    Dim tabLigneExceptUser As String() = Split(ligne, ",")
                    Dim dateExpire As DateTime = Convert.ToDateTime(tabLigneExceptUser(1))
                    If dateExpire >= Now.Date Then
                        h += 1
                        ReDim Preserve tabExcepUser(h)
                        tabExcepUser(h) = LCase(tabLigneExceptUser(0) & "," & tabLigneExceptUser(2))
                    End If
                End If
            Loop Until ligne Is Nothing
            monStreamReader.Close()
            monStreamReader.Dispose()

        Catch ex As Exception
            Commun.Journal("ERREUR : CreateTabExceptionUser : " & ex.Message, True)
        End Try
    End Sub
End Class