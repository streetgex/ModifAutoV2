Imports System.DirectoryServices
Imports System.IO
Imports System.Net.Mail
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Collections.ObjectModel



Public Class Creation
    Shared tabDatabase(1, 35) As String

    Sub createCompte(ByVal usrNom As String, ByVal usrPrenom As String, ByVal usrDest As String, usrID As String, ByVal usrLogin As String, ByVal genre As String, ByVal Optional usrListeDiff As String = "")

        NbUserDatabaseExchange()

        ' TRAITEMENT DU FICHIER DE CREATION DE COMPTES




        Dim objUser As DirectoryEntry
        Dim objOUUtilisateurs As DirectoryEntry


        Dim EqDescr As String = " "


        'Si l'utilisateur existe par rapport a l'employeeID et est dans l'ou Utilisateurs, on prend la ligne suivante du fichier
        If EmployeeIDExist(usrID) = True And DirectoryEntry.Exists("LDAP://CN=" & usrPrenom & " " & usrNom & "," & RecupDataini.RecupVar("[OUUtilisateursActifs]")) Then Exit Sub

        'fonction deja appelée dans la creation du fichier json
        If usrLogin = "" Then
            usrLogin = DetermineLogin(usrPrenom, usrNom, usrID)
        End If


        Dim usrPasswd As String = RandomPassword.Generate(8, 10)
        Dim usrEquipeinfo As String = Commun.RecupEquipeinfo(usrDest)
        If usrEquipeinfo = Nothing Then
            usrEquipeinfo = "externe"
        End If
        Dim usrChefEquipeinfoCN As String = Commun.FindAttribut(usrEquipeinfo & "_eq", "managedBy")
        Dim usrPathEquipeinfo As String = Commun.FindAttribut(usrEquipeinfo & "_eq", "url")
        Dim aliasSMTP As String() = Commun.DetermineAliasLibre(usrPrenom, usrNom, usrID)

        Commun.Journal("Debut de Creation de compte : " & usrNom & "," & usrPrenom & "," & usrDest & "," & usrID & "," & usrLogin & "," & usrListeDiff)

        EqDescr = Commun.FindAttribut(usrDest & " grp", "Description")

        Try
            'verification si un compte provisoire exsite
            Dim tempCtrlPathExist As String = UserExist(usrPrenom & " " & usrNom & " (Provisoire)")


            objOUUtilisateurs = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursActifs]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)



            If tempCtrlPathExist = "" Then

                objUser = objOUUtilisateurs.Children.Add("CN=" & usrPrenom & " " & usrNom, "user")
                objUser.Properties("SAMAccountName").Value = usrLogin
                objUser.Properties("userPrincipalName").Value = usrLogin & "@igbmc.fr"
                Commun.AppliquerChangement(objUser)

                objUser.Properties("Comment").Value += "Créé le: " & Strings.Left(CStr(Now), 10) & vbCrLf
                Commun.AppliquerChangement(objUser)
            Else

                objUser = New DirectoryEntry(tempCtrlPathExist, Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)

                Try
                    If objUser.Path = "LDAP://CN=" & usrPrenom & " " & usrNom & " (Provisoire)," & RecupDataini.RecupVar("[OUUtilisateursProvisoires]") Then

                        objUser.Rename("CN=" & usrPrenom & " " & usrNom)
                        Commun.AppliquerChangement(objUser)
                        objUser.Properties("Comment").Value += "Transformé le: " & Strings.Left(CStr(Now), 10) & vbCrLf
                        objUser.Properties("AccountExpires").Value = 0
                        Commun.AppliquerChangement(objUser)

                    End If

                    objUser.MoveTo(objOUUtilisateurs)
                    Commun.AppliquerChangement(objUser)

                    Commun.SetADLDAPProperty(objUser, "accountDeletionDate", "")
                    objUser.Properties("accountDeletionDT").Value = Nothing
                    objUser.Properties("accountDeactivationDT").Value = Nothing
                    objUser.Properties("description").Value = Nothing
                    objUser.Properties("SAMAccountName").Value = usrLogin
                    objUser.Properties("userPrincipalName").Value = usrLogin & "@igbmc.fr"
                    Commun.AppliquerChangement(objUser)
                    Commun.Journal("Creation de compte : le compte existait déja, il a été modifié : " & usrLogin, True)

                Catch ex As Exception
                    Commun.Journal("ERREUR : Creation de compte : Modification du compte provisoire, une erreur s'est produite : " & usrLogin & " : " & ex.Message, True)
                End Try


            End If

            objUser.Properties("SN").Value = usrNom
            objUser.Properties("givenName").Value = usrPrenom
            objUser.Properties("Displayname").Value = usrPrenom & " " & usrNom
            objUser.Properties("displayNamePrintable").Value = usrNom & " " & usrPrenom
            'Definition du genre: Homme Femme
            objUser.Properties("extensionAttribute2").Add(genre)
            Commun.AppliquerChangement(objUser)
            objOUUtilisateurs.Close()
            objOUUtilisateurs.Dispose()
            objOUUtilisateurs = Nothing
        Catch ex As Exception
            Commun.Journal("ERREUR : Creation de compte : Une erreur s'est produite : " & usrLogin & " : " & ex.Message, True)
        End Try

        Try
            objUser.Invoke("SetPassword", New Object() {usrPasswd})
            Commun.AppliquerChangement(objUser)

            objUser.Properties("ObjectClass").Add("posixaccount")

            objUser.Properties("mail").Value = usrLogin & "@igbmc.fr"
            objUser.Properties("Department").Value = EqDescr
            Commun.AppliquerChangement(objUser)


            objUser.Properties("EmployeeID").Value = usrID
            objUser.Properties("company").Value = "IGBMC"
            objUser.Properties("DepartmentNumber").Value = usrDest
            Commun.AppliquerChangement(objUser)
        Catch
            Commun.Journal("ERREUR : Creation de compte : Attributs : " & usrLogin, True)
        End Try

        'Ajout de la valeur posixAccount à la class objectClass du compte
        Try
            objUser.Properties("ObjectClass").Add("posixaccount")
            Commun.AppliquerChangement(objUser)
        Catch
            Commun.Journal("ERREUR : Ajout PosixAccount au compte : Attributs : " & usrLogin, True)
        End Try

        Try
            'Determination des données Unix
            'Recherche du gid du groupe de l'equipe de l'utilisateur.
            Dim GIDNumber As String = Commun.FindAttribut(usrEquipeinfo & "_eq", "gidNumber")

            'Recherche
            Dim unixHomeDirectory As String = Replace(Replace(usrPathEquipeinfo, "\\", "\"), "\", "/") & "/" & usrLogin
            If InStr(usrPathEquipeinfo, "labo4") > 0 Then
                unixHomeDirectory = Replace(Replace(usrPathEquipeinfo, "\\", "\"), "\", "/") & "/" & usrEquipeinfo & "/" & usrLogin
            End If

            'Creation du compte LDAP Unix
            Dim uidNumber As String = Commun.UIDNumberMini() ' (usrNom, usrPrenom, usrLogin, usrPasswd, unixHomeDirectory, GIDNumber)
            '5 valeurs unix a integrer dans l'AD : gidNumber,uidNumber,loginShell,homeDirectory,uid
            Commun.SetADLDAPProperty(objUser, "gidNumber", GIDNumber)
            'objUser.Properties("uidNumber").Value = uidNumber
            Commun.SetADLDAPProperty(objUser, "uidNumber", uidNumber)
            'objUser.Properties("loginShell").Value = "/bin/tcsh"
            Commun.SetADLDAPProperty(objUser, "loginShell", "/bin/bash")
            'objUser.Properties("unixHomeDirectory").Value = unixHomeDirectory
            Commun.SetADLDAPProperty(objUser, "unixHomeDirectory", unixHomeDirectory)
            Commun.SetADLDAPProperty(objUser, "uid", usrLogin)
            Commun.AppliquerChangement(objUser)

        Catch
            Commun.Journal("ERREUR : Creation de l'utilisateur Unix : " & usrLogin, True)
        End Try

        'AJOUT DU GROUPE EQUIPE COMPTABLE
        If EqDescr <> " " And EqDescr <> "" And EqDescr <> Nothing Then
            Try
                Commun.AddRemoveADGroup(usrLogin, usrDest & " grp", "Add")
            Catch ex As Exception
                Commun.Journal("ERREUR : Ajout a l'equipe comptable : " & usrLogin & " : " & ex.Message, True)
            End Try
        End If

        'ajout au groupe administratif en fonction de l'equipe info
        Try

            Dim membreEquipeADM As Boolean = Commun.IsMembreEquipeAdministratif(objUser)

            If membreEquipeADM = True Then
                Commun.AddRemoveADGroup(usrLogin, "G_ADMINISTRATIF_Users", "Add")
            End If


        Catch e As Exception
            Commun.Journal("ERREUR : ajout à l'equipe ADMINISTRATIVE de l'utilisateur : " & e.Message & " : " & usrLogin, True)
        End Try

        'Determiner le Flag sur "le mot de passe n'expire jamais"
        Try
            Dim objUserNT = GetObject("WinNT://igbmc/" & usrLogin)
            objUserNT.Put("userFlags", &H10000)
            objUserNT.setinfo()
            objUserNT = Nothing
        Catch
        End Try

        'creation de la boite mail Exchange
        Try
            Dim petiteDB As String = TrouverPetiteDatabase()
            commandePWSMailbox(usrLogin, petiteDB)
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Creation du compte mail : " & e.Message & " : " & usrLogin, True)
        End Try

        'Creation des alias de mails


        Try

            Dim adress As New List(Of String)
            adress.Add("smtp:" & aliasSMTP(0) & "@igbmc.fr")
            'adress.Add("smtp:" & aliasSMTP(0) & "@igbmc.u-strasbg.fr")
            adress.Add("SMTP:" & usrLogin & "@igbmc.fr")
            'adress.Add("smtp:" & usrLogin & "@igbmc.u-strasbg.fr")
            objUser.Properties("proxyAddresses").Value = adress.ToArray()
            Commun.AppliquerChangement(objUser)
            Dim a As Integer = UBound(aliasSMTP)
            If a > 0 Then
                For a = 1 To a
                    If aliasSMTP(a) <> usrLogin Then
                        adress.Add("smtp:" & aliasSMTP(a) & "@igbmc.fr")
                        'adress.Add("smtp:" & aliasSMTP(a) & "@igbmc.u-strasbg.fr")
                        objUser.Properties("proxyAddresses").Value = adress.ToArray()
                        Commun.AppliquerChangement(objUser)
                    End If
                Next
            End If

            objUser.Properties("msExchExtensionAttribute16").Value = aliasSMTP(0) & "@igbmc.fr"
            Commun.AppliquerChangement(objUser)
            Commun.Journal("Creation des alias la boite mail Réussie: " & usrLogin)
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Creation des alias mail : " & e.Message & " : " & usrLogin, True)
        End Try

        'Ajout a la liste de diffusion "annouce"
        Try
            Commun.AddRemoveADGroup(usrLogin, "Announce", "Add")
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Ajout liste de diffusions ANNOUCE : " & e.Message & " : " & usrLogin, True)
        End Try

        'Ajout a la liste de diffusion "ce"
        Try
            Commun.AddRemoveADGroup(usrLogin, "ce", "Add")
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Ajout liste de diffusions CE : " & e.Message & " : " & usrLogin, True)
        End Try

        'Ajout a la liste de diffusion "scientific info"
        Try
            Commun.AddRemoveADGroup(usrLogin, "Scientific Info", "Add")
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Ajout liste de diffusions SCIENTIFIC INFO : " & e.Message & " : " & usrLogin, True)
        End Try

        'Ajout a la liste de diffusion "spb info"
        Try
            Commun.AddRemoveADGroup(usrLogin, "spb-info", "Add")
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Ajout liste de diffusions spb-info : " & e.Message & " : " & usrLogin, True)
        End Try

        'Ajout a la liste de diffusion "phd" ou "postdoc"
        Try
            If usrListeDiff = "phd" Or usrListeDiff = "postdoc" Then
                Commun.AddRemoveADGroup(usrLogin, usrListeDiff, "Add")
            End If
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Ajout liste de diffusions PHD ou POSTDOC : " & e.Message & " : " & usrLogin, True)
        End Try

        'Ajout au groupe autorisant l'acces par VPN
        Try
            Commun.AddRemoveADGroup(usrLogin, "G_Acces VPN-SSTP", "Add")
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Ajout groupe VPN : " & e.Message & " : " & usrLogin, True)
        End Try

        'Ajout au groupe autorisant l'acces par VPN
        Try
            Commun.AddRemoveADGroup(usrLogin, "LicenceO365ProPlus", "Add")
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Ajout groupe pour les licenses Office 365 Pro Plus : " & e.Message & " : " & usrLogin, True)
        End Try

        'creation de l'attribut "msExchUsageLocation" synchronisé avec l'attribut "UsageLocation" pour Microsoft Azure
        Try
            Commun.SetADLDAPProperty(objUser, "msExchUsageLocation", "FR")
            Commun.AppliquerChangement(objUser)
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : attribut ""msExchUsageLocation"" : " & e.Message & " : " & usrLogin, True)
        End Try

        'Enregistrement du login et de l'alias mail dans l'historique
        Commun.ajoutAliasFichierHisto(aliasSMTP(0), usrID)
        Commun.ajoutAliasFichierHisto(usrLogin, usrID)
        'Creer l'alias sur IGBMCSERVICES

        'Dim textmail As String = corpMailCreation(usrPrenom, usrNom, usrLogin, usrPasswd, aliasSMTP, usrPathEquipeinfo)
        Dim textMail = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN""><HTML><HEAD><META HTTP-EQUIV=""CONTENT-TYPE"" " &
        "CONTENT=""text/html; charset=windows-1252""><TITLE></TITLE><STYLE TYPE=""text/css""><!--P { margin-bottom: 0.21cm }--></STYLE></HEAD><BODY LANG=""fr-FR"" DIR=""LTR"">Bienvenue a l'IGBMC " _
                    & usrPrenom & " " & usrNom & " (compte ouvert le " & DateTime.Today.ToString("dd/MM/yyyy") & ")</FONT></FONT></P><P ALIGN=CENTER STYLE=""margin-bottom: 0cm""><BR></P><TABLE WIDTH=642" _
                    & " BORDER=1 BORDERCOLOR=""#000000"" CELLPADDING=0 CELLSPACING=0><COL WIDTH=150><COL WIDTH=490><TR><TD COLSPAN=2 WIDTH=640 VALIGN=TOP><center>nom d'utilisateur: <B>" _
                    & usrLogin & "</B><BR>mot de passe (" & Len(usrPasswd) & " caract&egrave;res): <B>" & usrPasswd & "</B><BR>(&Agrave; changer sur https://password.igbmc.fr)<BR><BR>Votre adresse e-mail est :<BR><B>" _
                    & aliasSMTP(0) & "@igbmc.fr</B><BR>ou bien " & usrLogin & "@igbmc.fr</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Messagerie et agenda</center></TD><TD WIDTH=490><center>" _
                    & "Vous pouvez acc&eacute;der &agrave; vos mails et &agrave; votre agenda personnel via l'adresse https://igbmcmail.igbmc.fr/owa <BR></center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>" _
                    & "Commandes de cellules/s&eacute;quen&ccedil;age</center></TD><TD WIDTH=490><center>Il faut demander un compte &agrave; Michel Offner (tel: 3277) ou Tony Moutaux (tel: 3395)</center></TD></TR><TR VALIGN=TOP>" _
                    & "<TD WIDTH=150><center>Espace disque</center></TD><TD WIDTH=490><center>Vous disposez d'un espace de stockage sauvegard&eacute; auquel vous pouvez acc&eacute;der via l'adresse " _
                    & usrPathEquipeinfo & " .</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Qui contacter en cas de probl&egrave;me?</center></TD><TD WIDTH=490>" _
                    & "<center>Envoyez un mail &agrave; <B>helpdesk@igbmc.fr</B></center></TD></TR></TABLE><BR><TABLE WIDTH=642 BORDER=1 BORDERCOLOR=""#000000"" CELLPADDING=0 CELLSPACING=0><COL WIDTH=150>" _
                    & "<COL WIDTH=490><TR><TD COLSPAN=2 WIDTH=640 VALIGN=TOP><center>Your login: <B>" & usrLogin & "</B><BR>Your password: <B>" & usrPasswd & "</B><BR>(Please change it on https://password.igbmc.fr)" _
                    & "<BR><BR>Your email address Is:<BR><B>" & aliasSMTP(0) & "@igbmc.fr</B><BR>Or " & usrLogin & "@igbmc.fr</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Email and Calendar</center></TD><TD WIDTH=490><center>" _
                    & "You can access to your mail And to the shared calendar at https://igbmcmail.igbmc.fr/owa <BR></center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Orders of cells, sequencing</center></TD><TD WIDTH=490><center>" _
                    & "If you need to order cells, you must ask for an account to Michel Offner (3277) Or Tony Moutaux (3395).</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Disk space</center></TD><TD WIDTH=490><center>" _
                    & "You have a backed up data storage space that you can access by: " & usrPathEquipeinfo & "</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Contact</center></TD><TD WIDTH=490><center>Send an email to: <B>helpdesk@igbmc.fr</B>" _
                    & "</center></TD></TR></TABLE></BODY></HTML>"
        Commun.SendEmail(RecupDataini.RecupVar("[AdminScriptLogin]") & "@igbmc.fr", "serviceinfo@igbmc.fr", "[Creation de compte] " & usrPrenom & " " & usrNom, textMail)
        Commun.Journal("Fin de Creation de compte : " & usrLogin)
        Json.SendJson("login=" & usrLogin & "&domain=%40igbmc.fr&alias=" & aliasSMTP(0), "persons/" & usrID & "/email", "AD", "POST")
        objUser.Properties("msExchExtensionAttribute16").Value = LCase(aliasSMTP(0)) & "@igbmc.fr"
        Commun.AppliquerChangement(objUser)
        'End If

        objUser.Close()
        objUser.Dispose()
        objUser = Nothing
    End Sub

    Shared Sub NbUserDatabaseExchange()

        Dim Ldap As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursActifs]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
        Dim dirSearcher As DirectorySearcher = New DirectorySearcher(Ldap)
        dirSearcher.PageSize = 2000
        dirSearcher.SearchScope = SearchScope.Subtree

        tabDatabase(0, 0) = "DB-E12d-users"
        tabDatabase(0, 1) = "DB-E12e-users"
        tabDatabase(0, 2) = "DB-E12f-users"
        tabDatabase(0, 3) = "DB-E12g-users"
        tabDatabase(0, 4) = "DB-E12h-users"
        tabDatabase(0, 5) = "DB-E12i-users"
        tabDatabase(0, 6) = "DB-F21d-users"
        tabDatabase(0, 7) = "DB-F21e-users"
        tabDatabase(0, 8) = "DB-F21f-users"
        tabDatabase(0, 9) = "DB-F21g-users"
        tabDatabase(0, 10) = "DB-F21h-users"
        tabDatabase(0, 11) = "DB-F21i-users"
        tabDatabase(0, 12) = "DB-G31d-Users"
        tabDatabase(0, 13) = "DB-G31e-Users"
        tabDatabase(0, 14) = "DB-G31f-Users"
        tabDatabase(0, 15) = "DB-G31g-Users"
        tabDatabase(0, 16) = "DB-G31h-Users"
        tabDatabase(0, 17) = "DB-G31i-Users"
        tabDatabase(0, 18) = "DB-H13d-users"
        tabDatabase(0, 19) = "DB-H13e-users"
        tabDatabase(0, 20) = "DB-H13f-users"
        tabDatabase(0, 21) = "DB-H13g-users"
        tabDatabase(0, 22) = "DB-H13h-users"
        tabDatabase(0, 23) = "DB-H13i-users"
        tabDatabase(0, 24) = "DB-I23d-users"
        tabDatabase(0, 25) = "DB-I23e-users"
        tabDatabase(0, 26) = "DB-I23f-users"
        tabDatabase(0, 27) = "DB-I23g-users"
        tabDatabase(0, 28) = "DB-I23h-users"
        tabDatabase(0, 29) = "DB-I23i-users"
        tabDatabase(0, 30) = "DB-J32d-Users"
        tabDatabase(0, 31) = "DB-J32e-Users"
        tabDatabase(0, 32) = "DB-J32f-Users"
        tabDatabase(0, 33) = "DB-J32g-Users"
        tabDatabase(0, 34) = "DB-J32h-Users"
        tabDatabase(0, 35) = "DB-J32i-Users"

        For i = 0 To UBound(tabDatabase, 2)
            dirSearcher.Filter = "(&(homeMDB=CN=" & tabDatabase(0, i) & ",CN=Databases,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=First Organization,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=igbmc,DC=u-strasbg,DC=fr))"
            Dim results As SearchResultCollection = dirSearcher.FindAll
            Dim j As Integer = 0
            For Each utilisateur In results
                j += 1
            Next
            tabDatabase(1, i) = j
            results.Dispose()
            results = Nothing
        Next i
        dirSearcher.Dispose()
        dirSearcher = Nothing
        Ldap.Close()
        Ldap.Dispose()
        Ldap = Nothing


    End Sub

    Function EmployeeIDExist(ByVal EmployeeID As String) As Boolean
        Dim resultat As Boolean = False
        Dim objAD As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
        Dim searcher As DirectorySearcher = New DirectorySearcher(objAD)
        searcher.PageSize = 2000
        searcher.Filter = "(EmployeeID=" & EmployeeID & ")"

        Dim result As SearchResult = searcher.FindOne
        If Not result Is Nothing Then
            resultat = True
        End If
        searcher.Dispose()
        objAD.Close()
        objAD.Dispose()
        Return resultat
    End Function

    Function DetermineLogin(ByVal prenom As String, ByVal nom As String, ByVal userID As String)
        Dim login As String = ""

        Dim tablogin As String(,) = Commun.CreateTabHistoAliasLogin("login")
        Dim index As Integer = Commun.MultiIndexOf(tablogin, userID, 0)

        'si le numero d'emloyé existe deja dans l'historique, on réatribue le meme login
        If index <> -1 And userID <> "0" Then
            login = tablogin(1, index)

        Else
            'sinon, on le construit
            prenom = Replace(prenom, " ", "")
            nom = Replace(nom, " ", "")
            prenom = Replace(prenom, "'", "")
            nom = Replace(nom, "'", "")
            prenom = Replace(prenom, "-", "")
            nom = Replace(nom, "-", "")
            prenom = LCase(prenom)
            nom = LCase(nom)

            Dim n As Integer = 7
            Dim p As Integer = 1
            Dim ctrlLoginLibre = False

            Do
                login = Strings.Left(nom, n) & Strings.Left(prenom, p)
                If Commun.ctrlAliasDispo(login) = True Then
                    GoTo sortie
                End If
                n -= 1
                p += 1
            Loop Until n = 1

            Dim i As Integer = 1
            Do
                login = Strings.Left(nom, 6) & Strings.Left(prenom, 1) & i
                If Commun.ctrlAliasDispo(login) = True Then
                    GoTo sortie
                End If
                i += 1
            Loop
        End If

sortie:
        Return login

    End Function

    Function UserExist(ByVal CNAVerifier As String) As String

        Dim monEntry As New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursActifs]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
        Dim maRecherche As DirectorySearcher = New DirectorySearcher
        maRecherche.PageSize = 2000
        Dim resultat As String = ""

        Try

            maRecherche.Filter = "(&(objectClass=user) (CN=" + CNAVerifier + "))"
            maRecherche.SearchScope = SearchScope.Subtree

            Dim result As SearchResult = maRecherche.FindOne()
            monEntry.Close()
            monEntry.Dispose()
            monEntry = Nothing

            If Not result Is Nothing Then
                Return result.Path
            End If
            maRecherche.Dispose()
            maRecherche = Nothing
        Catch ex As Exception

        End Try

        Return resultat

    End Function

    Shared Function TrouverPetiteDatabase() As String
        Dim NBpetit As Integer = tabDatabase(1, 0)
        Dim DBpetit As String = tabDatabase(0, 0)
        Dim ipetit As Integer = 0
        For i = 1 To UBound(tabDatabase, 2)
            If tabDatabase(1, i) < NBpetit Then
                NBpetit = tabDatabase(1, i)
                DBpetit = tabDatabase(0, i)
                ipetit = i
            End If
        Next
        tabDatabase(1, ipetit) = NBpetit + 1
        Return DBpetit
    End Function

    Shared Sub commandePWSMailbox(ByVal login As String, ByVal db As String)
        Dim ctrlDomain As String = "serv-ad1"
        Try
            Dim pCredential As PSCredential
            Dim pConnectionInfo As WSManConnectionInfo
            Dim pRunspace As Runspace
            Dim pShell As PowerShell
            Dim pCommand As PSCommand
            Dim pCommand1 As PSCommand
            Dim pResult As Collection(Of PSObject)
            Dim pResult1 As Collection(Of PSObject)
            Dim pCommand2 As PSCommand
            Dim pResult2 As Collection(Of PSObject)
            Dim pCommand3 As PSCommand
            Dim pResult3 As Collection(Of PSObject)

            '-- set credentials      
            pCredential = DirectCast(Nothing, PSCredential) 'New PSCredential("igbmc\steph", CreateSecurePasswordString("aaaaaa"))

            '-- set connection info
            pConnectionInfo = New WSManConnectionInfo(New Uri("http://" & RecupDataini.RecupVar("[CasExchangeServer]") & "/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", pCredential)

            '-- create remote runspace
            pRunspace = RunspaceFactory.CreateRunspace(pConnectionInfo)
            pRunspace.Open()

            '-- create powershell
            pShell = PowerShell.Create
            pShell.Runspace = pRunspace

            '-- create command
            pCommand = New PSCommand
            With pCommand
                .AddCommand("Enable-mailbox")
                .AddParameter("identity", login)
                .AddParameter("alias", login)
                .AddParameter("Database", db)
                .AddParameter("DomainController", ctrlDomain)
            End With

            '-- add command to powershell
            pShell.Commands = pCommand

            '-- invoke the powershell
            pResult = pShell.Invoke

            pCommand1 = New PSCommand

            With pCommand1
                .AddCommand("Set-CASMailbox")
                .AddParameter("identity", login)
                .AddParameter("ActiveSyncEnabled", True)
                .AddParameter("ImapEnabled", False)
                .AddParameter("PopEnabled", False)
                .AddParameter("DomainController", ctrlDomain)
            End With

            pShell.Commands = pCommand1
            pResult1 = pShell.Invoke

            pCommand2 = New PSCommand

            With pCommand2
                .AddCommand("Set-MailboxCalendarConfiguration")
                .AddParameter("identity", login)
                .AddParameter("FirstWeekOfYear", "FirstFourDayWeek")
                '.AddParameter("WeatherUnit", "Celsius")
                .AddParameter("DomainController", ctrlDomain)
            End With

            pShell.Commands = pCommand2
            pResult2 = pShell.Invoke

            pCommand3 = New PSCommand

            With pCommand3
                .AddCommand("Set-MailboxRegionalConfiguration")
                .AddParameter("identity", login)
                .AddParameter("TimeZone", "Romance Standard Time")
                .AddParameter("Language", "fr-FR")
                .AddParameter("LocalizeDefaultFolderName", True)
                .AddParameter("DomainController", ctrlDomain)
            End With

            pShell.Commands = pCommand3
            pResult3 = pShell.Invoke

            pRunspace.Close()
            pRunspace.Dispose()
            pRunspace = Nothing

            Commun.Journal("Creation de la boite mail Réussie: " & login)
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Creation du compte mail: " & e.Message & " : " & login, True)
        End Try

    End Sub
End Class
