Imports System.DirectoryServices
Imports System.IO
Imports System.Net.Mail
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Collections.ObjectModel
Imports System.Security
Imports Tamir.SharpSsh

Public Class Form1

    'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    'Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Integer
    'Const ADS_UF_PASSWD_CANT_CHANGE = &H40
    'Dim Fexiste As Boolean
    'Dim cmdFtp As String
    'Dim cmdAccAD As String
    'Dim cmdDelAD As String
    'Dim cmdModAD As String
    'Dim fichierAdd As String
    'Dim fichierDel As String
    'Dim fichierMod As String
    'Dim objDom As ActiveDs.IADsContainer

    Dim objGroupe As Object
    Dim tabDatabase(1, 8) As String
    Dim tabExcepUser As String()
    
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'NbUserDatabaseExchange()
        'Dim petiteDB As String = TrouverPetiteDatabase()
        Commun.fichierLog = "c:\temp\LogAutoCompte.log"
        Commun.Journal("Debut du Traitement")

        CtrlGroupAdmins()



        ' RECUPERATION FICHIERS PAR FTP
        Dim persoFTP As FTPP = New FTPP()

        persoFTP.m_FTPSite = "ftp://" & Commun.RecupVar("[FCCserver]")
        persoFTP.UserName = Commun.RecupVar("[FCCAccount]")
        persoFTP.Password = Commun.RecupVar("[FCCPassword]")

        Dim listeFichierC As List(Of String) = persoFTP.GetFileList("/" & Commun.RecupVar("[FCCFilePath]"), "c", ".txt")
        Dim listeFichierS As List(Of String) = persoFTP.GetFileList("/" & Commun.RecupVar("[FCCFilePath]"), "s", ".txt")

        Try
            For Each fileServ In listeFichierC
                persoFTP.GetFile("/" & Commun.RecupVar("[FCCFilePath]") & "/" & fileServ, "c:\temp\" & fileServ)
                persoFTP.UploadFile("c:\temp\" & fileServ, "/cmpttmp/" & fileServ)
            Next

            For Each fileServ In listeFichierS

                persoFTP.GetFile("/" & Commun.RecupVar("[FCCFilePath]") & "/" & fileServ, "c:\temp\" & fileServ)
                persoFTP.UploadFile("c:\temp\" & fileServ, "/cmpttmp/" & fileServ)
            Next

        Catch ex As Exception
            Commun.Journal("ECHEC : FTP : Recuperation du fichier : " & ex.Message, True)
        End Try

        Call fichierRep("create") 'traite les fichiers de création de compte
        Call fichierRep("suppr")

        'Mise a jour des utilisateurs provisoires

        Dim OUProvisoire As DirectoryEntry = New DirectoryEntry("LDAP://OU=Users Provisoires,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim searcher As DirectorySearcher = New DirectorySearcher(OUProvisoire)
        searcher.Filter = "(&(objectClass=user))"

        Dim result As SearchResultCollection = searcher.FindAll()
        Dim usr1 As ActiveDs.IADsContainer
        Dim dateExpD As Date
        Dim dateExpT As String

        For Each utilisateurProv As SearchResult In result
            usr1 = GetObject(utilisateurProv.Path)
            If CStr(Strings.Left(usr1.AccountExpirationDate, 10)) <> Strings.Right(usr1.Description, 10) Then
                usr1.Put("description", "COMPTE PROVISOIRE expire le: " & Strings.Left(CStr(usr1.AccountExpirationDate), 10))
                usr1.SetInfo()
            End If
            If usr1.AccountExpirationDate < Now Then
                usr1.Put("userAccountControl", &H2) 'Desactive le compte
                usr1.SetInfo()
            End If
            dateExpT = Strings.Left(usr1.Get("description"), 17)
            dateExpD = usr1.AccountExpirationDate 'CDate(VB6.Format(Strings.Right(dateExpT, 10), "dd/mm/yyyy"))
            If dateExpD < DateAdd(Microsoft.VisualBasic.DateInterval.Month, -3, Now) And usr1.AccountDisabled = True And dateExpT = "COMPTE PROVISOIRE" Then
                Dim userASuppr As DirectoryEntry = utilisateurProv.GetDirectoryEntry
                OUProvisoire.Children.Remove(userASuppr)
                DelUnixUser(userASuppr.Properties("SAMAccountName").Value)
                userASuppr.Close()
                userASuppr = Nothing
            End If
            usr1 = Nothing
        Next
        OUProvisoire.Close()
        OUProvisoire = Nothing


        'Désactivation des comptes de l'OU "Users Inconnus"
        Dim objUser As ActiveDs.IADsUser
        Dim OUUsersInconnus As ActiveDs.IADsContainer = GetObject("LDAP://OU=Users Inconnus,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
        For Each objUser In OUUsersInconnus
            objUser.Put("userAccountControl", &H2) 'Desactive le compte
            objUser.SetInfo()
        Next
        OUUsersInconnus = Nothing
        objUser = Nothing

        'Suppression des comptes dans l'OU Exception qui n'ont rien a y faire
        Dim OUException As New DirectoryEntry("LDAP://OU=Exceptions,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim searchOUException As DirectorySearcher = New DirectorySearcher(OUException)
        searchOUException.Filter = "(&(objectClass=user))"
        searchOUException.PropertiesToLoad.Add("SamAccountName")
        Dim searchOUExceptionResult As SearchResultCollection = searchOUException.FindAll()
        For Each userOUException As SearchResult In searchOUExceptionResult
            If PresenceExceptionUser(userOUException.Properties("SamAccountName")(0)) = False Then
                userOUException.GetDirectoryEntry.DeleteTree()
            End If
        Next

        'Ajout au groupe G_Domain_DisableOpenSession des comptes de l'OU "Comptes Désactivés"

        Dim OUComptesDesactives As New DirectoryEntry("LDAP://OU=Comptes Désactivés,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim GroupDisabled As New DirectoryEntry("LDAP://CN=G_Domain_DisableOpenSession,OU=REFUS,OU=Groupes,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim GroupDomainUsers As New DirectoryEntry("LDAP://CN=Utilisa. du domaine,CN=Users,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim searchDesactiv As DirectorySearcher = New DirectorySearcher(OUComptesDesactives)
        searchDesactiv.Filter = "(&(objectClass=user))"
        Dim searchDesactivResult As SearchResultCollection = searchDesactiv.FindAll()
        For Each objUserDesactiv As SearchResult In searchDesactivResult

            Dim RID As Integer = Commun.PrimaryGroupId(GroupDisabled.Properties("objectSid").Value)
            Dim DNDesactivUser As String = Replace(objUserDesactiv.Path, "LDAP://", "")

            Dim DirEntry As New DirectoryEntry("LDAP://" & DNDesactivUser)
            If DirEntry.Properties("primaryGroupID").Value <> RID Then
                Try
                    'Ajout au groupe G_Domaine_DisableOpenSession
                    GroupDisabled.Invoke("Add", New Object() {objUserDesactiv.Path})
                Catch
                End Try

                Try
                    'definition du groupe G_Domain_DisableOpenSession comme groupe principal
                    DirEntry.Properties("primaryGroupID").Value = RID
                    DirEntry.CommitChanges()
                Catch
                End Try

                Try
                    'Retrait du groupe Utilisa. du domaine
                    GroupDomainUsers.Invoke("Remove", New Object() {objUserDesactiv.Path})
                Catch
                End Try


            End If
        Next

        GroupDomainUsers.Close()
        GroupDomainUsers = Nothing
        GroupDisabled.Close()
        GroupDisabled = Nothing
        OUComptesDesactives.Close()
        OUComptesDesactives = Nothing


        For Each fileServ In listeFichierC
            persoFTP.DeleteFile("/" & Commun.RecupVar("[FCCFilePath]") & "/" & fileServ)
        Next

        For Each fileServ In listeFichierS
            persoFTP.DeleteFile("/" & Commun.RecupVar("[FCCFilePath]") & "/" & fileServ)
        Next

        Try
            Shell(My.Application.Info.DirectoryPath & "\MAJZoneInfo.exe")
        Catch
            Commun.Journal("ECHEC:Lancement de MAJZoneInfo", True)
        End Try

        Commun.Journal("Fin du Traitement" & vbCrLf)

        If Commun.controlSendMail = True Then
            Commun.SendEmail(Commun.RecupVar("[AdminScriptLogin]") & "@igbmc.fr", Commun.RecupVar("[MailErrors]"), "AutoCompte.NET : Rapport d'erreur", Commun.journalECHECMail)
        End If
        Application.Exit()
    End Sub
    Private Sub fichierRep(ByVal type As String)

        Dim Compte As Integer
        Dim fich As Object
        Dim rep As String = ""
        Dim mask As String = ""

        If type = "create" Then
            rep = "c:\temp"
            mask = "c*.txt"
            type = "create"
        End If

        If type = "suppr" Then
            rep = "c:\temp"
            mask = "s*.txt"
            type = "suppr"
        End If

        Dim repertoire = Directory.GetFiles(rep, mask)
        Compte = -1
        For Each fich In repertoire
            Compte = Compte + 1
            If type = "create" Then
                Call createCompte(fich)
            End If

            If type = "suppr" Then
                Call SupprimCompte(fich)
            End If
        Next
    End Sub
    Private Sub createCompte(ByVal fichier As String)

        NbUserDatabaseExchange()


        ' TRAITEMENT DU FICHIER DE CREATION DE COMPTES
        'On Error Resume Next

        Dim Tc() As String
        Dim Champc() As String

        Dim objUser As DirectoryEntry
        Dim objOUUtilisateurs As DirectoryEntry
        If FileLen(fichier) > 0 Then

            Dim EqDescr As String = " "
            FileOpen(4, fichier, OpenMode.Input)
            Dim i As Integer = -1
            While Not EOF(4)
                i = i + 1
                ReDim Preserve Tc(i)
                Tc(i) = LineInput(4)

                Tc(i) = Replace(Tc(i), """", "")
                Champc = Split(Tc(i), ",")
                If InStr(Tc(i), ",") <> 0 Then 'evite les lignes blanches
                    Commun.Journal("Debut de Creation de compte : " & Champc(2))

                    If DirectoryEntry.Exists("LDAP://" & Commun.TransformeSAMACCOUNTenCN(Champc(4) & " grp")) Then
                        Dim destAD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.TransformeSAMACCOUNTenCN(Champc(4) & " grp"))
                        EqDescr = destAD.Properties("Description").Value
                        destAD.Close()
                        destAD = Nothing
                    End If

                    Try
                        'verification si un compte provisoire exsite
                        Dim tempCtrlPathExist As String = ""
                        '           UserExist controle l'existance d'un utilisateur uniquement dans l'OU Utilisateurs
                        tempCtrlPathExist = UserExist(Champc(1) & " " & Champc(0))

                        objOUUtilisateurs = New DirectoryEntry("LDAP://OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")


                        If tempCtrlPathExist = "" Then

                            objUser = objOUUtilisateurs.Children.Add("CN=" & Champc(1) & " " & Champc(0), "user")
                            objUser.Properties("SAMAccountName").Value = Champc(2)
                            objUser.Properties("userPrincipalName").Value = Champc(2) & "@igbmc.fr"
                            objUser.CommitChanges()

                            objUser.Properties("SN").Value = Champc(0)
                            objUser.Properties("givenName").Value = Champc(1)
                            objUser.Properties("Displayname").Value = Champc(1) & " " & Champc(0)
                            objUser.Properties("Comment").Value += "Créé le: " & Strings.Left(CStr(Now), 10) & vbCrLf
                            objUser.CommitChanges()
                        Else
                            objUser = New DirectoryEntry(tempCtrlPathExist)
                            If objUser.Path <> "LDAP://CN=" & Champc(1) & " " & Champc(0) & ",OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr" Then
                                objUser.MoveTo(objOUUtilisateurs)
                            End If
                            Commun.Journal("ECHEC : Creation de compte : le compte existait déja, il a été modifié : " & Champc(2), True)
                        End If

                        objOUUtilisateurs.Close()
                        objOUUtilisateurs = Nothing

                    Catch
                        Commun.Journal("ECHEC : Creation de compte : le compte existait déja, une erreur s'est produite : " & Champc(2), True)
                    End Try

                    Try
                        objUser.Invoke("SetPassword", New Object() {Champc(3)})
                        objUser.CommitChanges()

                        objUser.Properties("ObjectClass").Add("posixaccount")

                        objUser.Properties("mail").Value = Champc(2) & "@igbmc.fr"
                        'objUser.Properties("physicalDeliveryOfficeName").Value = "----"
                        'objUser.Properties("TelephoneNumber").Value = "----"
                        objUser.Properties("Department").Value = EqDescr
                        objUser.CommitChanges()


                        objUser.Properties("EmployeeID").Value = Champc(9)
                        objUser.Properties("company").Value = "IGBMC"
                        objUser.Properties("DepartmentNumber").Value = Champc(4)
                        objUser.CommitChanges()
                    Catch
                        Commun.Journal("ECHEC : Creation de compte : Attributs : " & Champc(2), True)
                    End Try

                    'Ajout de la valeur posixAccount à la class objectClass du compte
                    Try
                        objUser.Properties("ObjectClass").Add("posixaccount")
                        objUser.CommitChanges()
                    Catch
                        Commun.Journal("ECHEC : Ajout PosixAccount au compte : Attributs : " & Champc(2), True)
                    End Try


                    'definition du GID du groupe de l'utilisateur
                    'Si le groupe de l'equipe n'existe pas dans l'AD, creation du groupe
                    If Not DirectoryEntry.Exists("LDAP://CN=" & Champc(5) & "_eq,OU=Equipes,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr") Then
                        Dim GID = Commun.GIDMini
                        Commun.getFileFTP("/" & Commun.RecupVar("[FCCFilePath]") & "/equipeinfo.txt", "c:\temp\equipeinfo.txt")
                        Dim monStreamReader As StreamReader = New StreamReader("c:\temp\equipeinfo.txt")
                        Dim ligne As String
                        Dim reponse As String = ""

                        Do
                            ligne = monStreamReader.ReadLine()
                            If Strings.Left(ligne, Len(Champc(5))) = Champc(5) Then
                                Commun.CreationGroupesEquipeInfo(Split(ligne, ",")(0), Split(ligne, ",")(1), Split(ligne, ",")(4))
                            End If

                        Loop Until ligne Is Nothing
                        monStreamReader.Close()
                    End If

                    'Recherche du gid du groupe de l'equipe de l'utilisateur.
                    Dim groupeAD As DirectoryEntry = New DirectoryEntry("LDAP://CN=" & Champc(5) & "_eq,OU=Equipes,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr")
                    Dim GIDNumber As String = groupeAD.Properties("gidNumber").Value
                    groupeAD.Close()
                    groupeAD = Nothing

                    'Creation du compte LDAP Unix
                    Try
                        Dim donneesUnix As String = CreatUnixUser(Champc(0), Champc(1), Champc(2), Champc(3), Champc(5), Champc(7) & "\" & Champc(2), GIDNumber)
                        '4 valeurs unix a integrer dans l'AD : gidNumber,uidNumber,loginShell,homeDirectory
                        Dim tabDonneesUnix As String() = Split(donneesUnix, ",")
                        objUser.Properties("uidNumber").Value = tabDonneesUnix(0)
                        objUser.Properties("loginShell").Value = tabDonneesUnix(1)
                        objUser.Properties("unixHomeDirectory").Value = tabDonneesUnix(2)

                        '1 valeur qui correspond au login
                        objUser.Properties("uid").Value = Champc(2)
                        objUser.CommitChanges()
                    Catch
                        Commun.Journal("ECHEC : Creation de l'utilisateur Unix : " & Champc(2), True)
                    End Try

                    'AJOUT DU GROUPE EQUIPE COMPTABLE
                    If EqDescr <> " " Then
                        Try
                            objGroupe = GetObject("LDAP://CN=" & Champc(4) & " grp,OU=Equipes,DC=igbmc,DC=u-strasbg,DC=fr")
                            objGroupe.Add(objUser.Path)
                            objGroupe = Nothing
                        Catch
                            Commun.Journal("ECHEC : l'utilisateur était deja membre de l'equipe comptable : " & Champc(2), True)
                        End Try
                    End If

                    'ajout au groupe administratif en fonction de l'equipe info
                    Try

                        Dim objGroupeADM = GetObject("LDAP://CN=G_ADMINISTRATIF_Users,OU=Droits,OU=Groupes,DC=igbmc,DC=u-strasbg,DC=fr")
                        Dim membreEquipeADM As Boolean = Commun.MembreEquipeAdministratif(objUser)

                        If membreEquipeADM = True Then
                            objGroupeADM.Add(objUser.Path)
                        End If

                        objGroupeADM = Nothing

                    Catch e As Exception
                        Commun.Journal("ECHEC : ajout à l'equipe ADMINISTRATIVE de l'utilisateur : " & e.Message & " : " & Champc(2), True)
                    End Try

                    'Determiner le Flag sur "l'utilisateur ne peut pas changer de mot de passe" et "le mot de passe n'expire jamais"
                    Try
                        Dim objUserNT = GetObject("WinNT://igbmc/" & Champc(2))
                        objUserNT.Put("userFlags", &H10040)
                        objUserNT.setinfo()
                        objUserNT = Nothing
                    Catch
                    End Try

                    Try
                        'Creation des alias de mails
                        Dim NomCN As String = objUser.Properties("distinguishedName").Value
                        Dim Ldap As DirectoryEntry = New DirectoryEntry("LDAP://" & NomCN)
                        Dim adress As New List(Of String)
                        Dim aliasSMTP = Replace(Champc(1), " ", "-") & "." & Replace(Champc(0), " ", "-")
                        aliasSMTP = Replace(aliasSMTP, "'", "")

                        adress.Add("smtp:" & aliasSMTP & "@igbmc.fr")
                        adress.Add("smtp:" & aliasSMTP & "@igbmc.u-strasbg.fr")
                        Ldap.Properties("proxyAddresses").Value = adress.ToArray()
                        Ldap.CommitChanges()
                        Ldap.Close()
                    Catch e As Exception
                        Commun.Journal("ECHEC : Creation de compte : Creation des alias mail : " & e.Message & " : " & Champc(2), True)
                    End Try

                    'creation de la boite mail Exchange
                    Try
                        Dim petiteDB As String = TrouverPetiteDatabase()
                        commandePWSMailbox(Champc(2), petiteDB)
                    Catch e As Exception
                        Commun.Journal("ECHEC : Creation de compte : Creation du compte mail : " & e.Message & " : " & Champc(2), True)
                    End Try


                    'Ajout a la liste de diffusion "annouce"
                    Try
                        objGroupe = GetObject("LDAP://CN=Announce,OU=Listes de diffusion,OU=Exchange,DC=igbmc,DC=u-strasbg,DC=fr")
                        objGroupe.Add(objUser.Path)
                        objGroupe = Nothing
                    Catch e As Exception
                        Commun.Journal("ECHEC : Creation de compte : Ajout liste de diffusions ANNOUCE : " & e.Message & " : " & Champc(2), True)
                    End Try

                    'Ajout a la liste de diffusion "ce"
                    Try
                        objGroupe = GetObject("LDAP://CN=ce,OU=Listes de diffusion,OU=Exchange,DC=igbmc,DC=u-strasbg,DC=fr")
                        objGroupe.Add(objUser.Path)
                        objGroupe = Nothing
                    Catch e As Exception
                        Commun.Journal("ECHEC : Creation de compte : Ajout liste de diffusions CE : " & e.Message & " : " & Champc(2), True)
                    End Try


                    'Ajout a la liste de diffusion "phd" ou "postdoc"
                    Try
                        If Champc(8) = "phd" Or Champc(8) = "postdoc" Then
                            objGroupe = GetObject("LDAP://CN=" & Champc(8) & ",OU=Listes de diffusion,OU=Exchange,DC=igbmc,DC=u-strasbg,DC=fr")
                            objGroupe.Add(objUser.Path)
                            objGroupe = Nothing
                        End If
                    Catch e As Exception
                        Commun.Journal("ECHEC : Creation de compte : Ajout liste de diffusions PHD ou POSTDOC : " & e.Message & " : " & Champc(2), True)
                    End Try

                    objUser.Close()
                    objUser = Nothing
                    Dim textMail = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN""><HTML><HEAD><META HTTP-EQUIV=""CONTENT-TYPE"" CONTENT=""text/html; charset=windows-1252""><TITLE></TITLE><STYLE TYPE=""text/css""><!--P { margin-bottom: 0.21cm }--></STYLE></HEAD><BODY LANG=""fr-FR"" DIR=""LTR"">Bienvenue a l'IGBMC " & Champc(1) & " " & Champc(0) & " (compte ouvert le " & DateTime.Today.ToString("dd/MM/yyyy") & ")</FONT></FONT></P><P ALIGN=CENTER STYLE=""margin-bottom: 0cm""><BR></P><TABLE WIDTH=642 BORDER=1 BORDERCOLOR=""#000000"" CELLPADDING=0 CELLSPACING=0><COL WIDTH=150><COL WIDTH=490><TR><TD COLSPAN=2 WIDTH=640 VALIGN=TOP><center>nom d'utilisateur: <B>" & Champc(2) & "</B><BR>mot de passe (8 caract&egrave;res): <B>" & Champc(3) & "</B><BR>(&Agrave; changer sur http://informatique.igbmc.u-strasbg.fr)<BR><BR>Votre adresse e-mail est :<BR><B>" & Champc(1) & "." & Champc(0) & "@igbmc.fr</B><BR>ou bien " & Champc(2) & "@igbmc.fr</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Messagerie et agenda</center></TD><TD WIDTH=490><center>Vous pouvez acc&eacute;der &agrave; vos mails et &agrave; votre agenda personnel via l'adresse https://igbmcmail.igbmc.fr/owa <BR></center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Commandes de cellules/s&eacute;quen&ccedil;age</center></TD><TD WIDTH=490><center>Il faut demander un compte &agrave; Michel Offner (tel: 3277) ou Tony Moutaux (tel: 3395)</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Espace disque</center></TD><TD WIDTH=490><center>Vous disposez d'un espace de stockage sauvegard&eacute; auquel vous pouvez acc&eacute;der via l'adresse " & Champc(7) & " si vous utiliser un PC sous Windows.</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Qui contacter en cas de probl&egrave;me?</center></TD><TD WIDTH=490><center>Envoyez un mail &agrave; <B>helpdesk@igbmc.fr</B></center></TD></TR></TABLE><BR><TABLE WIDTH=642 BORDER=1 BORDERCOLOR=""#000000"" CELLPADDING=0 CELLSPACING=0><COL WIDTH=150><COL WIDTH=490><TR><TD COLSPAN=2 WIDTH=640 VALIGN=TOP><center>Your login: <B>" & Champc(2) & "</B><BR>Your password: <B>" & Champc(3) & "</B><BR>(Please change it on http://informatique.igbmc.u-strasbg.fr)<BR><BR>Your email address is:<BR><B>" & Champc(1) & "." & Champc(0) & "@igbmc.fr</B><BR>or " & Champc(2) & "@igbmc.fr</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Email and Calendar</center></TD><TD WIDTH=490><center>You can access to your mail and to the shared calendar at https://igbmcmail.igbmc.fr/owa <BR></center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Orders of cells, sequencing</center></TD><TD WIDTH=490><center>If you need to order cells, you must ask for an account to Michel Offner (3277) or Tony Moutaux (3395).</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Disk space</center></TD><TD WIDTH=490><center>You can access to your shared disk space at:<BR>Windows: " & Champc(7) & "</center></TD></TR><TR VALIGN=TOP><TD WIDTH=150><center>Contact</center></TD><TD WIDTH=490><center>Send an email to: <B>helpdesk@igbmc.fr</B></center></TD></TR></TABLE></BODY></HTML>"


                    Commun.SendEmail(Commun.RecupVar("[AdminScriptLogin]") & "@igbmc.fr", "serviceinfo@igbmc.fr", "[Creation de compte] " & Champc(1) & " " & Champc(0), textMail)
                    Commun.Journal("Fin de Creation de compte : " & Champc(2))
                End If
            End While
            FileClose(4)

        End If
        Kill(fichier)

    End Sub

    Public Sub SupprimCompte(ByVal fichier As String)
        ' TRAITEMENT DU FICHIER DE SUPPRESSION DE COMPTES
        Dim Ts() As String
        Dim k As Integer

        If FileLen(fichier) > 0 Then
            'objDom = GetObject("WinNT://" & DomNT)

            FileOpen(3, fichier, OpenMode.Input)
            k = -1

            Dim OUDisable As DirectoryEntry = New DirectoryEntry("LDAP://OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
            Dim dirSearcher As DirectorySearcher = New DirectorySearcher(OUDisable)
            Dim ouAdmins As DirectoryEntry = New DirectoryEntry("LDAP://OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr")
            Dim searcherouAdmins As DirectorySearcher = New DirectorySearcher(ouAdmins)

            While Not EOF(3)
                k = k + 1
                ReDim Preserve Ts(k)
                Ts(k) = LineInput(3)
                If Len(Ts(k)) > 2 Then 'evite les lignes blanches

                    Try

                        dirSearcher.Filter = "(&(objectClass=user) (SAMAccountName=" & Ts(k) & "))"
                        dirSearcher.SearchScope = SearchScope.Subtree
                        Dim result As SearchResult = dirSearcher.FindOne()

                        If Not result Is Nothing Then
                            result.GetDirectoryEntry.DeleteTree()
                            OUDisable.CommitChanges()
                            DelUnixUser(Ts(k))
                        End If

                    Catch ex As Exception
                        Commun.Journal("ECHEC : Suppression de compte : " & Ts(k) & " : " & ex.Message, True)
                    End Try



                    Try
                        'Suppression du compte adm s'il existe

                        searcherouAdmins.Filter = "(&(objectClass=user) (SAMAccountName=" & Ts(k) & "Adm))"
                        Dim result As SearchResult = searcherouAdmins.FindOne()
                        If Not result Is Nothing Then
                            result.GetDirectoryEntry.DeleteTree()
                            ouAdmins.CommitChanges()
                        End If

                    Catch ex As Exception
                        Commun.Journal("ECHEC : Suppression du compte ADM : " & Ts(k) & "Adm : " & ex.Message, True)
                    End Try

                    'Suppression du dossier de profil itinerant
                    Try
                        If Directory.Exists("\\labo1\home1\profils\" & Ts(k)) Then
                            'fermer les fichiers ouverts par l'utilisateur
                            Dim ProcessProperties As New ProcessStartInfo
                            ProcessProperties.FileName = "openfiles"
                            ProcessProperties.Arguments = "/disconnect /s home1 /a " & Ts(k)
                            ProcessProperties.WindowStyle = ProcessWindowStyle.Hidden
                            Dim myProcess As Process = Process.Start(ProcessProperties)

                            Directory.Delete("\\labo1\home1\profils\" & Ts(k), True)
                        End If
                    Catch ex As Exception
                        Commun.Journal("ECHEC : Suppression du dossier Profil Itinerant : " & Ts(k) & " : " & ex.Message, True)
                    End Try

                    'Suppression du dossier de profil itinerant du compte adm
                    Try
                        If Directory.Exists("\\labo1\home1\profils\" & Ts(k) & "Adm") Then
                            Directory.Delete("\\labo1\home1\profils\" & Ts(k) & "Adm", True)
                        End If
                    Catch ex As Exception
                        Commun.Journal("ECHEC : Suppression du dossier Profil Itinerant Adm : " & Ts(k) & "Adm : " & ex.Message, True)
                    End Try
                End If
            End While
            dirSearcher = Nothing
            OUDisable.Close()
            OUDisable = Nothing
            searcherouAdmins = Nothing
            ouAdmins.Close()
            ouAdmins = Nothing
            FileClose(3)
        End If

        Kill(fichier)

    End Sub
    Function SearchPathAD(ByVal filtreAttribut As String, ByVal filtreValue As String, ByVal sortieAttribut As String) As String
        Dim AD As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")
        Dim SearchAD As DirectorySearcher = New DirectorySearcher(AD)
        SearchAD.PropertiesToLoad.Add(sortieAttribut)
        SearchAD.Filter = "(&(" & filtreAttribut & "=" & filtreValue & "))"
        Dim result As SearchResult = SearchAD.FindOne()
        AD.Close()
        AD = Nothing
        If result Is Nothing Then
            Return ""
        Else
            Return result.Properties(sortieAttribut)(0)
        End If

    End Function
    Public Sub CtrlGroupAdmins()

        Dim maListeMembres As New ArrayList
        Dim samaccountGroup As String = ""
        Dim listMembresAutorises As String = ""
        For compteur = 0 To 4
            Select Case compteur
                Case 0
                    samaccountGroup = "Admins du domaine"
                    listMembresAutorises = Commun.RecupVar("[Group_Admins_du_domaine]")
                    'listMembresAutorises = "Administrateur,steph,stephadm,tina,guiseithadm,SERV-TMG$,SERV-CLUSTER2$,SERV-EXCHANGE$,userprog,krbtgt"
                Case 1
                    samaccountGroup = "Administrateurs de l'entreprise"
                    listMembresAutorises = Commun.RecupVar("[Group_Administrateurs_de_l_entreprise]")
                    'listMembresAutorises = "Administrateur,steph,stephadm,guiseithadm,userprog"
                Case 2
                    samaccountGroup = "Administrateurs"
                    listMembresAutorises = Commun.RecupVar("[Group_Administrateurs]")
                    'listMembresAutorises = "Administrateur,steph,stephadm,Administrateurs,Admins du domaine,scripsteph"
                Case 3
                    samaccountGroup = "Administrateurs du schéma"
                    listMembresAutorises = Commun.RecupVar("[Group_Administrateurs_du_schema]")
                    'listMembresAutorises = "Administrateur,steph,stephadm"
                Case 4
                    samaccountGroup = "Administrateurs DHCP"
                    listMembresAutorises = Commun.RecupVar("[Group_Administrateurs_DHCP]")
                    'listMembresAutorises = "Administrateur,steph,stephadm,guiseithadm,userprog"
            End Select

            If InStr(listMembresAutorises, "Administrateur,steph,stephadm") <> 0 Then

                Dim tabMembresAutorises = Split(listMembresAutorises, ",")
                Dim monGroupe As DirectoryEntry = New DirectoryEntry("LDAP://" & SearchPathAD("SAMAccountName", samaccountGroup, "distinguishedName"))
                Try
                    ' Groupe dont les membres sont à lister
                    Dim tabMembre As String() = Commun.MembresDuGroupe(samaccountGroup, True)
                    For Each unMembre In tabMembre

                        Dim userLogin As String = SearchPathAD("distinguishedName", unMembre.ToString, "SAMAccountName") 'user.Properties("SAMAccountName").Value

                        If Array.IndexOf(tabMembresAutorises, userLogin) = -1 Then
                            monGroupe.Invoke("Remove", New Object() {"LDAP://" & unMembre.ToString})
                        End If
                    Next unMembre

                    For i = 0 To UBound(tabMembresAutorises)
                        Dim DNUser As String = SearchPathAD("SAMAccountName", tabMembresAutorises(i), "distinguishedName")
                        Dim found As Integer = Array.IndexOf(tabMembre, DNUser)
                        If found = -1 Then
                            monGroupe.Invoke("Add", New Object() {"LDAP://" & DNUser})
                        End If
                    Next i
                    monGroupe.Close()
                Catch ex As Exception
                    Commun.Journal("ECHEC : Controle des groupes Admins : case " & compteur & " : " & ex.Message, True)
                End Try
                monGroupe.Close()
                monGroupe = Nothing
            Else
                Commun.Journal("ECHEC : Controle des groupes Admins : Liste Vide : case " & compteur, True)
            End If
        Next

    End Sub

    Public Shared Function IsPrimaryGroup(ByVal GroupDE As DirectoryEntry, ByVal AccountPrimaryGroupID As Integer) As Boolean
        Try

            Dim GroupIDBytes() As Byte = DirectCast(GroupDE.Properties("objectSid").Value, Byte())
            Dim GroupSID As New Security.Principal.SecurityIdentifier(GroupIDBytes, 0)
            Dim SplitSID() As String = GroupSID.Value.Split("-"c)
            Return AccountPrimaryGroupID = CInt(SplitSID(SplitSID.Length - 1))
        Catch ex As Exception
            Throw New ApplicationException("Error comparing SID of group to account primary group ID: " & ex.Message)
        End Try
    End Function
    Public Function UserExist(ByVal CNAVerifier As String) As String

        Dim monEntry As New DirectoryEntry("LDAP://OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim maRecherche As DirectorySearcher = New DirectorySearcher
        Dim resultat As String = ""

        Try

            maRecherche.Filter = "(&(objectClass=user) (CN=" + CNAVerifier + "))"
            maRecherche.SearchScope = SearchScope.Subtree

            Dim result As SearchResult = maRecherche.FindOne()
            monEntry.Close()
            monEntry = Nothing

            If Not result Is Nothing Then
                Return result.Path
            End If
            maRecherche = Nothing
        Catch ex As Exception

        End Try

        Return resultat

    End Function
    Sub NbUserDatabaseExchange()

        Dim Ldap As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")
        Dim dirSearcher As DirectorySearcher = New DirectorySearcher(Ldap)
        dirSearcher.SearchScope = SearchScope.Subtree


        tabDatabase(0, 0) = "DB2-F21a"
        tabDatabase(0, 1) = "DB2-F21b"
        tabDatabase(0, 2) = "DB2-F21c"
        tabDatabase(0, 3) = "DB3-G31a"
        tabDatabase(0, 4) = "DB3-G31b"
        tabDatabase(0, 5) = "DB3-G31c"
        tabDatabase(0, 6) = "DB4-H13a"
        tabDatabase(0, 7) = "DB4-H13b"
        tabDatabase(0, 8) = "DB4-H13c"

        For i = 0 To UBound(tabDatabase, 2)
            dirSearcher.Filter = "(&(homeMDB=CN=" & tabDatabase(0, i) & ",CN=Databases,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=First Organization,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=igbmc,DC=u-strasbg,DC=fr))"
            Dim results As SearchResultCollection = dirSearcher.FindAll
            Dim j As Integer = 0
            For Each utilisateur In results
                j += 1
            Next
            tabDatabase(1, i) = j
        Next i

    End Sub
    Function TrouverPetiteDatabase() As String
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
    Shared Function CreatUnixUser(ByVal nom As String, ByVal prenom As String, ByVal login As String, ByVal passwd As String, ByVal nomEQ As String, ByVal homedirectory As String, ByVal GIDNumber As Integer) As String
        Dim LdapU As DirectoryEntry = New DirectoryEntry("LDAP://130.79.78.178:1389/dc=igbmc,dc=fr", "uid=steph, ou=People, dc=igbmc,dc=fr", "testldap", AuthenticationTypes.ServerBind)
        Dim ResultFields() As String = {"uidNumber"}
        Dim searcherLdapU As DirectorySearcher = New DirectorySearcher(LdapU)
        Dim CtrlUIDNumberLibre As Boolean = False
        Dim i As Integer = -1
        Dim nouveauUIDNumber As Integer

        searcherLdapU.Filter = "(&(objectClass=posixAccount))"
        searcherLdapU.SearchScope = SearchScope.Subtree
        searcherLdapU.PropertiesToLoad.AddRange(ResultFields)
        Dim tabUIDN() As Integer
        Dim resultLdapU As SearchResultCollection = searcherLdapU.FindAll()
        For Each result As SearchResult In resultLdapU
            i += 1
            ReDim Preserve tabUIDN(i)
            tabUIDN(i) = result.Properties("uidNumber")(0).ToString
        Next
        Dim k = 900
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



        Dim homeDirectoryUnix = Replace(homedirectory, "\\", "/")
        homeDirectoryUnix = Replace(homeDirectoryUnix, "\", "/")

        'creation de l'utilisateur sur le LDAP 
        LdapU = New DirectoryEntry("LDAP://130.79.78.178:1389/ou=People,dc=igbmc,dc=fr", "uid=steph, ou=People, dc=igbmc,dc=fr", "testldap", AuthenticationTypes.ServerBind)

        Dim objUserUnix As DirectoryEntry = LdapU.Children.Add("uid=" & login, "inetOrgPerson")
        Commun.SetADLDAPProperty(objUserUnix, "cn", prenom & " " & UCase(nom))
        Commun.SetADLDAPProperty(objUserUnix, "sn", UCase(nom))

        Dim tabClass(3) As String
        tabClass(0) = "top"
        tabClass(1) = "shadowAccount"
        tabClass(2) = "posixAccount"
        tabClass(3) = "inetOrgPerson"

        objUserUnix.Properties("objectClass").AddRange(tabClass)
        Commun.SetADLDAPProperty(objUserUnix, "gidNumber", GIDNumber)
        Commun.SetADLDAPProperty(objUserUnix, "homeDirectory", homeDirectoryUnix)
        Commun.SetADLDAPProperty(objUserUnix, "uidNumber", nouveauUIDNumber)
        objUserUnix.CommitChanges()

        Commun.SetADLDAPProperty(objUserUnix, "loginShell", "/bin/tcsh")
        Commun.SetADLDAPProperty(objUserUnix, "description", homeDirectoryUnix)
        Commun.SetADLDAPProperty(objUserUnix, "givenName", prenom)
        objUserUnix.CommitChanges()

        Commun.SetADLDAPProperty(objUserUnix, "userPassword", passwd)
        objUserUnix.CommitChanges()

        Return objUserUnix.Properties("uidNumber").Value & "," & objUserUnix.Properties("loginShell").Value & "," & objUserUnix.Properties("homeDirectory").Value

        objUserUnix.Close()
        objUserUnix = Nothing
        LdapU.Close()
        LdapU = Nothing

    End Function
    Sub DelUnixUser(ByVal login As String)
        Try
            'Dim LdapUser As DirectoryEntry = New DirectoryEntry("LDAP://130.79.78.178:1389/uid=" & login & ",ou=People,dc=igbmc,dc=fr", "uid=steph, ou=People, dc=igbmc,dc=fr", "testldap", AuthenticationTypes.ServerBind)
            Dim LdapU As DirectoryEntry = New DirectoryEntry("LDAP://130.79.78.178:1389/dc=igbmc,dc=fr", "uid=steph, ou=People, dc=igbmc,dc=fr", "testldap", AuthenticationTypes.ServerBind)
            Dim searcherLdapU As DirectorySearcher = New DirectorySearcher(LdapU)
            searcherLdapU.Filter = "(&(objectClass=posixAccount)(uid=" & login & "))"
            searcherLdapU.SearchScope = SearchScope.Subtree
            Dim resultLdapU As SearchResult = searcherLdapU.FindOne()
            If Not resultLdapU Is Nothing Then
                resultLdapU.GetDirectoryEntry.DeleteTree()
            End If

            LdapU.Close()
            LdapU = Nothing
        Catch ex As Exception
            Commun.Journal("ECHEC : Suppression du compte Unix : " & ex.Message & " : " & login, True)
        End Try

    End Sub
    Sub commandePWSMailbox(ByVal aliasMail As String, ByVal db As String)
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

            '-- set credentials      
            pCredential = DirectCast(Nothing, PSCredential) 'New PSCredential("igbmc\steph", CreateSecurePasswordString("aa676665"))

            '-- set connection info
            pConnectionInfo = New WSManConnectionInfo(New Uri("http://serv-cashub1/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", pCredential)

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
                .AddParameter("identity", aliasMail)
                .AddParameter("alias", aliasMail)
                .AddParameter("Database", db)
                .AddParameter("DomainController", ctrlDomain)
            End With

            '-- add command to powershell
            pShell.Commands = pCommand

            '-- invoke the powershell
            pResult = pShell.Invoke

            pCommand1 = New PSCommand

            With pCommand1
                .AddCommand("set-CASMailbox")
                .AddParameter("identity", aliasMail)
                .AddParameter("ActiveSyncEnabled", True)
                .AddParameter("ImapEnabled", False)
                .AddParameter("PopEnabled", False)
                .AddParameter("DomainController", ctrlDomain)
            End With

            pShell.Commands = pCommand1
            pResult1 = pShell.Invoke
            pRunspace.Close()
        Catch e As Exception
            Commun.Journal("ECHEC : Creation de compte : Creation du compte mail : " & e.Message & " : " & aliasMail, True)
        End Try

    End Sub
    Function PresenceExceptionUser(ByVal login As String) As Boolean
        'verifie la presence de l'utilisateur dans le fichier des exceptions, pour ne pas supprimer son compte
        Dim resultat As Boolean = False
        Try
            If tabExcepUser Is Nothing Then
                Dim monStreamReader As StreamReader = New StreamReader("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\UsersExceptions.txt")
                Dim ligne As String
                Dim i As Integer = -1
                Do
                    ligne = monStreamReader.ReadLine()
                    If Not ligne Is Nothing Then
                        Dim tabLigneExceptUser As String() = Split(ligne, ",")
                        i += 1
                        ReDim Preserve tabExcepUser(i)
                        tabExcepUser(i) = tabLigneExceptUser(0)
                    End If
                Loop Until ligne Is Nothing
                monStreamReader.Close()
            End If

            If Not tabExcepUser Is Nothing Then
                If Array.IndexOf(tabExcepUser, login) = -1 Then
                    resultat = False
                Else
                    resultat = True
                End If
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : PresenceExceptionUser : login : " & login & " : " & ex.Message, True)
        End Try

        Return resultat

    End Function
    Shared Sub commandSSH(ByVal nomZone As String)
        'Abolument remplacer les - par des _ dans le nom du volume
        Dim nomVolume As String = Replace(nomZone, "-", "_")
        Dim ssh As SshExec = New Tamir.SharpSsh.SshExec("nxprod-adm", "admin", "Stor@ge!igbmc") '("serv-nagios2", "root", "Sc&infO")
        Dim pattern As String = "root@serv-nagios2:~#" '"<NXPROD> CLI >"

        ssh.Connect()
        Dim command As String = "access nas-volumes add " & nomVolume & " 1.5 TB -security_style MIXED"
        Dim Res = ssh.RunCommand(command)
        command = "data-protection snapshots policies set-hourly-policy " & nomVolume & " on -hourlyat 8,10,12,14,16,18,20 -hourlykeep 8"
        Res = ssh.RunCommand(command)
        command = "data-protection snapshots policies set-daily-policy " & nomVolume & " on -dailyat 0,1,2,3,4,5 -dailystarttimehour 8 -dailykeep 7"
        Res = ssh.RunCommand(command)
        command = "data-protection snapshots policies set-weekly-policy " & nomVolume & " on -weeklyday 5 -weeklystarttimehour 20 -weeklykeep 4"
        Res = ssh.RunCommand(command)
        command = "data-protection replication nas-replication add " & nomVolume & " NXREPLICA"
        Res = ssh.RunCommand(command)
        command = "access cifs-shares add " & nomZone & " " & nomVolume & " /" & nomZone & " -create_dir_on_the_fly"
        Res = ssh.RunCommand(command)

        ssh.Close()
    End Sub
    'Sub commandePWSMailbox1(ByVal aliasMail As String)


    '    Dim pCredential As PSCredential
    '    Dim pConnectionInfo As WSManConnectionInfo
    '    Dim pRunspace As Runspace
    '    Dim pShell As PowerShell
    '    Dim pCommand As PSCommand
    '    Dim pCommand1 As PSCommand
    '    Dim pResult As Collection(Of PSObject)
    '    Dim pResult1 As Collection(Of PSObject)

    '    '-- set credentials      
    '    pCredential = DirectCast(Nothing, PSCredential) 'New PSCredential("igbmc\steph", CreateSecurePasswordString("aa676665"))

    '    '-- set connection info
    '    pConnectionInfo = New WSManConnectionInfo(False, "serv-lync13.igbmc.u-strasbg.fr", 5985, "WSMan", "http://schemas.microsoft.com/powershell/Microsoft.Powershell", pCredential)

    '    '-- create remote runspace
    '    pRunspace = RunspaceFactory.CreateRunspace(pConnectionInfo)
    '    pRunspace.Open()

    '    '-- create powershell
    '    pShell = PowerShell.Create
    '    pShell.Runspace = pRunspace

    '    '-- create command
    '    pCommand = New PSCommand
    '    With pCommand
    '        .AddCommand("get-csuser")
    '        .AddParameter("identity", aliasMail)
    '        '.AddParameter("alias", aliasMail)
    '        '.AddParameter("Database", db)
    '        '.AddParameter("DomainController", "serv-ad2")
    '    End With

    '    '-- add command to powershell
    '    pShell.Commands = pCommand

    '    '-- invoke the powershell
    '    pResult = pShell.Invoke
    'End Sub
End Class
