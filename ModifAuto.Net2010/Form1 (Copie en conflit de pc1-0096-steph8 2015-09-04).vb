Imports System.DirectoryServices
Imports System.IO
Imports System.Net.Mail
Imports ActiveDs
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Web.Security

Public Class Form1

    Shared withJson As String = "json" 'Valeur possible : json ftp debug temp(fichiers dans le dossier c:\temp)
    Shared tabPersoMonoEquipe As String(,)
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Shared tabExcepUser(,) As String
    Shared destinationsJsonIsTrue

    Shared Sub exec(ByVal cmd As String, Optional ByVal argu As String = "")
        Dim p As New Process
        p.StartInfo.UseShellExecute = False
        p.StartInfo.CreateNoWindow = True

        p.StartInfo.Arguments = argu
        p.StartInfo.FileName = cmd
        p.Start()
        p.WaitForExit(5000)
    End Sub

    Shared Sub Main()
        ExpirationMDP()

        SansContrat()
        'Dim mail = MailCloture("Veronique", "03/12/2015")
        'Commun.SendEmail("noreply@igbmc.fr", "parasote@igbmc.fr" & ";Cc:serviceinfo@igbmc.fr", "ARRET DU COMPTE", mail)

        'Dim listExtensionsXivo As String = Json.SendJson("", "extensions", "Xivo", "GET")
        'Dim donnesUNIXLdap As String = DonneesUnix("fillioll", "laverie-igbmc_eq")
        If Environment.MachineName = "SERV-AD1" Then withJson = "json"

        Commun.fichierLog = "c:\temp\Log" & Application.ProductName & ".log"
        Commun.Journal("Debut de traitement")


        Dim ctrlCreationFichier As Boolean = CreationFichiers()

        If ctrlCreationFichier = True Then
            GestionDestinationsDepartements()
            modifDonneesAD()

            If Environment.MachineName = "SERV-AD1" Then
                ReactiveDesactiveCompte()
            End If

            ' UPLOAD DES FICHIERS PAR FTP
            Try

                'Commun.UploadFileFTP("c:\temp\eq.txt", "/" & RecupDataini.RecupVar("[FCCFilePath]") & "/eq.txt")
                'Commun.UploadFileFTP("c:\temp\listep.txt", "/" & RecupDataini.RecupVar("[FCCFilePath]") & "/listep.txt")
                File.Copy("c:\temp\eq.txt", RecupDataini.RecupVar("[CheminPartage]") & "\todo\eq.txt", True)
                File.Copy("c:\temp\listep.txt", RecupDataini.RecupVar("[CheminPartage]") & "\todo\listep.txt", True)
                If withJson = "json" Then
                    'Commun.UploadFileFTP("c:\temp\listepersoJson.txt", "/" & RecupDataini.RecupVar("[FCCFilePath]") & "/listepersoJson.txt")
                    File.Copy("c:\temp\listepersoJson.txt", RecupDataini.RecupVar("[CheminPartage]") & "\todo\listepersoJson.txt", True)
                    File.Copy("c:\temp\listepersoJson.txt", RecupDataini.RecupVar("[CheminPartage]") & "\cmpttmp\listepersoJson" & Now.ToString("dd") & "-" & Now.ToString("MM") & "-" & Now.ToString("yy") & "-" & Now.ToString("HH") & "h.txt", True)
                End If
            Catch ex As Exception
                Commun.Journal("ERREUR : FTP : Upload du fichier : " & ex.Message, True)
            End Try

            Sleep(5000)

            Try
                If withJson <> "ftp" Then
                    'Kill("c:\temp\listepersoJson.txt")
                End If
            Catch
                Commun.Journal("ERREUR : suppression du fichier temporaire listepersoJson.txt", True)
            End Try
            Try
                If withJson = "ftp" Then
                    Kill("c:\temp\listep1.txt")
                End If
            Catch
                Commun.Journal("ERREUR : suppression du fichier temporaire listep1.txt", True)
            End Try
            Try
                Kill("c:\temp\listep.txt")
            Catch
                Commun.Journal("ERREUR : suppression du fichier temporaire listep.txt", True)
            End Try
            Try
                'Kill("c:\temp\eq.txt")
            Catch
                Commun.Journal("ERREUR : suppression du fichier temporaire eq.txt", True)
            End Try

        End If

        If Hour(Now) = 1 Or Hour(Now) = 2 Then
            'Gestion de l'expiration des mot de passe des comptes adm
            ExpirationMDP()

            '15 jours avant la suppression
            EnvoiMailCompteExpireXjour(15)

            'SansContrat()

            'gestion de l'envoi des mails pour les comptes désactivé pour leur suppression
            '1 jour apres la désactivation
            'EnvoiMailCompteExpireXjour(DateDiff("d", Now, DateTime.UtcNow.AddMonths(3)))

            GroupesDynamiques()

            'Changement tous les mois pair, le premier du mois, des mots de passe des comptes prestataires de l'imagerie
            If System.DateTime.Now.ToString("dd") = "01" And EstPair(System.DateTime.Now.ToString("MM")) = True Then
                ChangePasswordAccountPrestaImagerie()
            End If

            UpdateFichierHistoAlias()
        End If


        If Commun.controlSendMail = True Then
            Commun.SendEmail(RecupDataini.RecupVar("[AdminScriptLogin]") & "@igbmc.fr", "steph@igbmc.fr", "ModifAuto.NET : Rapport d'erreur", Commun.journalECHECMail)
        End If

        Commun.Journal("Fin de traitement" & vbCrLf)
        'CreationAuto.start()
    End Sub
    Shared Sub modifDonneesAD()
        'COMPARAISON DU PERSONNEL ENTRE HIER ET AUJOURDH'HUI ET CREATION DU FICHIER DE MODIFICATIONS


        Dim Ldap As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")

        Dim ancienEQ As String = ""

        Dim test As String = ""

        Dim listExtensionsXivo As String = ""
        Try
            listExtensionsXivo = Json.SendJson("", "extensions", "Xivo", "GET")
        Catch ex As Exception
            Commun.Journal("ERREUR : modifDonneesAD : Recuperation des données Xivo", True)
        End Try
        Dim usrTempsTavailTemp As Integer = 100

        Dim objuser As DirectoryEntry





        For n = 0 To UBound(tabPersoMonoEquipe, 2)


            Dim usrNom As String = tabPersoMonoEquipe(0, n)
            Dim usrPrenom As String = tabPersoMonoEquipe(1, n)
            Dim usrLogin As String = tabPersoMonoEquipe(2, n)
            Dim usrDestNomCourt As String = tabPersoMonoEquipe(3, n)
            Dim usrDestNomLong As String = tabPersoMonoEquipe(4, n)
            Dim usrTempsTravail As Integer = Convert.ToInt32(tabPersoMonoEquipe(5, n))
            Dim usrBureaux As String = tabPersoMonoEquipe(6, n)
            Dim usrTelephones As String = tabPersoMonoEquipe(7, n)
            Dim usrEquipeInfo As String = tabPersoMonoEquipe(8, n)
            Dim usrChef As String = tabPersoMonoEquipe(9, n)
            Dim CNChef As String = ""
            CNChef = Commun.TransformeSAMACCOUNTenCN(usrChef)
            Dim usrID As String = tabPersoMonoEquipe(10, n)
            Dim usrAliasMailLong As String = tabPersoMonoEquipe(11, n) 'avec @igbmc.fr


            Try
                'Si l'utilistateur n'existe pas, on continue avec le suivant
                If Not DirectoryEntry.Exists("LDAP://" & Commun.TransformeSAMACCOUNTenCN(usrLogin)) Then
                    Commun.Journal("ERREUR : Modification des données : " & usrLogin & " : L'Utilisateur n'existe pas", True)
                    Continue For
                End If

                Dim searcher As DirectorySearcher = New DirectorySearcher(Ldap)
                searcher.Filter = "(&(objectClass=user) (SAMAccountName=" & usrLogin & "))"
                Dim result As SearchResult = searcher.FindOne()

                objuser = result.GetDirectoryEntry()

                result = Nothing
                searcher.Dispose()
                searcher = Nothing

                'recuperation de l'ancienne equipe comptable
                If objuser.Properties("DepartmentNumber").Value <> "" Then
                    ancienEQ = objuser.Properties("DepartmentNumber").Value
                End If
                'On Error Resume Next


                'Gestion des Thumbnails
                'Thumbn.ComparePhoto(TransformeSAMACCOUNTenCN(usrLogin), usrID, objuser)
                Thumbn.ComparePhoto(usrID, objuser)

                'Gestion des données basiques
                Try
                    If ((objuser.Properties("SN").Value <> usrNom) Or (objuser.Properties("givenName").Value <> usrPrenom) Or (objuser.Properties("physicalDeliveryOfficeName").Value <> usrBureaux)) Then
                        objuser.Properties("SN").Value = usrNom
                        objuser.Properties("givenName").Value = usrPrenom


                        objuser.Properties("mail").Value = usrLogin & "@igbmc.fr"

                        If usrBureaux = "" Then
                            objuser.Properties("physicalDeliveryOfficeName").Clear()
                        Else
                            Commun.SetADLDAPProperty(objuser, "physicalDeliveryOfficeName", usrBureaux)
                        End If
                        Commun.AppliquerChangement(objuser)
                        Commun.Journal("Changement des Données basiques Réussi : " & usrLogin)
                    End If
                Catch ex As Exception
                    Commun.Journal("ERREUR : Modification des données basiques : " & usrLogin & " : " & ex.Message, True)
                End Try

                'Mise a jour du mail principal
                Try
                    Dim mailPrincipalcourt = TrouverMailPrincipal(usrLogin, "court")
                    If mailPrincipalcourt <> objuser.Properties("mail").Value Then
                        objuser.Properties("mail").Value = mailPrincipalcourt
                        Commun.AppliquerChangement(objuser)
                    End If
                Catch ex As Exception
                    Commun.Journal("ERREUR : Modification du mail principal : " & usrLogin & " : " & ex.Message, True)
                End Try


                'Données Téléphonique
                Dim tabPhoneFichier As String() = Split(usrTelephones, ";")
                Dim telephonePrincipal As String = tabPhoneFichier(0)
                Dim IPphone As String = telephonePrincipal
                Try
                    Dim tabOtherphoneAD() As String
                    If Not objuser.Properties("otherTelephone").Value Is Nothing Then
                        ReDim tabOtherphoneAD(objuser.Properties("otherTelephone").Count - 1)
                        objuser.Properties("otherTelephone").CopyTo(tabOtherphoneAD, 0)
                    End If

                    's'il y a des numeros dans le fichier
                    If UBound(tabPhoneFichier) > 0 Then
                        Dim tabOtherphoneFichier As String()

                        For h = 1 To UBound(tabPhoneFichier)
                            ReDim Preserve tabOtherphoneFichier(h - 1)
                            tabOtherphoneFichier(h - 1) = tabPhoneFichier(h)
                        Next h


                        'si des numeros sont deja présents dans l'AD
                        If Not tabOtherphoneAD Is Nothing Then

                            'si le nombre de numero de telephone n'est pas le meme dans le fichier que dans l'AD
                            If UBound(tabOtherphoneFichier) <> UBound(tabOtherphoneAD) Then
                                'on remplace la valeur de l'AD
                                objuser.Properties("otherTelephone").Value = tabOtherphoneFichier
                                Commun.AppliquerChangement(objuser)
                                Commun.Journal("Changement d'attribut ""OtherTelephone"" Réussi : " & usrLogin)
                                'Si le nombre de numero est egal
                            Else
                                'on verifie qu'un des numeros a changé
                                For Each valeur As String In tabOtherphoneAD
                                    Dim index As Integer = Array.IndexOf(tabOtherphoneFichier, valeur)
                                    'Si un a changé, on remplace tout
                                    If index = -1 Then
                                        objuser.Properties("otherTelephone").Value = tabOtherphoneFichier
                                        Commun.AppliquerChangement(objuser)
                                        Commun.Journal("Changement d'attribut ""OtherTelephone"" Réussi : " & usrLogin)
                                    End If
                                Next valeur
                            End If
                            's'il n'y a pas de numero dans l'AD
                        Else
                            'on remplace tout par les valeurs du fichier
                            objuser.Properties("otherTelephone").Add(" ")
                            objuser.Properties("otherTelephone").Value = tabOtherphoneFichier
                            Commun.AppliquerChangement(objuser)
                            Commun.Journal("Creation de l'attribut ""OtherTelephone"" Réussi : " & usrLogin)
                        End If
                        tabOtherphoneFichier = Nothing
                        tabOtherphoneAD = Nothing
                        's'il n'y a pas de numero dans le fichier,
                    Else
                        'on nettoie l'attribut dans l'AD
                        If objuser.Properties.Contains("otherTelephone") Then
                            objuser.Properties("otherTelephone").Clear()
                            Commun.AppliquerChangement(objuser)
                            Commun.Journal("Suppression de l'attribut ""OtherTelephone"" Réussi : " & usrLogin)
                        End If
                    End If
                    tabPhoneFichier = Nothing

                    IPphone = Replace(IPphone, "(+3336948) ", "")
                    IPphone = Replace(IPphone, "(+3338865) ", "")
                    IPphone = Replace(IPphone, ")", "")
                    IPphone = Replace(IPphone, "+33 ", "0")
                    If Len(IPphone) = 4 Then
                        'si le numero de poste est trouvé dans le xivo
                        If InStr(listExtensionsXivo, """exten"": """ & IPphone & """") <> 0 Then
                            'on efface l'attribut IPphone
                            IPphone = ""
                        End If
                    End If
                Catch ex As Exception
                    Commun.Journal("ERREUR : Traitement des données téléphoniques " & usrLogin & " : " & ex.Message, True)
                End Try


                Try
                    If objuser.Properties("TelephoneNumber").Value <> telephonePrincipal Then
                        If telephonePrincipal = "" Then
                            objuser.Properties("TelephoneNumber").Clear()
                        Else
                            Commun.SetADLDAPProperty(objuser, "TelephoneNumber", telephonePrincipal)
                            'objuser.Properties("TelephoneNumber").Add(telephonePrincipal)
                        End If
                        Commun.AppliquerChangement(objuser)
                        Commun.Journal("Modification de l'attribut ""TelephoneNumber"" Réussi : " & usrLogin)
                    End If
                Catch ex As Exception
                    Commun.Journal("ERREUR : Modification des données téléphoniques ""TelephoneNumber : " & usrLogin & " : " & telephonePrincipal & " : " & ex.Message, True)
                End Try

                Try
                    If objuser.Properties("IPphone").Value <> IPphone Then
                        If IPphone = "" Or IPphone = Nothing Then
                            objuser.Properties("IpPhone").Clear()
                        Else
                            Commun.SetADLDAPProperty(objuser, "IpPhone", IPphone)
                        End If
                        Commun.AppliquerChangement(objuser)
                        Commun.Journal("Modification de l'attribut ""IpPhone"" Réussi : " & usrLogin)
                    End If
                Catch ex As Exception
                    Commun.Journal("ERREUR : Modification des données téléphoniques ""IPPhone : " & usrLogin & " : " & IPphone & " : " & ex.Message, True)
                End Try

                'gestion des données de service
                Try
                    If (objuser.Properties("Company").Value <> "IGBMC") Or (objuser.Properties("Department").Value <> usrDestNomLong) Or (objuser.Properties("DepartmentNumber").Value <> usrDestNomCourt) Then

                        objuser.Properties("Company").Value = "IGBMC"
                        objuser.Properties("Department").Value = usrDestNomLong
                        objuser.Properties("DepartmentNumber").Value = usrDestNomCourt
                        Commun.AppliquerChangement(objuser)
                        Commun.Journal("Modification de l'attribut ""Compagny"",""Department"" et ""DepartmentNumber"" Réussi : " & usrLogin)

                    End If
                Catch
                    Commun.Journal("ERREUR : Modification : Departement de l'utilisateur : " & usrLogin, True)
                End Try

                'Données UNIX
                Try
                    Dim uidAD As String = objuser.Properties("SAMAccountName").Value
                    Dim uidNumberAD As String = objuser.Properties("uidnumber").Value
                    Dim gidNumberAD As String = Commun.FindAttribut(usrEquipeInfo & "_eq", "gidNumber")
                    Dim cheminZoneLabo As String = Commun.FindAttribut(usrEquipeInfo & "_eq", "url")
                    Dim unixHomeDirectoryAD As String = Replace(Replace(cheminZoneLabo, "\\", "\"), "\", "/") & "/" & usrLogin
                    If InStr(cheminZoneLabo, "labo4") > 0 Then
                        unixHomeDirectoryAD = Replace(cheminZoneLabo, "\\", "\")
                        unixHomeDirectoryAD = Replace(unixHomeDirectoryAD, "\", "/")
                        unixHomeDirectoryAD = Replace(unixHomeDirectoryAD, "-", "_")
                        unixHomeDirectoryAD = unixHomeDirectoryAD & "/" & usrEquipeInfo & "/" & usrLogin
                    End If
                    Dim loginShellAD As String = objuser.Properties("loginShell").Value

                    If objuser.Properties("uid").Value <> uidAD Then
                        Commun.SetADLDAPProperty(objuser, "uid", uidAD)
                        Commun.AppliquerChangement(objuser)
                        Commun.Journal("Modification de l'attribut AD ""uid"" réussi : " & usrLogin)
                    End If

                    If objuser.Properties("uidNumber").Value <> uidNumberAD Then
                        Commun.SetADLDAPProperty(objuser, "uidNumber", uidNumberAD)
                        Commun.AppliquerChangement(objuser)
                        Commun.Journal("Modification de l'attribut AD ""uidNumber"" réussi : " & usrLogin)
                    End If

                    If objuser.Properties("gidNumber").Value <> gidNumberAD Then
                        Commun.SetADLDAPProperty(objuser, "gidNumber", gidNumberAD)
                        Commun.AppliquerChangement(objuser)
                        Commun.Journal("Modification de l'attribut AD ""gidNumber"" réussi : " & usrLogin)
                    End If

                    If objuser.Properties("unixHomeDirectory").Value <> unixHomeDirectoryAD Then
                        Commun.SetADLDAPProperty(objuser, "unixHomeDirectory", unixHomeDirectoryAD)
                        Commun.AppliquerChangement(objuser)
                        Commun.Journal("Modification de l'attribut AD ""unixHomeDirectory"" réussi : " & usrLogin)
                    End If
                    If objuser.Properties("loginShell").Value <> loginShellAD Then
                        Commun.SetADLDAPProperty(objuser, "loginShell", unixHomeDirectoryAD)
                        Commun.AppliquerChangement(objuser)
                        Commun.Journal("Modification de l'attribut AD ""loginShellAD"" réussi : " & usrLogin)
                    End If

                    Dim donneesUNIXAd As String = uidNumberAD & "," & gidNumberAD & "," & unixHomeDirectoryAD & "," & loginShellAD
                    '4 valeurs Unix récupérées : uidnumber,gidnumber,homedirectory,loginShell
                    If usrLogin <> "" And usrEquipeInfo <> "" Then
                        Dim donnesUNIXLdap As String = DonneesUnix(usrLogin, usrEquipeInfo)
                        If donneesUNIXAd <> donnesUNIXLdap Then
                            Dim objUserUnix As DirectoryEntry = New DirectoryEntry("LDAP://130.79.78.178:1389/uid=" & uidAD & ",ou=People,dc=igbmc,dc=fr", "uid=" & RecupDataini.RecupVar("[AdminScriptLogin]") & ", ou=People, dc=igbmc,dc=fr", RecupDataini.RecupVar("[AdminScriptPassword]"), AuthenticationTypes.ServerBind)
                            Dim tabDonneesUnix As String() = Split(donnesUNIXLdap, ",")
                            If uidNumberAD <> tabDonneesUnix(0) Then
                                Try
                                    Commun.SetADLDAPProperty(objUserUnix, "uidnumber", uidNumberAD)
                                    'Commun.SetADLDAPProperty(objuser, "uidnumber", tabDonneesUnix(0))
                                    Commun.Journal("Modification des données Unix sur le LDAP ""uidnumber"" réussi : " & usrLogin)
                                Catch ex As Exception
                                    Commun.Journal("ERREUR : Modification des attributs Unix ""uidnumber"" : " & usrLogin & " : " & ex.Message, True)
                                End Try
                            End If
                            If gidNumberAD <> tabDonneesUnix(1) Then
                                Try
                                    Commun.SetADLDAPProperty(objUserUnix, "gidnumber", gidNumberAD)
                                    'Commun.SetADLDAPProperty(objuser, "gidnumber", tabDonneesUnix(1))
                                    Commun.AppliquerChangement(objUserUnix)
                                    logSerge(Now() & " : " & usrLogin & " : gidNumber     : " & gidNumberAD & " (" & tabDonneesUnix(1) & ")" & vbCrLf)
                                    Commun.Journal("Modification des données Unix sur le LDAP ""gidnumber"" réussi : " & usrLogin)
                                Catch ex As Exception
                                    Commun.Journal("ERREUR : Modification des attributs Unix ""gidnumber"" : " & usrLogin & " : " & ex.Message, True)
                                End Try
                            End If
                            If unixHomeDirectoryAD <> tabDonneesUnix(2) Then
                                Try
                                    Commun.SetADLDAPProperty(objUserUnix, "homeDirectory", unixHomeDirectoryAD)
                                    'Commun.SetADLDAPProperty(objuser, "unixHomeDirectory", tabDonneesUnix(2))
                                    Commun.AppliquerChangement(objUserUnix)
                                    logSerge(Now() & " : " & usrLogin & " : homeDirectory : " & unixHomeDirectoryAD & " (" & tabDonneesUnix(2) & ")" & vbCrLf)
                                    Commun.Journal("Modification des données Unix sur le LDAP ""homeDirectory"" réussi : " & usrLogin)
                                Catch ex As Exception
                                    Commun.Journal("ERREUR : Modification des attributs Unix ""homeDirectory"" : " & usrLogin & " : " & ex.Message, True)
                                End Try
                            End If
                            If loginShellAD <> tabDonneesUnix(3) Then
                                If loginShellAD Is Nothing Then loginShellAD = "/bin/tcsh"
                                'objUserUnix.Properties("loginShell").Add(loginShellAD)
                                Try
                                    Commun.SetADLDAPProperty(objUserUnix, "loginShell", loginShellAD)
                                    'Commun.SetADLDAPProperty(objuser, "loginShell", loginShellAD)
                                    Commun.AppliquerChangement(objUserUnix)
                                    Commun.Journal("Modification des données Unix sur le LDAP ""loginshell"" réussi : " & usrLogin)
                                Catch ex As Exception
                                    Commun.Journal("ERREUR : Modification des attributs Unix ""loginshell"" : " & usrLogin & " : " & ex.Message, True)
                                End Try
                            End If
                            objUserUnix.Close()
                            objUserUnix.Dispose()
                            objUserUnix = Nothing
                            tabDonneesUnix = Nothing
                        End If
                    End If
                Catch ex As Exception
                    Commun.Journal("ERREUR : Modification des attributs Unix de l'utilisateur : " & usrLogin & " : " & ex.Message, True)
                End Try


                'Chef d'équipe
                Try

                    If CNChef <> "" Then

                        If objuser.Properties("Manager").Value <> CNChef Then
                            objuser.Properties("Manager").Value = CNChef
                            Commun.AppliquerChangement(objuser)
                            Commun.Journal("Modification de l'attribut ""Manager"" Réussi : " & usrLogin)
                        End If
                        'On Error GoTo 0
                    End If
                Catch
                    Commun.Journal("ERREUR : Chef d'equipe de l'utilisateur : " & usrLogin & " : " & CNChef, True)
                End Try
            Catch ex As Exception
                Commun.Journal("ERREUR : Controle de modification de l'utilisateur : " & usrLogin, True)
            End Try

            'gestion des appartenance aux groupes de destinations (grp)
            CompararaisonAjoutRetraitDestinations(objuser)
            'Correction equipe

            'ajout et retrait du groupe ADMINISTRATIF en fonction de la nouvelle equipe
            Try

                Dim membreGroupADM As Boolean = Commun.AppartientGroup(objuser.Path, "G_ADMINISTRATIF_Users")
                Dim membreEquipeADM As Boolean = Commun.IsMembreEquipeAdministratif(objuser)

                If membreEquipeADM = True And membreGroupADM = False Then
                    Commun.AddRemoveADGroup(objuser.Path, "G_ADMINISTRATIF_Users", "Add")
                End If
                If membreEquipeADM = False And membreGroupADM = True Then
                    Commun.AddRemoveADGroup(objuser.Path, "G_ADMINISTRATIF_Users", "Remove")
                End If

            Catch
                Commun.Journal("ERREUR : Modification : Changer l'equipe de l'utilisateur : " & usrLogin, True)
            End Try


            'Modification du displayname
            Try
                If objuser.Properties("displayName").Value <> usrPrenom & " " & usrNom Or objuser.Properties("displayNamePrintable").Value <> usrNom & " " & usrPrenom Then
                    objuser.Properties("displayName").Value = usrPrenom & " " & usrNom
                    Commun.SetADLDAPProperty(objuser, "displayNamePrintable", usrNom & " " & usrPrenom)
                    Commun.AppliquerChangement(objuser)
                    Commun.Journal("Modification de l'attribut ""displayName"" Réussi : " & usrLogin)
                End If
            Catch ex As Exception
                Commun.Journal("ERREUR : Modification : changement displayName : " & usrLogin & " : " & ex.Message, True)
            End Try

            'Controle si l'utilisateur a changé de nom
            If objuser.Properties("cn").Value <> objuser.Properties("displayName").Value Then
                'Modification des alias de mails
                Try
                    modifAliasMail(objuser, usrPrenom, usrNom)
                Catch ex As Exception
                    Commun.Journal("ERREUR : Modification : Ajout d'Alias : " & usrLogin & " : " & ex.Message, True)
                End Try

                Try
                    Dim objDom As ActiveDs.IADsContainer = GetObject(objuser.Parent.Path)
                    objDom.MoveHere(objuser.Path, "cn=" & objuser.Properties("displayName").Value)
                    objDom = Nothing

                Catch ex As Exception
                    Commun.Journal("ERREUR : Modification : Renomer l'objet utilisateur : " & usrLogin & " : " & ex.Message, True)
                End Try
            End If

            'Controle et Modification de l'alias sur IGBMCSERVICES
            'Dim aliasmailAD As String = TrouverMailPrincipal(usrLogin, "long")
            'objuser.Properties("msExchExtensionCustomAttribute1").Clear()
            Dim aliasmailAD As String = LCase(objuser.Properties("msExchExtensionAttribute16").Value)
            If aliasmailAD <> usrAliasMailLong Then
                Try
                    Dim aliasMailpourIGBMCSERVICES As String = Replace(aliasmailAD, "@igbmc.fr", "")
                    Json.SendJson("login=" & objuser.Properties("SAMAccountName").Value & "&domain=%40igbmc.fr&alias=" & aliasMailpourIGBMCSERVICES, "persons/" & usrID & "/email", "AD", "PUT")
                    objuser.Properties("msExchExtensionAttribute16").Value = LCase(aliasmailAD)
                    Commun.AppliquerChangement(objuser)

                Catch ex As Exception
                    Commun.Journal("ERREUR : Modification alias Mail sur IGBMCSERVICES : " & objuser.Properties("SAMAccountName").Value & " : " & ex.Message, True)
                End Try
            End If


            objuser.Close()
            objuser.Dispose()
            objuser = Nothing

        Next n

        Ldap.Close()
        Ldap.Dispose()
        Ldap = Nothing
        Commun.Journal("Gestion des modifications utilisateurs réussie", False)
    End Sub
    Shared Function CreationFichiers()
        Dim result As Boolean = False
        If withJson = "ftp" Then

            CreationFichierParFTP()
            Commun.getFileFTP("/" & RecupDataini.RecupVar("[FCCFilePath]") & "/eq.txt", "c:\temp\eq.txt")
        ElseIf withJson = "json" Then
            Dim dataJsonDestinationTrue As String = Json.SendJson("", "destinations?is_team=true", "AD", "GET")
            If dataJsonDestinationTrue = Nothing Then End
            destinationsJsonIsTrue = Json.DeserializeJson(dataJsonDestinationTrue, "destinations")
            creationFichierParJson()
        ElseIf withJson = "debug" Then
            File.Copy(RecupDataini.RecupVar("[CheminPartage]") & "\todo\eq.txt", "c:\temp\eq.txt", True)
            File.Copy(RecupDataini.RecupVar("[CheminPartage]") & "\todo\listepersoJson.txt", "c:\temp\listepersoJson.txt", True)
            'Commun.getFileFTP("/" & RecupDataini.RecupVar("[FCCFilePath]") & "/listepersoJson.txt", "c:\temp\listepersoJson.txt")
            'Commun.getFileFTP("/" & RecupDataini.RecupVar("[FCCFilePath]") & "/eq.txt", "c:\temp\eq.txt")
        ElseIf withJson = "temp" Then

        Else
            End
        End If

        Try
            'DETERMINER L'EQUIPE PRINCIPALE
            If withJson = "ftp" Then 'Or withJson = "temp" Then
                FileOpen(1, "c:\temp\listep1.txt", OpenMode.Input) 'fichier du personnel sans accent multi equipe
            Else
                FileOpen(1, "c:\temp\listepersoJson.txt", OpenMode.Input) 'fichier du personnel sans accent multi equipe
            End If
            Dim ligneP As String
            Dim lignePerso As String()
            Dim champE As String()
            Dim ligneE As String
            Dim j As Integer

            Dim g As Integer = 0
            While Not EOF(1)

                ligneP = LineInput(1)
                lignePerso = Split(ligneP, ",")
                'evite les lignes vides
                If lignePerso.Length > 1 Then
                    ReDim Preserve tabPersoMonoEquipe(12, g)
                    tabPersoMonoEquipe(0, g) = lignePerso(0)
                    tabPersoMonoEquipe(1, g) = lignePerso(1)
                    tabPersoMonoEquipe(2, g) = lignePerso(2)
                    tabPersoMonoEquipe(3, g) = lignePerso(3)
                    tabPersoMonoEquipe(4, g) = lignePerso(4)
                    tabPersoMonoEquipe(5, g) = lignePerso(5)
                    tabPersoMonoEquipe(6, g) = lignePerso(6)
                    tabPersoMonoEquipe(7, g) = lignePerso(7)
                    tabPersoMonoEquipe(8, g) = lignePerso(8)
                    tabPersoMonoEquipe(9, g) = lignePerso(9)
                    tabPersoMonoEquipe(10, g) = lignePerso(10)
                    tabPersoMonoEquipe(11, g) = lignePerso(11)
                    'tabPersoMonoEquipe(12, g) = lignePerso(12)
                    lignePerso = Nothing

                    If tabPersoMonoEquipe Is Nothing Then Throw New Exception("tabPerso est vide")

                    If g <> 0 Then
                        'comparaison de 2 lignes du meme utilisateur (dans le cas d'equipe multiple) 
                        If tabPersoMonoEquipe(10, g) = tabPersoMonoEquipe(10, g - 1) Then
                            'Si le temps de travail dans une equipe, pour le meme utilisateur, est superieur a la ligne precedente, on reecrit le debut de la ligne precedente
                            If tabPersoMonoEquipe(5, g) > tabPersoMonoEquipe(5, g - 1) Then
                                tabPersoMonoEquipe(0, g - 1) = tabPersoMonoEquipe(0, g)
                                tabPersoMonoEquipe(1, g - 1) = tabPersoMonoEquipe(1, g)
                                tabPersoMonoEquipe(2, g - 1) = tabPersoMonoEquipe(2, g)
                                tabPersoMonoEquipe(3, g - 1) = tabPersoMonoEquipe(3, g)
                                tabPersoMonoEquipe(4, g - 1) = tabPersoMonoEquipe(4, g)
                                tabPersoMonoEquipe(5, g - 1) = tabPersoMonoEquipe(5, g)
                                tabPersoMonoEquipe(8, g - 1) = tabPersoMonoEquipe(8, g)
                                tabPersoMonoEquipe(9, g - 1) = tabPersoMonoEquipe(9, g)
                                g = g - 1
                            Else
                                g = g - 1
                            End If
                        End If
                    End If
                    g = g + 1
                End If
            End While

            FileClose(1)

            FileOpen(1, "c:\temp\listep.txt", OpenMode.Output) 'fichier du personnel sans accent mono equipe

            For i = 0 To g - 1
                tabPersoMonoEquipe(6, i) = Replace(tabPersoMonoEquipe(6, i), "BATIMENT ICS", "ICS")
                tabPersoMonoEquipe(6, i) = Replace(tabPersoMonoEquipe(6, i), "C.E.B.G.S", "CEBGS")
                tabPersoMonoEquipe(6, i) = Replace(tabPersoMonoEquipe(6, i), "E.S.B.S.", "ESBS")

                Dim tabTelTemp = Split(tabPersoMonoEquipe(7, i), ";")
                Dim telTemp As String = ""
                For j = 0 To UBound(tabTelTemp)
                    If tabTelTemp(j) <> "----" Then

                        If Len(tabTelTemp(j)) = 4 Then
                            Dim prefixTel As String = ""

                            If Strings.Left(tabTelTemp(j), 2) = "32" Or Strings.Left(tabTelTemp(j), 2) = "33" Or Strings.Left(tabTelTemp(j), 2) = "34" Or Strings.Left(tabTelTemp(j), 2) = "35" Or Strings.Left(tabTelTemp(j), 2) = "56" Or Strings.Left(tabTelTemp(j), 2) = "57" Then
                                prefixTel = "(+3338865) "
                            End If

                            If Strings.Left(tabTelTemp(j), 2) = "50" Or Strings.Left(tabTelTemp(j), 2) = "51" Or Strings.Left(tabTelTemp(j), 2) = "52" Then
                                prefixTel = "(+3336948) "
                            End If


                            telTemp = telTemp & ";" & prefixTel & tabTelTemp(j)
                        ElseIf Len(tabTelTemp(j)) = 10 And Strings.Left(tabTelTemp(j), 1) = "0" Then
                            'Remplacement uniquement du premier "0"
                            telTemp = telTemp & ";" & Replace(tabTelTemp(j), "0", "+33 ", 1, 1)
                        Else
                            telTemp = ";" & tabTelTemp(j)
                        End If
                    Else
                        telTemp = telTemp & ";"
                    End If
                Next j
                telTemp = Strings.Right(telTemp, Len(telTemp) - 1)
                tabPersoMonoEquipe(7, i) = telTemp
                PrintLine(1, tabPersoMonoEquipe(0, i) & "," & tabPersoMonoEquipe(1, i) & "," & tabPersoMonoEquipe(2, i) & "," & tabPersoMonoEquipe(3, i) & "," & tabPersoMonoEquipe(4, i) & "," & tabPersoMonoEquipe(5, i) & "," & tabPersoMonoEquipe(6, i) & "," & tabPersoMonoEquipe(7, i) & "," & tabPersoMonoEquipe(8, i) & "," & tabPersoMonoEquipe(9, i) & "," & tabPersoMonoEquipe(10, i) & "," & tabPersoMonoEquipe(11, i))
            Next i
            FileClose(1)


            'CREATION DE LA LISTE DES DESTINATIONS

            If withJson = "ftp" Then
                FileOpen(2, "c:\temp\eq.txt", OpenMode.Output)
                FileOpen(1, "c:\temp\listep1.txt", OpenMode.Input)

                j = 0
                Dim tableauE As String(,)
                While Not EOF(1)
                    Dim ctrl As Integer = 0
                    ligneE = LineInput(1)
                    champE = Split(ligneE, ",")
                    For i = 0 To j
                        ReDim Preserve tableauE(2, j)
                        If champE(3) = tableauE(0, i) Then
                            GoTo sortie1
                        Else
                            If ctrl = j Then
                                tableauE(0, j) = champE(3)
                                tableauE(1, j) = champE(4)
                                'chef d'equipe comptable
                                tableauE(2, j) = champE(9)
                                PrintLine(2, champE(3) & "," & champE(4) & "," & champE(9))
                                j = j + 1

                            End If
                            ctrl = ctrl + 1
sortie1:
                        End If
                    Next i
                End While

                FileClose(1)
                FileClose(2)
            End If
            result = True
        Catch ex As Exception
            Commun.Journal("ERREUR : Creation des fichiers et des tableaux : " & ex.Message, True)
            Commun.SendEmail(RecupDataini.RecupVar("[AdminScriptLogin]") & "@igbmc.fr", "steph@igbmc.fr", "ModifAuto.NET : Rapport d'erreur", Commun.journalECHECMail)
            Return result
            Exit Function
        End Try
        Return result
    End Function
    Shared Sub GestionDestinationsDepartements()
        'CREATION DU TABLEAU DES DESTINATIONS
        Dim tableauDest As String()
        FileOpen(1, "c:\temp\eq.txt", OpenMode.Input)

        Dim m As Integer = -1
        While Not EOF(1)
            Dim ligneDest As String = LineInput(1)
            'evite les lignes vides
            If InStr(ligneDest, ",") <> 0 Then
                m += 1
                ReDim Preserve tableauDest(m)
                tableauDest(m) = ligneDest
            End If
        End While

        FileClose(1)

        'CREATION DU TABLEAU DES DEPARTEMENTS
        Dim tableauDepartment As String()
        FileOpen(1, "c:\temp\eq.txt", OpenMode.Input)

        Dim n As Integer = -1
        While Not EOF(1)

            Dim ligneDepartment As String = LineInput(1)

            'evite les lignes vides
            If InStr(ligneDepartment, ",") <> 0 Then
                ligneDepartment = Split(ligneDepartment, ",")(4) & "," & Split(ligneDepartment, ",")(5) & "," & Split(ligneDepartment, ",")(6) & ","
                If n = -1 Then
                    n += 1
                    ReDim Preserve tableauDepartment(n)
                    tableauDepartment(n) = ligneDepartment
                ElseIf Array.IndexOf(tableauDepartment, ligneDepartment) = -1 Then
                    n += 1
                    ReDim Preserve tableauDepartment(n)
                    tableauDepartment(n) = ligneDepartment
                End If
            End If
        End While

        FileClose(1)



        If tableauDest Is Nothing Then Throw New Exception("tableauDest est vide")



        'MODIFICATION ET AJOUT DES NOUVELLES DESTINATIONS
        Dim OUDomDest As DirectoryEntry = New DirectoryEntry("LDAP://OU=Equipes,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim dirDestSearcher As DirectorySearcher = New DirectorySearcher(OUDomDest)

        For l = 0 To UBound(tableauDest)
            Dim temp() As String = Split(tableauDest(l), ",")
            Dim SAMAccountGroup As String = temp(0) & " grp"
            Dim description As String = temp(1)
            Dim IdDest As String = temp(2)
            Dim samManager As String = temp(3)
            Dim department_nomCourt As String = temp(4)
            Dim department_nom As String = temp(5)

            Try
                If tableauDest(l) <> "" Then
                    'L'attribut "languageCode" contient le groupe ID de GDPI
                    dirDestSearcher.Filter = "(&(objectClass=group) (languageCode=" & IdDest & "))"
                    'dirDestSearcher.Filter = "(&(objectClass=group) (sAMAccountName=" & SAMAccountGroup & "))"
                    dirDestSearcher.SearchScope = SearchScope.OneLevel
                    Dim result As SearchResult = dirDestSearcher.FindOne()
                    Dim dirDest As New DirectoryEntry
                    If result Is Nothing Then
                        'Creation d'une nouvelle destination
                        Commun.NouveauGroupe(OUDomDest, SAMAccountGroup, description)
                        dirDest = New DirectoryEntry("LDAP://" & Commun.TransformeSAMACCOUNTenCN(SAMAccountGroup))
                        'L'attribut "languageCode" contient le groupe ID de GDPI
                        dirDest.Properties("languageCode").Value = IdDest
                        Commun.AppliquerChangement(dirDest)
                    Else
                        dirDest = New DirectoryEntry(result.Path)
                    End If


                    If samManager <> "?" Then
                        'Modification des destinations 
                        Dim CNchefEquipe As String = TransformeSAMACCOUNTenCN(samManager)

                        If dirDest.Properties("SAMAccountName").Value <> SAMAccountGroup Then
                            dirDest.Properties("SAMAccountName").Value = SAMAccountGroup
                            Commun.AppliquerChangement(dirDest)
                            dirDest.Rename("CN=" & SAMAccountGroup)
                            Commun.AppliquerChangement(dirDest)
                        End If

                        If dirDest.Properties("managedBy").Value <> CNchefEquipe Then
                            dirDest.Properties("managedBy").Value = CNchefEquipe
                            Commun.AppliquerChangement(dirDest)
                            End
                            If dirDest.Properties("description").Value <> description Then
                                dirDest.Properties("description").Value = description
                                Commun.AppliquerChangement(dirDest)
                            End If
                            If dirDest.Properties("department").Value <> department_nom Then
                                dirDest.Properties("department").Value = department_nom
                                Commun.AppliquerChangement(dirDest)
                            End If

                        End If

                        dirDest.Close()
                        dirDest.Dispose()
                        dirDest = Nothing
                        result = Nothing
                    End If
                End If
            Catch ex As Exception
                Commun.Journal("ERREUR : Gestion des Destinations : Ajout/Modification " & SAMAccountGroup & " : " & ex.Message, True)
            End Try

            ''gestion des departements
            'Dim OUDomDepartment As DirectoryEntry = New DirectoryEntry("LDAP://OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr")
            'If TransformeSAMACCOUNTenCN("Dpt_" & department_nomCourt) = "" Then
            '    Commun.NouveauGroupe(OUDomDepartment, "Dpt_" & department_nomCourt, "Departement " & department_nom)
            'End If
            'OUDomDepartment.Close()
            'OUDomDepartment.Dispose()
            'OUDomDepartment = Nothing

        Next l


        'gestion des departements
        Dim OUDomDepart As DirectoryEntry = New DirectoryEntry("LDAP://OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim dirDepartSearcher As DirectorySearcher = New DirectorySearcher(OUDomDepart)
        For p = 0 To UBound(tableauDepartment)
            Dim temp() As String = Split(tableauDepartment(p), ",")
            Dim SAMAccountDepart As String = "Dpt_" & temp(0)
            Dim department_nom As String = temp(1)
            Dim IdDepart As String = temp(2)
            Try
                If tableauDepartment(p) <> "" Then
                    'L'attribut "languageCode" contient le departement ID de GDPI
                    dirDepartSearcher.Filter = "(&(objectClass=group) (languageCode=" & IdDepart & "))"
                    'dirDestSearcher.Filter = "(&(objectClass=group) (sAMAccountName=" & SAMAccountGroup & "))"
                    dirDepartSearcher.SearchScope = SearchScope.OneLevel
                    Dim result As SearchResult = dirDepartSearcher.FindOne()
                    Dim dirDepart As New DirectoryEntry
                    If result Is Nothing Then
                        'Creation d'un nouveau departement
                        Try
                            Commun.NouveauGroupe(OUDomDest, SAMAccountDepart, "Departement " & department_nom)
                        Catch

                        End Try
                        dirDepart = New DirectoryEntry("LDAP://" & Commun.TransformeSAMACCOUNTenCN(SAMAccountDepart))
                        'L'attribut "languageCode" contient le groupe ID de GDPI
                        dirDepart.Properties("languageCode").Value = IdDepart
                        Commun.AppliquerChangement(dirDepart)
                    Else
                        dirDepart = New DirectoryEntry(result.Path)
                    End If

                    'Modification des departements 
                    If dirDepart.Properties("SAMAccountName").Value <> SAMAccountDepart Then
                        dirDepart.Properties("SAMAccountName").Value = SAMAccountDepart
                        Commun.AppliquerChangement(dirDepart)
                        dirDepart.Rename("CN=" & SAMAccountDepart)
                        Commun.AppliquerChangement(dirDepart)
                    End If

                    If dirDepart.Properties("description").Value <> "Departement " & department_nom Then
                        dirDepart.Properties("description").Value = "Departement " & department_nom
                        Commun.AppliquerChangement(dirDepart)
                    End If

                    dirDepart.Close()
                    dirDepart.Dispose()
                    dirDepart = Nothing
                    result = Nothing
                End If

            Catch ex As Exception
                Commun.Journal("ERREUR : Gestion des Departements: Ajout/Modification " & SAMAccountDepart & " : " & ex.Message, True)
            End Try


        Next p

        '----------------------------------------------------------------------------------------------------------------------------------
        'CONTROLE DES DESTINATIONS
        dirDestSearcher.Filter = "(&(objectClass=group))"
        dirDestSearcher.SearchScope = SearchScope.OneLevel
        dirDestSearcher.PropertiesToLoad.Add("SAMAccountName")
        dirDestSearcher.PropertiesToLoad.Add("description")
        Dim result1 As SearchResultCollection = dirDestSearcher.FindAll

        For Each grpResult As SearchResult In result1
            Dim SAMAccountGroup As String = grpResult.Properties("SAMAccountName")(0)
            Try
                Dim descriptionGroup As String = ""
                If grpResult.Properties.Contains("description") Then
                    descriptionGroup = grpResult.Properties("description")(0)
                End If
                'Controle si l'equipe Existe dans la BDP
                Dim ctrlGrpExistBDP As Boolean = False
                For l = 0 To UBound(tableauDest)
                    If SAMAccountGroup = Split(tableauDest(l), ",")(0) & " grp" Or SAMAccountGroup = "EXTERNE grp" Then
                        ctrlGrpExistBDP = True
                        Exit For
                    End If
                Next l

                'Si l'equipe n'existe pas dans la BDP
                If ctrlGrpExistBDP = False Then
                    Dim nbrDeMembreDuGroupe As Integer = grpResult.GetDirectoryEntry.Properties("Member").Count
                    Dim groupNonModifieDepuis As Integer = DateDiff(DateInterval.Month, grpResult.GetDirectoryEntry.Properties("whenChanged").Value, Now)
                    'Si le groupe n'a pas été modifié depuis plus de 3 mois
                    If groupNonModifieDepuis > 2 Then
                        'Si le groupe est vide et
                        If nbrDeMembreDuGroupe = 0 Then
                            OUDomDest.Children.Remove(grpResult.GetDirectoryEntry)
                            Commun.AppliquerChangement(OUDomDest)
                            Commun.Journal("Equipe supprimée : " & descriptionGroup & " (" & SAMAccountGroup & ")")
                            'Pour ne pas passer dans la fonction de controle d'association destination<->equipeinfo
                            Continue For

                        Else
                            Throw New Exception("Nettoyage : Le groupe n'est pas vide")
                        End If
                    End If
                End If

                'Controle si la destination est associé a une Equipeinfo 
                If grpResult.GetDirectoryEntry.Properties("Member").Count <> 0 Then
                    If Commun.RecupEquipeinfo(Replace(SAMAccountGroup, " grp", "")) = "" Then

                        Commun.SendEmail("administrateur@igbmc.fr", RecupDataini.RecupVar("[MailEquipeinfo]"), "Destination non associée a une Equipe Info", "La destination " & descriptionGroup & " (" & SAMAccountGroup & "), dans l'active Directory, n'est actuellement pas associé à une Equipe info." & vbCrLf & "Cette destination contient un ou des utilisateurs." & vbCrLf & "Tant que cette association ne sera pas faite, les utilisateurs n'auront pas d'acces a leur Zone Labo")
                    End If
                End If
            Catch ex As Exception
                Commun.Journal("ERREUR : Gestion des Destinations : Controle : " & SAMAccountGroup & " : " & ex.Message, True)
            End Try
        Next grpResult
        dirDestSearcher.Dispose()
        result1.Dispose()
        result1 = Nothing
        'dirSearcher = Nothing
        '----------------------------------------------------------------------------------------------------------------------------------

        OUDomDest.Close()
        OUDomDest.Dispose()
        OUDomDest = Nothing
        Erase tableauDest
        Commun.Journal("Gestion des destinations et départements réussie", False)

    End Sub

    Shared Sub ReactiveDesactiveCompte()

        Dim OUSource As DirectoryEntry
        Dim ouDestDesactive As DirectoryEntry = New DirectoryEntry("LDAP://OU=Comptes Désactivés,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim ouDestException As DirectoryEntry = New DirectoryEntry("LDAP://OU=Exceptions,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim ouDestUser As DirectoryEntry = New DirectoryEntry("LDAP://OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")

        For cas = 0 To 2
            If cas = 0 Then
                OUSource = ouDestDesactive
            ElseIf cas = 1 Then
                OUSource = ouDestException
            Else
                OUSource = ouDestUser
            End If

            Dim searcher As DirectorySearcher = New DirectorySearcher(OUSource)
            searcher.PageSize = 2000
            searcher.Filter = "(&(objectClass=user))"
            ' pour la recherche non recursive(SearchScope.OneLevel)
            searcher.SearchScope = SearchScope.OneLevel
            Dim users As SearchResultCollection = searcher.FindAll()
            searcher.Dispose()
            searcher = Nothing
            For Each user As SearchResult In users
                Dim DirEntry = user.GetDirectoryEntry
                Dim login As String = DirEntry.Properties("SAMAccountName").Value()
                Dim except As String = exceptionUser(login)
                Dim ctrlPresentBDP As Boolean = PresentBDP(login)



                'Cas d'un compte Actif par BDP ou par Exception
                If ctrlPresentBDP = True Or except <> "False" Then
                    Commun.GestionGroupeUserActive(DirEntry)
                    'Si l'utilistateur est present dans la base du perso mais n'est pas dans l'OU Utilisateur
                    If ctrlPresentBDP = True And cas <> 2 Then
                        Try
                            Commun.SetADLDAPProperty(DirEntry, "description", "")
                            Commun.SetADLDAPProperty(DirEntry, "Comment", "Réactivé le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf, True)
                            Commun.SetADLDAPProperty(DirEntry, "accountDeletionDate", "")
                            'DirEntry.Properties("accountDeletionDate").Clear()
                            Commun.AppliquerChangement(DirEntry)
                        Catch ex As Exception
                            Commun.Journal("ERREUR : réactivation du compte(propriétés) : " & DirEntry.Name & " : " & ex.Message, True)
                        End Try

                        Try

                        Catch ex As Exception
                            Commun.Journal("ERREUR : date d'expiration du compte (deplacement Vers OU Utilisateurs) : " & DirEntry.Name, True)
                        End Try

                        'Réactivation du compte et deplacement dans l'OU Utilisateurs

                        'Active l'utilisateur ADM s'il existe
                        Commun.ReactiveDesactiveCompteAdm(login & "adm", "Active")
                        Try
                            CompararaisonAjoutRetraitDestinations(DirEntry)
                            DirEntry.MoveTo(ouDestUser)
                        Catch ex As Exception
                            Commun.Journal("ERREUR : réactivation du compte (deplacement Vers OU Utilisateurs) : " & DirEntry.Name, True)
                        End Try
                    End If

                    'Si l'utilisateur n'est pas present dans la base du perso mais est en exception
                    If ctrlPresentBDP = False And except <> "False" Then
                        Try
                            If DirEntry.Properties("description").Value <> "EXCEPTION jusqu'au: " & except Then
                                Commun.SetADLDAPProperty(DirEntry, "description", "EXCEPTION jusqu'au: " & except)
                                Commun.SetADLDAPProperty(DirEntry, "Comment", "Exception le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf, True)
                                Commun.SetADLDAPProperty(DirEntry, "accountDeletionDate", "")
                                Commun.AppliquerChangement(DirEntry)
                            End If
                        Catch ex As Exception
                            Commun.Journal("ERREUR : réactivation du compte(propriétés) : " & DirEntry.Name, True)
                        End Try

                        'si il n'est pas deja dans l'OU Exception
                        If cas <> 1 Then

                            'Active l'utilisateur ADM s'il existe
                            Commun.ReactiveDesactiveCompteAdm(login & "adm", "Active")

                            Commun.GestionGroupeUserActive(DirEntry)

                            Try
                                'Le deplacer dans l'OU Exception
                                DirEntry.MoveTo(ouDestException)

                            Catch ex As Exception
                                Commun.Journal("ERREUR : réactivation du compte (deplacement vers OU Exceptions) : " & DirEntry.Name, True)
                            End Try
                        End If
                    End If
                End If

                'Si l'utilisateur n'est pas present dans la base du perso ni dans les exceptions et, est dans l'OU Utilisateur
                If ctrlPresentBDP = False And except = "False" And cas <> 0 Then
                    Try
                        'si il est dans l'OU Exception
                        If cas = 1 Then
                            Commun.SetADLDAPProperty(DirEntry, "description", "EXCEPTION Désactivé le: " & Strings.Left(CStr(Now), 10))
                            Commun.SetADLDAPProperty(DirEntry, "accountDeletionDate", Strings.Left(DateTime.UtcNow.AddMonths(3).ToString, 10))
                            Commun.AppliquerChangement(DirEntry)
                        End If

                        'si il est dans l'OU Utilisateurs
                        If cas = 2 Then
                            Commun.SetADLDAPProperty(DirEntry, "description", "Désactivé le: " & Strings.Left(CStr(Now), 10))
                            Commun.SetADLDAPProperty(DirEntry, "accountDeletionDate", Strings.Left(DateTime.UtcNow.AddMonths(3).ToString, 10))
                            Commun.AppliquerChangement(DirEntry)
                        End If
                        Commun.SetADLDAPProperty(DirEntry, "Comment", "Désactivé le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf, True)
                        Commun.AppliquerChangement(DirEntry)

                    Catch ex As Exception
                        Commun.Journal("ERREUR : désactivation du compte(propriétés) : " & DirEntry.Name, True)
                    End Try

                    'Desactive l'utilisateur ADM s'il existe
                    Commun.ReactiveDesactiveCompteAdm(login & "adm", "Desactive")

                    'Le deplacer dans l'ou Comptes Désactivés
                    Commun.GestionGroupeUserDesactive(DirEntry)
                    CompararaisonAjoutRetraitDestinations(DirEntry)

                    Dim adresseMail As String = DirEntry.Properties("Mail").Value
                    Dim prenom As String = DirEntry.Properties("givenName").Value
                    Dim dateDeSuppressionPrevue As String = DirEntry.Properties("accountDeletionDate").Value
                    Dim mail = MailCloture(prenom, dateDeSuppressionPrevue)
                    Try

                        Commun.SendEmail("noreply@igbmc.fr", adresseMail & ";Cc:serviceinfo@igbmc.fr", "ARRET DU COMPTE", mail)
                        Commun.Journal("Mail de fermeture de compte envoyé (m-3) : " & adresseMail)

                    Catch ex As Exception
                        Commun.Journal("ERREUR : Mail de fermeture de compte envoyé (m-3) : " & adresseMail & " : " & ex.Message, True)
                    End Try

                    Try
                        DirEntry.MoveTo(ouDestDesactive)
                    Catch ex As Exception
                        Commun.Journal("ERREUR : désactivation du compte (deplacement vers OU Désactivés) : " & DirEntry.Name & " : " & ex.Message, True)
                    End Try
                End If

                'Mise de la valeur a "False" apres le traitement pour eviter une boucle dans for
                except = "False"


                DirEntry.Close()
                DirEntry.Dispose()
                DirEntry = Nothing
            Next user
            users.Dispose()
            users = Nothing

            OUSource.Close()

        Next cas

        OUSource.Dispose()
        OUSource = Nothing

        ouDestUser.Close()
        ouDestUser.Dispose()
        ouDestUser = Nothing

        ouDestException.Close()
        ouDestException.Dispose()
        ouDestException = Nothing

        ouDestDesactive.Close()
        ouDestDesactive.Dispose()
        ouDestDesactive = Nothing
        Commun.Journal("Gestion de l'activation/desactivation des comptes réussie", False)
    End Sub
    Shared Function MailCloture(ByVal prenom As String, ByVal dateDeSuppressionPrevue As String)
        Return "Bonjour " & prenom & ", " & vbCrLf & vbCrLf & "Votre compte (mail, acces informatique, zone labo, profil,...) à l'IGBMC va etre definitivement supprimé le " & dateDeSuppressionPrevue & "." & vbCrLf & "Les données de votre zone Labo ont été mises à la disposition de votre chef d'equipe. Pour les récupérer, veuillez le contacter." & vbCrLf & "Veuillez récuperer les éléments de votre boite mail que vous souhaitez conserver, avant cette date." & vbCrLf & "Apres cette date, plus aucune récuperation, ne sera possible, elles seront definitivement detruites." & vbCrLf & vbCrLf & "En cas de besoin d'assistance, vous pouvez envoyer un mail à helpdesk@igbmc.fr" & vbCrLf & vbCrLf & "Le service Informatique."
    End Function

    Shared Function TransformeSAMACCOUNTenCN(ByVal login As String) As String
        Dim Ldap As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")
        Dim searcher As DirectorySearcher = New DirectorySearcher(Ldap)
        searcher.Filter = "(& (SAMAccountName=" & login & "))"
        Dim result As SearchResult = searcher.FindOne()
        If result Is Nothing Then
            Return ""
        Else
            Dim pathCN As String = Replace(result.Path, "LDAP://", "")
            Return pathCN
        End If
        searcher.Dispose()
        searcher = Nothing
        Ldap.Close()
        Ldap.Dispose()
        Ldap = Nothing
    End Function
    Shared Function UserMembreDeDestination(ByVal username As String) As String()
        Dim appartientA As String()
        Dim j As Integer = 0
        Dim Entry As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")
        Dim Searcher As New DirectorySearcher(Entry)
        Searcher.SearchScope = DirectoryServices.SearchScope.Subtree
        Searcher.Filter = "(&(objectcategory=user)(SAMAccountName=" & username & "))"
        Dim res As SearchResult = Searcher.FindOne
        For i = 0 To res.Properties("memberOf").Count() - 1
            Dim temp As String = Replace(res.Properties("memberOf")(i).ToString, ",OU=Equipes,DC=igbmc,DC=u-strasbg,DC=fr", "")
            temp = Replace(temp, "CN=", "")
            If Strings.Right(temp, 4) = " grp" Then
                ReDim Preserve appartientA(j)
                appartientA(j) = temp
                j += 1
            End If
        Next i
        Searcher.Dispose()
        Searcher = Nothing
        Return appartientA
        Entry.Close()
        Entry.Dispose()
        Entry = Nothing
        appartientA = Nothing
    End Function

    Shared Function EquipeComptableFichierUser(ByVal login As String) As String()
        Try
            Dim srJ As StreamReader
            If withJson = "ftp" Then
                srJ = New StreamReader("c:\temp\listep1.txt", System.Text.Encoding.Default)
            Else
                srJ = New StreamReader("c:\temp\listepersoJson.txt", System.Text.Encoding.Default)
            End If
            Dim ligneFichier As String
            Dim equipeUser() As String
            Dim j As Integer = 0
            Do
                ligneFichier = srJ.ReadLine

                If ligneFichier Is Nothing Then
                    Exit Do
                Else
                    If InStr(ligneFichier, ",") <> 0 Then
                        Dim temp As String() = Split(ligneFichier, ",")
                        If temp(2) = login Then
                            ReDim Preserve equipeUser(j)
                            equipeUser(j) = temp(3) & " grp"
                            j += 1
                        End If
                        temp = Nothing
                    End If
                End If
            Loop
            srJ.Close()
            srJ.Dispose()
            Return equipeUser
        Catch ex As Exception
            Commun.Journal("ERREUR : Recuperation des Destinations : " & login & " : " & ex.Message, True)
        End Try

    End Function
    Public Shared Function chercheDepartementparRapporDestination(ByVal destination As String) As String
        Try
            Dim srJ As StreamReader
            srJ = New StreamReader("c:\temp\eq.txt", System.Text.Encoding.Default)

            destination = Replace(destination, " grp", "")

            Dim ligneFichier As String
            Dim departement As String = ""
            Dim j As Integer = 0
            Do
                ligneFichier = srJ.ReadLine

                If ligneFichier Is Nothing Then
                    Exit Do
                Else
                    If InStr(ligneFichier, ",") <> 0 Then
                        Dim temp As String() = Split(ligneFichier, ",")
                        If temp(0) = destination Then
                            departement = "Dpt_" & temp(4)
                            Exit Do
                        End If
                        temp = Nothing
                    End If
                End If
            Loop

            srJ.Close()
            srJ.Dispose()
            Return departement
        Catch ex As Exception
            Commun.Journal("ERREUR : Recuperation du departement : " & destination & " : " & ex.Message, True)
        End Try
    End Function
    Public Shared Sub CompararaisonAjoutRetraitDestinations(ByVal usrDir As DirectoryEntry)
        Dim login As String = usrDir.Properties("SAMAccountName").Value
        Dim EquipeUserAD As String() = UserMembreDeDestination(login)
        Dim EquipeUserFichier As String() = EquipeComptableFichierUser(login)


        If Not EquipeUserAD Is Nothing And Not EquipeUserFichier Is Nothing Then
            For i = 0 To UBound(EquipeUserAD)
                'Dim destGrp As String = chercheDepartementparRapporDestination(EquipeUserFichier(i))
                'Commun.AddRemoveADGroup(login, destGrp, "Add")
                For j = 0 To UBound(EquipeUserFichier)
                    If EquipeUserFichier(j) <> "" Then
                        If EquipeUserAD(i) = EquipeUserFichier(j) Then
                            EquipeUserAD(i) = ""
                            EquipeUserFichier(j) = ""
                        End If
                    End If
                Next j
            Next i
        End If

        If Not EquipeUserFichier Is Nothing Then
            For j = 0 To UBound(EquipeUserFichier)
                If EquipeUserFichier(j) <> "" Then
                    Dim dptGrp As String = chercheDepartementparRapporDestination(EquipeUserFichier(j))
                    If Commun.AppartientGroup(login, "G_Domain_DisableOpenSession") = False And Commun.AppartientGroup(login, EquipeUserFichier(j)) = False Then
                        Commun.AddRemoveADGroup(login, (EquipeUserFichier(j)), "Add")
                        Commun.AddRemoveADGroup(login, dptGrp, "Add")
                        Commun.Journal("Succes : CompararaisonAjoutRetraitEquipesComptables : " & login & " ajouté dans " & EquipeUserFichier(j))
                    End If
                End If
            Next j
        End If

        If Not EquipeUserAD Is Nothing Then
            For i = 0 To UBound(EquipeUserAD)

                If EquipeUserAD(i) <> "" Then
                    Dim dptGrp As String = chercheDepartementparRapporDestination(EquipeUserAD(i))
                    If Commun.AppartientGroup(login, EquipeUserAD(i)) = True Then
                        Commun.AddRemoveADGroup(login, EquipeUserAD(i), "Remove")
                        Commun.AddRemoveADGroup(login, dptGrp, "Remove")
                        Commun.Journal("Succes : CompararaisonAjoutRetraitEquipesComptables : " & login & " retiré de " & EquipeUserAD(i))
                    End If
                End If
            Next i
        End If
        EquipeUserAD = Nothing
        EquipeUserFichier = Nothing

    End Sub

    Shared Function DonneesUnix(ByVal login As String, ByVal equipe As String) As String
        Try
            Dim LdapU As DirectoryEntry = New DirectoryEntry("LDAP://130.79.78.178:1389/ou=People,dc=igbmc,dc=fr", "", "", AuthenticationTypes.Anonymous)
            Dim ResultFields() As String = {"uidnumber", "gidnumber", "homedirectory", "loginShell"}
            Dim searcherLdapU As DirectorySearcher = New DirectorySearcher(LdapU)
            searcherLdapU.Filter = "uid=" & login
            searcherLdapU.PropertiesToLoad.AddRange(ResultFields)
            Dim resultLdapU As SearchResult = searcherLdapU.FindOne()

            Return resultLdapU.Properties("uidnumber")(0) & "," & resultLdapU.Properties("gidnumber")(0) & "," & resultLdapU.Properties("homedirectory")(0) & "," & resultLdapU.Properties("loginShell")(0)
            searcherLdapU.Dispose()
            searcherLdapU = Nothing
            LdapU.Close()
            LdapU.Dispose()
            LdapU = Nothing
        Catch ex As Exception
            Commun.Journal("ERREUR : DonneesUnix : " & login & " - " & equipe & " : " & ex.Message, True)
        End Try
    End Function

    Shared Sub modifAliasMail(ByVal objuser As DirectoryEntry, ByVal prenom As String, ByVal nom As String)

        Dim ctrlChangement As Boolean = False
        Dim aliasMail As String = prenom & "." & nom
        aliasMail = LCase(Replace(aliasMail, " ", "-"))
        aliasMail = LCase(Replace(aliasMail, "'", ""))

        Dim userID As String = objuser.Properties("EmployeeID").Value

        Dim newAliasMail As String = Commun.DetermineAliasLibre(prenom, nom, userID)

        Try

            Dim currentAddresses = objuser.Properties("proxyAddresses").Value
            Dim adress As New List(Of String)


            For Each value In currentAddresses
                Dim tempValue = Strings.Left(value, 5) & LCase(Strings.Right(value, Len(value) - 5))
                tempValue = tempValue
                adress.Add(tempValue)
            Next value
            'Definit tous les format d'adresse possible avec le nouveau nom
            Dim nvlAdresse1 As String = "smtp:" & aliasMail & "@igbmc.fr"
            Dim nvlAdresse2 As String = "SMTP:" & aliasMail & "@igbmc.fr"
            Dim nvlAdresse3 As String = "smtp:" & aliasMail & "@igbmc.u-strasbg.fr"
            Dim nvlAdresse4 As String = "SMTP:" & aliasMail & "@igbmc.u-strasbg.fr"

            Dim iof1 As Integer = adress.IndexOf(nvlAdresse1)
            Dim iof2 As Integer = adress.IndexOf(nvlAdresse2)
            Dim iof3 As Integer = adress.IndexOf(nvlAdresse3)
            Dim iof4 As Integer = adress.IndexOf(nvlAdresse4)

            If iof1 = -1 And iof2 = -1 Then
                adress.Add(nvlAdresse1)
                objuser.Properties("proxyAddresses").Value = adress.ToArray()
                Commun.AppliquerChangement(objuser)
                ctrlChangement = True
            End If
            If iof3 = -1 And iof4 = -1 Then
                adress.Add(nvlAdresse3)
                objuser.Properties("proxyAddresses").Value = adress.ToArray()
                Commun.AppliquerChangement(objuser)
                ctrlChangement = True
            End If
            currentAddresses = Nothing
            adress.Clear()
            adress = Nothing

            'Si il y a eu des changements, enregistrement de l'alias mail dans l'historique
            If ctrlChangement = True Then
                Commun.ajoutAliasFichierHisto(aliasMail, userID)
                objuser.Properties("msExchExtensionAttribute16").Value = aliasMail & "@igbmc.fr"
                objuser.CommitChanges()
            End If

        Catch ex As Exception
            Commun.Journal("ERREUR : Modification alias Mail sur AD : " & objuser.Properties("SAMAccountName").Value & " : " & ex.Message, True)
        End Try

    End Sub

    Shared Sub CreationFichierParFTP()
        Commun.getFileFTP("/" & RecupDataini.RecupVar("[FCCFilePath]") & "/listepersonew.txt", "c:\temp\listepers.txt")


        'SUPPRESSION DES ACCENTS DANS LE FICHIER DU PERSONNEL
        Dim Tc As String()
        FileOpen(2, "c:\temp\listep1.txt", OpenMode.Output) 'fichier du personnel sans accent multi equipe
        FileOpen(3, "c:\temp\listepers.txt", OpenMode.Input)

        Dim h As Integer = -1

        While Not EOF(3)
            h = h + 1
            ReDim Preserve Tc(h)
            Tc(h) = LineInput(3)
            Tc(h) = Replace(Tc(h), "â", "a")
            Tc(h) = Replace(Tc(h), "à", "a")
            Tc(h) = Replace(Tc(h), "é", "e")
            Tc(h) = Replace(Tc(h), "É", "E")
            Tc(h) = Replace(Tc(h), "È", "E")
            Tc(h) = Replace(Tc(h), "è", "e")
            Tc(h) = Replace(Tc(h), "ê", "e")
            Tc(h) = Replace(Tc(h), "ë", "e")
            Tc(h) = Replace(Tc(h), "ì", "i")
            Tc(h) = Replace(Tc(h), "î", "i")
            Tc(h) = Replace(Tc(h), "ï", "i")
            Tc(h) = Replace(Tc(h), "ü", "u")
            Tc(h) = Replace(Tc(h), "ù", "u")
            Tc(h) = Replace(Tc(h), "ö", "o")
            Tc(h) = Replace(Tc(h), "ô", "o")
            Tc(h) = Replace(Tc(h), "ç", "c")
            PrintLine(2, Tc(h))

        End While
        FileClose(2)
        FileClose(3)
        Tc = Nothing
    End Sub

    Shared Sub ExpirationMDP()
        Dim samaccoutname As String
        Try

            Dim OUAdmins As DirectoryEntry = New DirectoryEntry("LDAP://OU=AdmInfo,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr")
            Dim OUAdminsearcher As DirectorySearcher = New DirectorySearcher(OUAdmins)
            OUAdminsearcher.Filter = "(&(objectClass=user))"
            OUAdminsearcher.PropertiesToLoad.Add("pwdLastSet")
            OUAdminsearcher.PropertiesToLoad.Add("SamAccountName")
            OUAdminsearcher.PropertiesToLoad.Add("userAccountControl")
            For Each resultUser As SearchResult In OUAdminsearcher.FindAll()
                samaccoutname = resultUser.Properties("SamAccountName")(0)
                'si le compte est userprog continuer sans traiter
                If samaccoutname = "userprog" Then Continue For

                Dim rappelComplexite As String = ""
                Dim StrategieMDP As DirectoryEntry
                If Commun.AppartientGroup(samaccoutname, "G_SMDPM_Admins") = True Then
                    StrategieMDP = New DirectoryEntry("LDAP://CN=Strategie_MDPM_Admins,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr")
                ElseIf Commun.AppartientGroup(samaccoutname, "G_SMDPM_Users-Admins") = True Then
                    StrategieMDP = New DirectoryEntry("LDAP://CN=Strategie_MDPM_Users-Admins,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr")
                Else
                    Commun.Journal("ERREUR : Determiner strategie de mot de passe : " & samaccoutname, True)
                    GoTo sortie
                End If

                Dim nbrChar As String = StrategieMDP.Properties("msDS-MinimumPasswordLength").Value.ToString
                Dim historyLenght As String = StrategieMDP.Properties("msDS-PasswordHistoryLength").Value.ToString
                Dim aaaa = StrategieMDP.Properties("msDS-MaximumPasswordAge").Value
                Dim bbb = aaaa / 60 / 60 / -24 / 10 ^ 7
                Dim MaxPWDageJourPolicy As Integer = ConvertAttribute(StrategieMDP.Properties("msDS-MaximumPasswordAge").Value)

                StrategieMDP.Close()
                StrategieMDP.Dispose()
                StrategieMDP = Nothing

                rappelComplexite = "Votre mot de passe doit etre changé tous les " & MaxPWDageJourPolicy & " jours." & vbCrLf & "Il ne peut pas etre le meme que les " & historyLenght & " précédents." & vbCrLf & "Votre mot de passe doit contenir au moins :" & vbCrLf & vbTab & "- " & nbrChar & " caractères"
                rappelComplexite = rappelComplexite & vbCrLf & vbTab & "- 1 Majuscule" & vbCrLf & vbTab & "- 1 Minuscule" & vbCrLf & vbTab & "- 1 Chiffre" & vbCrLf & vbTab & "- 1 Caractère spécial (non-alphabétique)"

                Dim lastSetPWD As DateTime = Format(New DateTime(1601, 1, 2).AddTicks(resultUser.Properties("pwdLastSet")(0)), "dd/MM/yyyy")
                Dim expirationPWD As DateTime = lastSetPWD.AddDays(MaxPWDageJourPolicy - 1)

                If resultUser.Properties("userAccountControl")(0) <> 512 And samaccoutname <> "Modele_Admin_Service" Then
                    Dim userAD As DirectoryEntry = New DirectoryEntry(resultUser.Path)
                    userAD.Properties("userAccountControl").Value = 512
                    Commun.AppliquerChangement(userAD)
                    userAD.Close()
                    userAD.Dispose()
                    userAD = Nothing
                End If

                Dim demain As DateTime = Format(Date.Now.AddDays(1), "dd/MM/yyyy")
                If demain = expirationPWD Then

                    Dim Email As String = Strings.Left(samaccoutname, Len(samaccoutname) - 3) & "@igbmc.fr"
                    Dim corpMail As String = "Le mot de passe de votre compte d'administration (" & samaccoutname & ") va expirer demain (" & demain & ")." & vbCrLf & "Pensez à le changer en ouvrant une session sur un ordinateur du domaine ou avant qu'il ne soit expiré en vous connectant ici : http://password.igbmc.u-strasbg.fr/default.aspx" & vbCrLf & vbCrLf & rappelComplexite & vbCrLf & vbCrLf & vbCrLf & "Le service Informatique" & vbCrLf & "(Email généré automatiquement)"
                    Commun.SendEmail("serviceinfo@igbmc.fr", Email, "Expiration de Votre mot de passe de compte administrateur", corpMail)
                End If
            Next resultUser

sortie:
            OUAdminsearcher.Dispose()
            OUAdminsearcher = Nothing
            OUAdmins.Close()
            OUAdmins.Dispose()
            OUAdmins = Nothing
            Commun.Journal("Gestion des mots de passe des comptes ADM terminée avec succes", False)
        Catch ex As Exception
            Commun.Journal("ERREUR : Gestion de password des comptes ADM : " & samaccoutname & " : " & ex.Message, True)
        End Try

    End Sub
    Shared Function ConvertAttribute(ByVal li As Object) As Integer


        Dim int64Val As IADsLargeInteger = CType(li, IADsLargeInteger)

        If Not int64Val Is Nothing Then
            Dim largeInt As System.Int64 = int64Val.HighPart
            largeInt = largeInt << 32
            largeInt += int64Val.LowPart
            Return (Now.AddTicks(-largeInt) - Now).Days
        End If


    End Function
    Shared Function exceptionUser(ByVal login As String) As String
        'controle si l'utilisateur fait partie des exceptions (toujours valide) en retournant la date limite de l'exception
        'sinon retourne False
        Dim resultat As String = "False"
        Try
            If tabExcepUser Is Nothing Then
                Dim monStreamReader As StreamReader = New StreamReader("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\UsersExceptions.txt")
                Dim ligne As String
                Dim i As Integer = -1
                Do
                    ligne = monStreamReader.ReadLine()
                    If Not ligne Is Nothing Then
                        Dim tabLigneExceptUser As String() = Split(ligne, ",")
                        Dim dateExpire As DateTime = Convert.ToDateTime(tabLigneExceptUser(1))
                        If dateExpire >= Now Then
                            i += 1
                            ReDim Preserve tabExcepUser(1, i)
                            tabExcepUser(0, i) = tabLigneExceptUser(0)
                            tabExcepUser(1, i) = tabLigneExceptUser(1)
                        End If
                        tabLigneExceptUser = Nothing
                    End If
                Loop Until ligne Is Nothing
                monStreamReader.Close()
                'monStreamReader.Dispose()
            End If

            If Not tabExcepUser Is Nothing Then
                Dim positionTab As Integer = IndexOfMulti(tabExcepUser, login, 0)
                If positionTab = -1 Then
                    resultat = "False"
                Else
                    resultat = tabExcepUser(1, positionTab)
                End If
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : exceptionUser : login : " & login & " : " & ex.Message, True)
        End Try

        Return resultat

    End Function
    Shared Function IndexOfMulti(ByVal tab As String(,), ByVal recherche As String, ByVal colonne As Integer) As Integer
        Dim j As Integer = -1
        If Not tab Is Nothing Then

            Dim tabTemp As String()
            For i = 0 To UBound(tab, 2)
                ReDim Preserve tabTemp(i)
                tabTemp(i) = tab(colonne, i)
            Next i
            j = Array.IndexOf(tabTemp, recherche)
            tabTemp = Nothing
        End If
        Return j
    End Function

    Shared Sub GroupesDynamiques()


        Dim Entry As DirectoryEntry = New DirectoryEntry("LDAP://OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim Searcher As New DirectorySearcher(Entry)
        Searcher.PageSize = 2000
        Searcher.SearchScope = DirectoryServices.SearchScope.Subtree
        Searcher.PropertiesToLoad.Add("physicalDeliveryOfficeName")
        Searcher.PropertiesToLoad.Add("SamAccountName")
        Searcher.PropertiesToLoad.Add("Department")
        Searcher.Filter = "(&(objectcategory=user))"
        Dim result As SearchResultCollection = Searcher.FindAll

        For Each user As SearchResult In result

            If user.Properties.Contains("SamAccountName") Then

                Dim userSamaccount As String = user.Properties("SamAccountName")(0)
                If user.Properties.Contains("physicalDeliveryOfficeName") Then

                    Dim userBureau As String = user.Properties("physicalDeliveryOfficeName")(0)
                    If Strings.Left(userBureau, 3) = "IGB" Then
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel IGBMC", "Add")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel ICS", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel CBI", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel CEBGS", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel ESBS", "Remove")
                    End If

                    If Strings.Left(userBureau, 3) = "ICS" Then
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel IGBMC", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel ICS", "Add")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel CBI", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel CEBGS", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel ESBS", "Remove")
                    End If

                    If Strings.Left(userBureau, 3) = "CBI" Then
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel IGBMC", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel ICS", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel CBI", "Add")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel CEBGS", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel ESBS", "Remove")
                    End If

                    If Strings.Left(userBureau, 3) = "CEB" Then
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel IGBMC", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel ICS", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel CBI", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel CEBGS", "Add")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel ESBS", "Remove")
                    End If

                    If Strings.Left(userBureau, 3) = "ESB" Then
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel IGBMC", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel ICS", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel CBI", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel CEBGS", "Remove")
                        Commun.AddRemoveADGroup(userSamaccount, "Personnel ESBS", "Add")
                    End If
                End If

                'If user.Properties.Contains("Department") Then

                'Dim userDepartment As String = user.Properties("Department")(0)
                'If userDepartment = "INFO-MICROINFORMATIQUE" Then
                If Commun.AppartientGroup(userSamaccount, "INF-MICRO grp") = True Then 'Groupe_MicroInfo.Invoke("Add", New Object() {user.Path.ToString})
                    Commun.AddRemoveADGroup(userSamaccount, "Pole Micro Informatique", "Add")
                    If Commun.TransformeSAMACCOUNTenCN(userSamaccount & "adm") <> "" Then Commun.AddRemoveADGroup(userSamaccount & "adm", "Pole Micro Informatique ADM", "Add")
                Else
                    Commun.AddRemoveADGroup(userSamaccount, "Pole Micro Informatique", "Remove")
                    'If Commun.AppartientGroup(userSamaccount, "Pole Micro Informatique") = True Then Groupe_MicroInfo.Invoke("Remove", New Object() {user.Path.ToString})
                    If Commun.TransformeSAMACCOUNTenCN(userSamaccount & "adm") <> "" Then Commun.AddRemoveADGroup(userSamaccount & "adm", "Pole Micro Informatique ADM", "Remove")
                End If

                'If userDepartment = "INFO-RESEAUX ET SYSTEME" Then
                If Commun.AppartientGroup(userSamaccount, "INF-RES-SYST grp") = True Then
                    Commun.AddRemoveADGroup(userSamaccount, "Pole Reseau Informatique", "Add")
                    If Commun.TransformeSAMACCOUNTenCN(userSamaccount & "adm") <> "" Then Commun.AddRemoveADGroup(userSamaccount & "adm", "Pole Reseau Informatique ADM", "Add")
                Else
                    Commun.AddRemoveADGroup(userSamaccount, "Pole Reseau Informatique", "Remove")
                    'If Commun.AppartientGroup(userSamaccount, "Pole Reseau Informatique") = True Then Groupe_ReseauInfo.Invoke("Remove", New Object() {user.Path.ToString})
                    If Commun.TransformeSAMACCOUNTenCN(userSamaccount & "adm") <> "" Then Commun.AddRemoveADGroup(userSamaccount & "adm", "Pole Reseau Informatique ADM", "Remove")
                End If

                'If userDepartment = "INFO-DEVELOPPEMENT" Then
                If Commun.AppartientGroup(userSamaccount, "INF-DEV grp") = True Then
                    Commun.AddRemoveADGroup(userSamaccount, "Pole Dev Informatique IGBMC", "Add")
                    If Commun.TransformeSAMACCOUNTenCN(userSamaccount & "adm") <> "" Then Commun.AddRemoveADGroup(userSamaccount & "adm", "Pole Dev Informatique IGBMC ADM", "Add")
                Else
                    Commun.AddRemoveADGroup(userSamaccount, "Pole Dev Informatique IGBMC", "Remove")
                    'If Commun.AppartientGroup(userSamaccount, "Pole Dev Informatique IGBMC") = True Then Groupe_DevInfo.Invoke("Remove", New Object() {user.Path.ToString})
                    If Commun.TransformeSAMACCOUNTenCN(userSamaccount & "adm") <> "" Then Commun.AddRemoveADGroup(userSamaccount & "adm", "Pole Dev Informatique IGBMC ADM", "Remove")
                End If
                'End If


            End If
        Next user
        result.Dispose()

        Entry = Nothing
        Searcher = Nothing

        Dim Entry1 As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")
        Dim Searcher1 As New DirectorySearcher(Entry)
        Searcher1.SearchScope = DirectoryServices.SearchScope.Subtree
        Searcher1.PropertiesToLoad.Add("SamAccountName")
        Searcher1.Filter = "(&(objectCategory=computer))"
        Dim result1 As SearchResultCollection = Searcher1.FindAll

        For Each computer In result1
            Dim namePC = computer.Properties("SamAccountName")(0)
            If Strings.InStr(namePC, "-") > 0 Then
                Dim caracteresAvantTiret = Strings.Left(namePC, (Strings.InStr(namePC, "-") - 1))
                If Strings.LCase(Strings.Mid(namePC, (Strings.InStr(namePC, "-") - 1), 1)) = "a" Or Strings.Mid(namePC, (Strings.InStr(namePC, "-") - 1), 1) = "1" Or Strings.Mid(namePC, (Strings.InStr(namePC, "-") - 1), 1) = "2" Or Strings.Mid(namePC, (Strings.InStr(namePC, "-") - 1), 1) = "3" Or Strings.Mid(namePC, (Strings.InStr(namePC, "-") - 1), 1) = "4" Or Strings.LCase(Strings.Left(namePC, 4)) = "serv" Or Strings.LCase(Strings.Left(namePC, 8)) = "imagerie" Then
                    If Strings.LCase(Strings.Mid(namePC, (Strings.InStr(namePC, "-") - 1), 1)) = "a" Then
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs Administratif", "Add")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs IGBMC", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs ICS", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CEBGS", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CBI", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Serveurs", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs nom incorrect", "Remove")
                    End If
                    If Strings.Mid(namePC, (Strings.InStr(namePC, "-") - 1), 1) = "1" Then
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs Administratif", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs IGBMC", "Add")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs ICS", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CEBGS", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CBI", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Serveurs", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs nom incorrect", "Remove")
                    End If
                    If Strings.Mid(namePC, (Strings.InStr(namePC, "-") - 1), 1) = "2" Or Strings.LCase(Strings.Left(namePC, 8)) = "imagerie" Then
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs Administratif", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs IGBMC", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs ICS", "Add")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CEBGS", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CBI", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Serveurs", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs nom incorrect", "Remove")
                    End If
                    If Strings.Mid(namePC, (Strings.InStr(namePC, "-") - 1), 1) = "3" Then
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs Administratif", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs IGBMC", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs ICS", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CEBGS", "Add")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CBI", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Serveurs", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs nom incorrect", "Remove")
                    End If
                    If Strings.Mid(namePC, (Strings.InStr(namePC, "-") - 1), 1) = "4" Then
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs Administratif", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs IGBMC", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs ICS", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CEBGS", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CBI", "Add")
                        Commun.AddRemoveADGroup(namePC, "Serveurs", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs nom incorrect", "Remove")
                    End If
                    If Strings.LCase(Strings.Left(namePC, 4)) = "serv" Then
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs Administratif", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs IGBMC", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs ICS", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CEBGS", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs CBI", "Remove")
                        Commun.AddRemoveADGroup(namePC, "Serveurs", "Add")
                        Commun.AddRemoveADGroup(namePC, "Ordinateurs nom incorrect", "Remove")
                    End If
                Else
                    Commun.AddRemoveADGroup(namePC, "Ordinateurs Administratif", "Remove")
                    Commun.AddRemoveADGroup(namePC, "Ordinateurs IGBMC", "Remove")
                    Commun.AddRemoveADGroup(namePC, "Ordinateurs ICS", "Remove")
                    Commun.AddRemoveADGroup(namePC, "Ordinateurs CEBGS", "Remove")
                    Commun.AddRemoveADGroup(namePC, "Ordinateurs CBI", "Remove")
                    Commun.AddRemoveADGroup(namePC, "Serveurs", "Remove")
                    Commun.AddRemoveADGroup(namePC, "Ordinateurs nom incorrect", "Add")
                End If
            End If
        Next computer

        result1.Dispose()

        Entry1 = Nothing
        Searcher1 = Nothing
        Commun.Journal("Gestion des groupes dynamiques terminée avec succes", False)
    End Sub
    Shared Function PresentBDP(ByVal login As String) As Boolean
        For i = 0 To UBound(tabPersoMonoEquipe, 2)
            If login = tabPersoMonoEquipe(2, i) Then
                Return True
                Exit Function
            End If
        Next i
        Return False
    End Function
    Shared Function ShellSort(ByVal tab1 As String(), _
                        Optional ByVal loBound As Long = -1, _
                        Optional ByVal upBound As Long = -1) As String()

        Dim i As Long, j As Long, h As Long, v As String

        If loBound = -1 Then
            loBound = LBound(tab1)
        End If
        If upBound = -1 Then
            upBound = UBound(tab1)
        End If

        h = loBound
        Do
            h = 3 * h + 1
        Loop Until h > upBound

        Do
            h = h / 3
            For i = h To upBound
                v = tab1(i) : j = i
                Do While tab1(j - h) > v
                    tab1(j) = tab1(j - h) : j = j - h
                    If j <= h Then
                        Exit Do
                    End If
                Loop
                tab1(j) = v
            Next i
        Loop Until h = loBound
        Return tab1
    End Function
    Shared Sub creationFichierParJson()
        FileOpen(2, "c:\temp\eq.txt", OpenMode.Output)
        'Dim data As String = Json.SendJson("", "destinations?is_team=true", "AD", "GET")
        Dim tabResultJson As String()
        Dim lineJson As String = ""


        'FileOpen(2, "c:\temp\listep1.txt", OpenMode.Output) 'fichier du personnel sans accent multi equipe


        'Dim destinations = Json.DeserializeJson(data, "destinations")
        For d = 0 To UBound(destinationsJsonIsTrue)
            Dim IDdest As String = destinationsJsonIsTrue(d).ID
            Dim Dest_short_name As String = destinationsJsonIsTrue(d).short_name
            Dim dest_name As String = destinationsJsonIsTrue(d).name
            Dim department_id As String = destinationsJsonIsTrue(d).department_id
            Dim nbrUsersDest As Integer = 0
            Dim team_id As String = destinationsJsonIsTrue(d).team_group
            Dim group_id As String = destinationsJsonIsTrue(d).group_id
            Dim entity_id As String = destinationsJsonIsTrue(d).entity_id

            'data = Json.SendJson("", "destinations/", IDdest & "/teaminfos", "AD", "GET")
            'Dim equipeUser As String = JObject.Parse(data).SelectToken("teaminfo").SelectToken("name")
            'Dim chefEquipeUser As String = JObject.Parse(data).SelectToken("teaminfo").SelectToken("leader_login")
            'If chefEquipeUser = " " Or chefEquipeUser = "" Then chefEquipeUser = "?"
            Dim equipeUser As String = ""
            equipeUser = Commun.RecupEquipeinfo(Dest_short_name)
            If equipeUser = "" Then
                equipeUser = "externe"
            End If

            Dim chefEquipeinfoUser As String = ""
            chefEquipeinfoUser = Commun.TransformeSAMACCOUNTenCN(Commun.FindAttribut(equipeUser & "_eq", "managedBy"))




            ''definir le responsable de l'equipe info
            'If equipeUser <> " " And chefEquipeUser <> "?" And equipeUser <> Nothing And chefEquipeUser <> Nothing Then
            '    Dim volumePartage As DirectoryEntry = New DirectoryEntry("LDAP://CN=ZoneLabo " & equipeUser & ",OU=Partages,OU=Equipes,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr")
            '    volumePartage.Properties("managedBy").Value = TransformeSAMACCOUNTenCN(chefEquipeUser)
            '    volumePartage.CommitChanges()
            '    volumePartage.Close()
            '    volumePartage = Nothing
            'End If

            Dim data As String = Json.SendJson("", "persons?present=true&extern=false&current_destination=true&destination=" & IDdest, "AD", "GET")
            Dim persons = Json.DeserializeJson(data, "persons")
            nbrUsersDest = UBound(persons) + 1
            For p = 0 To UBound(persons)
                Dim lastname As String = persons(p).lastname
                Dim firstname As String = persons(p).firstname
                Dim login As String = persons(p).login
                Dim aliasMail As String = persons(p).email_alias
                Dim IDuser As String = persons(p).ID
                Dim dateEntree As Date = persons(p).entrance_date
                Dim contractRunning As Boolean = ContratEnCours(IDuser)

                'si l'utilisateur n'a pas de login et qu'il est entré il y a moins de 6 mois, on cree un fichier de creation de compte
                'ou que le compte n'existe pas dans l'AD
                If login = "" Or Commun.TransformeSAMACCOUNTenCN(login) = "" Then

                    If Now < dateEntree.AddMonths(6) And contractRunning = True Then
                        Dim dataLD As String = Json.SendJson("", "persons/" & IDuser & "/mailing_lists", "AD", "GET")
                        Dim ld As String = "other"

                        If InStr(dataLD, """name"":""postdoc""") > 0 Then
                            ld = "postdoc"
                        End If

                        If InStr(dataLD, """name"":""phd""") > 0 Then
                            ld = "phd"
                        End If

                        Dim newUser As String = lastname & "," & firstname & "," & login & ",," & Dest_short_name & ",,,," & ld & "," & IDuser & "," & aliasMail
                        Dim sw As New StreamWriter(RecupDataini.RecupVar("[CheminPartage]") & "\todo\c" & Now.ToString("dd") & "-" & Now.ToString("MM") & "-" & Now.ToString("yy") & "-" & Now.ToString("HH") & "h.txt", True)
                        sw.WriteLine(newUser)
                        sw.Close()
                        sw.Dispose()
                    End If
                Else
                    'Si l'utilisateur a un contract en cours
                    If contractRunning = True Then
                        Dim locations As JArray = persons(p).locations
                        Dim loc_building_roomT(locations.Count - 1) As String
                        Dim loc_phoneT(locations.Count - 1) As String
                        For l = 0 To locations.Count - 1
                            'deserialisation d'un objet Json "root"
                            Dim location As Json.locationC = JsonConvert.DeserializeObject(Of Json.locationC)(locations(l).ToString)
                            loc_building_roomT(l) = location.building & " " & Strings.Left(location.room, InStr(location.room, " ") - 1)
                            loc_phoneT(l) = Replace(Replace(location.phone, " ", ""), ".", "")
                            If loc_building_roomT(l) = "" Then loc_building_roomT(l) = "----"
                            If loc_phoneT(l) = "" Then loc_phoneT(l) = "----"
                        Next l
                        Dim BatimentUser As String = Join(loc_building_roomT, ";")
                        Dim TelUser As String = Join(loc_phoneT, ";")
                        Erase loc_phoneT, loc_building_roomT




                        lineJson = lastname & "," & firstname & "," & login & "," & Dest_short_name & "," & dest_name & "," & "100" & "," & BatimentUser & "," & TelUser & "," & equipeUser & "," & chefEquipeinfoUser & "," & IDuser & "," & aliasMail
                        If tabResultJson Is Nothing Then
                            ReDim Preserve tabResultJson(0)
                        Else
                            ReDim Preserve tabResultJson(UBound(tabResultJson) + 1)
                        End If

                        tabResultJson(UBound(tabResultJson)) = lineJson
                    End If
                End If
            Next p

            'Creation du fichier des destinations 
            If nbrUsersDest > 0 Then
                Dim destinationsJsonLeaders As String = Json.SendJson("", "destinations/" & IDdest & "/leaders", "AD", "GET")

                'Recherche du responsable de la destination dans la BDP avec recursivité
                Dim leaderLoginDest As String = LeaderDest(entity_id, group_id, department_id, team_id, IDdest)

                'Ajout des departements au fichier des équipes
                Dim dataDepartments As String = Json.SendJson("", "departments/" & department_id, "AD", "GET")
                'deserialisation d'un objet Json "root"
                Dim departmentsJson As Json.departmentC = JsonConvert.DeserializeObject(Of Json.departmentC)(dataDepartments.ToString)
                Dim depart_Short_name As String = departmentsJson.short_name
                Dim depart_name As String = departmentsJson.name
                Dim depart_ID As String = departmentsJson.id


                PrintLine(2, Dest_short_name & "," & dest_name & "," & IDdest & "," & leaderLoginDest & "," & depart_Short_name & "," & depart_name & "," & depart_ID)
            End If

        Next d
        FileClose(2)

        'inserer les externes dans le tableau
        tabResultJson = RecupJsonExterne(tabResultJson)

        'classement par ordre alphabetique
        tabResultJson = ShellSort(tabResultJson)

        'correction des données multi equipe dans le tableau
        For i = 1 To UBound(tabResultJson)
            Dim id1 As String = Split(tabResultJson(i), ",")(10)
            Dim id0 As String = Split(tabResultJson(i - 1), ",")(10)
            If id1 = id0 Then
                Dim data As String = Json.SendJson("", "persons/" & id0 & "/destinations?is_current_destinations=true&is_team=true&fast=true", "AD", "GET")
                Dim destinations = Json.DeserializeJson(data, "destinations")

                Dim tabDest_short_name As String()
                Dim tabDest_name As String()
                Dim tabDest_time_rate As String()


                For d = 0 To UBound(destinations)
                    ReDim Preserve tabDest_short_name(d)
                    ReDim Preserve tabDest_name(d)
                    ReDim Preserve tabDest_time_rate(d)

                    Dim destination = destinations(d).destination
                    Dim destinationData As Json.destinationC = JsonConvert.DeserializeObject(Of Json.destinationC)(destination.ToString)
                    tabDest_short_name(d) = destinationData.short_name
                    tabDest_name(d) = destinationData.name
                    tabDest_time_rate(d) = destinations(d).time_rate

                    Dim aRemplacer As String = destinationData.short_name & "," & destinationData.name & ",100"
                    Dim remplacerPar As String = destinationData.short_name & "," & destinationData.name & "," & destinations(d).time_rate
                    tabResultJson(i - 1) = Replace(tabResultJson(i - 1), aRemplacer, remplacerPar)
                    tabResultJson(i) = Replace(tabResultJson(i), aRemplacer, remplacerPar)
                Next d

            End If
        Next i
        If tabResultJson Is Nothing Then End
        'Ecriture du fichier
        FileOpen(1, "c:\temp\listepersoJson.txt", OpenMode.Output)
        For i = 0 To UBound(tabResultJson)
            If InStr(tabResultJson(i), ",") > 0 Then
                PrintLine(1, tabResultJson(i))
            End If
        Next i
        FileClose(1)
        Commun.Journal("Création du fichier listepersoJson.txt réussie", False)

    End Sub
    Shared Function RecupJsonExterne(ByVal tabResultjson As String()) As String()

        Dim data As String = Json.SendJson("", "persons?extern=true&present=true", "AD", "GET")

        Dim lineJson As String = ""

        Dim persons = Json.DeserializeJson(data, "persons")
        For p = 0 To UBound(persons)

            Dim lastname As String = persons(p).lastname
            Dim firstname As String = persons(p).firstname
            Dim login As String = persons(p).login
            Dim aliasMail As String = persons(p).email_alias
            'si l'utilisateur n'a pas de login alors on passe au suivant
            If login = "" Then Continue For
            Dim IDuser As String = persons(p).ID
            Dim contractRunning As Boolean = ContratEnCours(IDuser)

            If contractRunning = True Then
                Dim locations As JArray = persons(p).locations

                Dim loc_phoneT(locations.Count - 1) As String
                Dim loc_building_roomT(locations.Count - 1) As String
                For l = 0 To locations.Count - 1
                    Dim location As Json.locationC = JsonConvert.DeserializeObject(Of Json.locationC)(locations(l).ToString)
                    loc_building_roomT(l) = location.building & " " & Strings.Left(location.room, InStr(location.room, " ") - 1)
                    loc_phoneT(l) = location.phone
                    If loc_building_roomT(l) = "" Then loc_building_roomT(l) = "----"
                    If loc_phoneT(l) = "" Then loc_phoneT(l) = "----"
                Next l
                Dim BatimentUser As String = Join(loc_building_roomT, " ; ")
                Dim TelUser As String = Join(loc_phoneT, " ; ")
                Erase loc_phoneT, loc_building_roomT

                data = Json.SendJson("", "persons/" & IDuser & "/destinations?is_current_destinations=true&is_team=true", "AD", "GET")
                Dim destinations = Json.DeserializeJson(data, "destinations")


                lineJson = lastname & "," & firstname & "," & login & ",EXTERNE,EXTERNE,100," & BatimentUser & "," & TelUser & ",externe,guiseith," & IDuser & "," & aliasMail

                If tabResultjson Is Nothing Then
                    ReDim Preserve tabResultjson(0)
                Else
                    ReDim Preserve tabResultjson(UBound(tabResultjson) + 1)
                End If

                tabResultjson(UBound(tabResultjson)) = lineJson
            End If
        Next p
        Return tabResultjson


    End Function
    Shared Sub logSerge(ByVal valeur As String)
        If Not File.Exists("C:\temp\LogAttributUnix.txt") Then
            IO.File.Create("C:\temp\LogAttributUnix.txt")
        End If
        File.AppendAllText("C:\temp\LogAttributUnix.txt", valeur)
    End Sub
    Shared Sub ChangePasswordAccountPrestaImagerie()
        Try
            Dim passwordPrestaImagerieAdm = RandomPassword.Generate(8)
            Dim passwordPrestaImagerieUsr = RandomPassword.Generate(8)
            Using userEntry = New DirectoryEntry("LDAP://" & TransformeSAMACCOUNTenCN("cs-prestaadm"))
                userEntry.Invoke("SetPassword", New Object() {passwordPrestaImagerieAdm})
                userEntry.CommitChanges()
            End Using
            Using userEntry1 = New DirectoryEntry("LDAP://" & TransformeSAMACCOUNTenCN("cs-prestausr"))
                userEntry1.Invoke("SetPassword", New Object() {passwordPrestaImagerieUsr})
                userEntry1.CommitChanges()
            End Using
            Dim textMail = "Bonjour," & vbCrLf & vbCrLf & "Les mots de passe pour les prestataires externes ont changés." & vbCrLf & vbCrLf & "Le mot de passe pour le compte ""cs-prestaadm"" est : " & vbTab & vbTab & passwordPrestaImagerieAdm & vbCrLf & "Le mot de passe pour le compte ""cs-prestausr"" est : " & vbTab & vbTab & passwordPrestaImagerieUsr
            Commun.SendEmail("noreply@igbmc.fr", "groupe-mic-photon@igbmc.fr;Bcc:serviceinfo@igbmc.fr", "Nouveaux mots de passe (" & Now.ToString("MMMM" & " " & Now.ToString("yyyy") & ")"), textMail)

            Commun.Journal("Changement des mots de passe des comptes prestataires de l'imagerie réussi", False)

        Catch ex As Exception
            Commun.Journal("ERREUR : ChangePasswordAccountPrestaImagerie : " & ex.Message, True)
        End Try
    End Sub
    Shared Function TrouverMailPrincipal(ByVal login As String, ByVal courtOuLong As String)
        Try
            Dim mail As String = ""
            Dim objusr As DirectoryEntry = New DirectoryEntry("LDAP://" & TransformeSAMACCOUNTenCN(login))
            Dim currentAddresses = objusr.Properties("proxyAddresses").Value
            Dim adress As New List(Of String)

            If courtOuLong = "court" Then
                For Each value In currentAddresses
                    If Strings.Left(value, 5) = "SMTP:" Then
                        mail = Replace(value, "SMTP:", "")
                        Exit For
                    End If
                Next value
            End If

            If courtOuLong = "long" Then
                For Each value In currentAddresses
                    If Strings.Left(value, 5) = "smtp:" And Strings.Right(value, 8) = "igbmc.fr" Then
                        mail = Replace(value, "smtp:", "")
                        Exit For
                    End If
                Next value
            End If

            If mail = "" Then Throw New Exception("L'utilisateur n'a pas d'adresse mail principale (chercher dans l'attribut ""proxyAddresses"" une valeur commencant par ""SMTP:""")
            Return mail
        Catch ex As Exception
            Commun.Journal("ERREUR : TrouverMailPrincipal : " & login & ":" & ex.Message, True)
        End Try
    End Function
    Shared Sub EnvoiMailCompteExpireXjour(ByVal J As Integer)
        Dim adresseMail As String = ""
        Try
            Dim dateDeSuppressionPrevue As String = Strings.Left(CDate(Now.Day & "/" & Now.Month & "/" & Now.Year).AddDays(J).ToString, 10)
            Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://OU=Comptes Désactivés,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
                Using searcher As DirectorySearcher = New DirectorySearcher(objAD)
                    searcher.PropertiesToLoad.Add("mail")
                    searcher.PropertiesToLoad.Add("givenName")
                    searcher.Filter = "(&(objectClass=person)(objectClass=user)(accountDeletionDate=" & dateDeSuppressionPrevue & "))"


                    Dim results As SearchResultCollection = searcher.FindAll
                    For Each result As SearchResult In results
                        adresseMail = result.Properties("mail")(0)
                        Dim prenom As String = result.Properties("givenName")(0)
                        Dim mail = MailCloture(prenom, dateDeSuppressionPrevue)
                        Commun.SendEmail("noreply@igbmc.fr", adresseMail & ";Cc:serviceinfo@igbmc.fr", "ARRET DU COMPTE", mail)
                        Commun.Journal("Mail de fermeture de compte envoyé (-" & J & ") : " & adresseMail)
                    Next
                End Using
            End Using

            Commun.Journal("Envoi des mails de cloture de compte (j-" & J & ") terminé avec succes", False)
        Catch ex As Exception
            Commun.Journal("ERREUR : EnvoiMailCompteExpireXjour : Mail de fermeture de compte (-" & J & "): " & adresseMail & " : " & ex.Message, True)
        End Try
    End Sub
    Shared Function ContratEnCours(ByVal employeeID As String) As Boolean
        Dim dateactu As Date = Convert.ToDateTime(Now.ToString("yyyy-MM-dd"))

        ContratEnCours = False

        Dim data As String = Json.SendJson("", "persons/" & employeeID & "/contracts", "AD", "GET")

        Dim contracts = Json.DeserializeJson(data, "contracts")
        For c = 0 To UBound(contracts)
            Dim dateStartTxt As String = Strings.Left(contracts(c).start_date, 10)
            Dim dateEndTxt As String = Strings.Left(contracts(c).end_date, 10)
            Dim organism As String = contracts(c).organism_id

            Dim dateStart As Date = #1/1/3000#
            If dateStartTxt <> Nothing Then
                dateStart = Convert.ToDateTime(dateStartTxt)
            End If

            Dim dateEnd As Date = #1/1/3000#
            If dateEndTxt <> Nothing Then
                dateEnd = (Convert.ToDateTime(dateEndTxt)).AddDays(1)
            End If

            If dateStart <= dateactu And dateactu < dateEnd Then
                If organism = "2" Or organism = "3" Or organism = "4" Or organism = "5" Or organism = "12" Or organism = "13" Then
                    ContratEnCours = True
                End If
            End If
        Next c

        Return ContratEnCours

    End Function
    Shared Sub UpdateFichierHistoAlias()
        Dim objAD As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")
        Dim searcher As DirectorySearcher = New DirectorySearcher(objAD)
        searcher.Filter = "(&(proxyAddresses=*))"
        searcher.PageSize = 2000
        Dim results As SearchResultCollection = searcher.FindAll
        For Each result As SearchResult In results
            'si l'objet est dans l'OU NAP on continue avec l'objet suivant
            If InStr(result.Path, "OU=Nap,DC=igbmc,DC=u-strasbg,DC=fr") <> 0 Then Continue For
            Dim objADwithSMTPAddresses As DirectoryEntry = New DirectoryEntry(result.Path)
            'si l'attribut "SMTPAddresses" existe
            If objADwithSMTPAddresses.Properties.Contains("proxyAddresses") Then
                Dim aliasMail As String = ""
                Dim addresses As New List(Of String)
                Dim aliasPrecedent As String = ""
                'Pour toutes les adresses contenue dans l'attribut "proxyAddresses"
                For Each address In objADwithSMTPAddresses.Properties("proxyAddresses").Value
                    Dim addressTemp As String = LCase(address.ToString)
                    'on ne recupere que les adresses mail basé sur @igbmc.fr (pour eviter de recuperer les adresses externes des contacts ou de utilisateurs de messagerie
                    If Strings.Left(addressTemp, 5) = "smtp:" And InStr(addressTemp, "@igbmc.fr") > 0 Then
                        addressTemp = Replace(addressTemp, "smtp:", "")
                        'on recupere l'alias sans l'@ et le domaine
                        aliasMail = Strings.Left(addressTemp, InStr(addressTemp, "@") - 1)

                        Dim userID As String = ""
                        If objADwithSMTPAddresses.Properties.Contains("EmployeeID") Then
                            userID = objADwithSMTPAddresses.Properties("EmployeeID").Value
                        Else
                            userID = 0
                        End If

                        Dim adresseDansLeFichier As Boolean = False
                        'Dim ctrlAlias As Boolean = Commun.ctrlAliasDispo(aliasMail, userID)
                        Dim text As String = File.ReadAllText("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\HistoAlias.txt")
                        Dim index As Integer = text.IndexOf(aliasMail & "," & userID)
                        If index >= 0 Then
                            adresseDansLeFichier = True
                        End If

                        If adresseDansLeFichier = False Then
                            Dim index1 As Integer = text.IndexOf(aliasMail & ",0")
                            If index1 >= 0 Then
                                adresseDansLeFichier = True
                            End If
                        End If

                        'si l'alias ne fait pas deja partie de la liste d'adresse de l'utilisateur et qu'il n'est pas encore dans le fichier
                        If InStr(aliasPrecedent, "§" & aliasMail & "§") = 0 And adresseDansLeFichier = False Then
                            Dim sw As New StreamWriter("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\HistoAlias.txt", True)
                            'on ecrit l'alias dans le fichier
                            If userID <> "" Then
                                sw.WriteLine(aliasMail & "," & userID)
                            Else
                                sw.WriteLine(aliasMail & ",0")
                            End If
                            sw.Close()
                            sw.Dispose()

                        End If
                        aliasPrecedent += "§" & aliasMail & "§"
                    End If
                Next
            End If
            objADwithSMTPAddresses.Close()
            objADwithSMTPAddresses.Dispose()
            objADwithSMTPAddresses = Nothing
        Next
        searcher.Dispose()

        objAD.Close()
        objAD.Dispose()
        objAD = Nothing
        Commun.Journal("Mise a jour du fichier d'historique des Alias terminée avec succes", False)
    End Sub
    Shared Function EstPair(n As Long) As Boolean
        EstPair = (n Mod 2) = 0
    End Function
    Shared Sub SansContrat()
        Dim fichier As String = "c:\temp\PasDeContract.txt"
        If File.Exists(fichier) Then
            Kill(fichier)
        End If

        Dim objAD As DirectoryEntry = New DirectoryEntry("LDAP://OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")
        Dim searcher As DirectorySearcher = New DirectorySearcher(objAD)
        searcher.Filter = "(&(objectClass=person)(objectClass=user))"
        searcher.PropertiesToLoad.Add("EmployeeID")
        searcher.PropertiesToLoad.Add("DisplayName")
        searcher.SearchScope = SearchScope.OneLevel
        Dim results As SearchResultCollection = searcher.FindAll
        For Each result As SearchResult In results
            Dim IDuser As String = result.Properties("EmployeeID")(0)
            Dim ctrlContrat As Boolean = ContratEnCours(IDuser)
            If ctrlContrat = False Then
                Dim sw As New StreamWriter(fichier, True)
                'on ecrit l'alias dans le fichier
                sw.WriteLine(result.Properties("DisplayName")(0) & "," & IDuser)
                sw.Close()
                sw.Dispose()
            End If
        Next
        results.Dispose()
        searcher.Dispose()
        objAD.Close()
        objAD.Dispose()
        Commun.Journal("Ecriture du fichier du personnel sans contrat terminé avec succes", False)
    End Sub
    Shared Function LeaderDest(ByVal idEntity As String, ByVal idGroup As String, ByVal idDep As String, ByVal idTeam As String, ByVal idDest As String) As String
        Dim destinationsJsonLeaders As String = Json.SendJson("", "destinations/" & idDest & "/leaders", "AD", "GET")
        Dim result As String = ""
        Dim leaders = Json.DeserializeJson(destinationsJsonLeaders, "leaders")
        If UBound(leaders) > -1 Then
            result = leaders(0).login()
        End If

        If result = "" Then
            Dim TeamJsonLeaders As String = Json.SendJson("", "entities/" & idEntity & "/groups/" & idGroup & "/departments/" & idDep & "/teams/" & idTeam & "/leaders", "AD", "GET")
            leaders = Json.DeserializeJson(TeamJsonLeaders, "leaders")
            If UBound(leaders) > -1 Then
                result = leaders(0).login()
            End If
        End If

        If result = "" Then
            Dim DptJsonLeaders As String = Json.SendJson("", "departments/" & idDep & "/leaders", "AD", "GET")
            leaders = Json.DeserializeJson(DptJsonLeaders, "leaders")
            If UBound(leaders) > -1 Then
                result = leaders(0).login()
            End If
        End If

        If result = "" Then
            Dim EntityJsonLeaders As String = Json.SendJson("", "entities/" & idEntity & "/leaders", "AD", "GET")
            leaders = Json.DeserializeJson(EntityJsonLeaders, "leaders")
            If UBound(leaders) > -1 Then
                result = leaders(0).login()
            End If
        End If

        Return result
    End Function
End Class