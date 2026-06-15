Imports System.DirectoryServices
Imports System.IO
'Imports ActiveDs
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Collections.ObjectModel
Imports System.Web.Script.Serialization
Imports System.Text
Imports System.Runtime.InteropServices

Module Module1
    Public iniFilePath = "\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\ScriptStephV2.ini"
    Public ini As New IniFile(iniFilePath)
    Public withJson As String = "json" 'Valeur possible : json debug temp(fichiers dans le dossier c:\temp)
    'Public tabPersoMonoEquipe As String(,)
    Public listeUtilisateursRH As New List(Of UtilisateurRH)
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Public tabExcepUser(,) As String
    'Shared organismContractName As String
    Public dossierArchivePST As String = ini.ReadValue("MODIFAUTO", "dossierArchivePST", "\\Space2\archives-pst\") '"\\Space2\archives-pst\"
    Public dossierPhotos As String = ini.ReadValue("MODIFAUTO", "dossierPhotos", "\\Space2\photos-RH\") '"\\Space2\photos-RH\"
    Public sendMSreport As Boolean = ini.ReadValue("MODIFAUTO", "sendMSreport", False)
    Public sendMailOO As Boolean = ini.ReadValue("MODIFAUTO", "sendMailOO", False)

    Public OUUtilisateursExternes As String = ini.ReadValue("MODIFAUTO", "OUUtilisateursExternes")
    Public OUUtilisateurs As String = ini.ReadValue("MODIFAUTO", "OUUtilisateurs")
    Public OUUtilisateursDesactives As String = ini.ReadValue("MODIFAUTO", "OUUtilisateursDesactives")
    Public OUUtilisateursExceptions As String = ini.ReadValue("MODIFAUTO", "OUUtilisateursExceptions")
    Public OUUtilisateursActifs As String = ini.ReadValue("MODIFAUTO", "OUUtilisateursActifs")
    Public OUUtilisateursSortis As String = ini.ReadValue("MODIFAUTO", "OUUtilisateursSortis")

    Dim cheminMAJZoneInfo As String = ini.ReadValue("MODIFAUTO", "cheminMAJZoneInfo")

    Public directeurLogin As String = ini.ReadValue("MODIFAUTO", "LoginDirecteur")

    Public AdminScriptLogin As String = ini.ReadValue("GLOBAL", "AdminScriptLogin")
    Public AdminScriptPassword As String = ini.ReadValue("GLOBAL", "AdminScriptPassword_encrypted")

    Public nomFichierRapportMS As String = "c:\temp\MSrapport(" & Replace(Now.ToString("dd-MM-yyyy HH.mm"), "/", "-") & ").csv"
    Public auth As AuthenticationTypes = AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure ' 'AuthenticationTypes.Secure
    Public usersRH As New List(Of UtilisateurRH)



    Public Sub Main()
        Dim debutScript As DateTime = Now()
        'pour affichage correct de la barre de progression
        Console.OutputEncoding = System.Text.Encoding.UTF8

        Commun.fichierLog = "c:\temp\Log" & application.productname & "_" & Now.ToString("yyyy") & ".log"
        Sleep(3000)
        'Commun.Journal(New String("_"c, Math.Max(1, Console.WindowWidth - 1)), False)
        Commun.Journal("Debut de traitement")

        'choix du DC
        ADHelper.InitialiserDC()

        Dim listExtensionsXivo As String = ""
        If Environment.MachineName <> "SERV-AD1" Then
            'Gestion.GestionGroupeUserActive(New DirectoryEntry("LDAP://serv-ad2.igbmc.u-strasbg.fr/CN=Pietro GIRAUDO,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr", AdminScriptLogin, AdminScriptPassword, auth))
            'Dim dateNowU  = Now.ToUniversalTime.Date.ToString("yyyyMMddHHmmss.sZ")
            'ModEquipeDestinationDepartement.ChargerEquipeDestinationDepartement()
            'Dim a As New Creation
            'Dim aaa As String = a.UserExist("Stephane CERDAN")
            'Commun.ClearAttributDeactivationDeletionDate("steph")
            'UtilisateurADIndexSupprime.SupprimCompteADM()
            'Pws.ForceSyncPeperCutAD("serv-printtools.igbmc.u-strasbg.fr")
            'jsonMS.UserBadgeCodeNumber("6879")
            'Json.SendJson("login=duzgunm&domain=%40igbmc.fr&alias=mikail.duzgun", "persons/8096/email", "AD", "PUT")

        Else
            withJson = "json"
        End If

        'withJson = "debug"

        'Creation et envoi du fichier de controle MicroSesame
        If sendMSreport = True Then
            If Hour(Now) = 3 Or Hour(Now) = 4 Then
                Try
                    ctrlMS.recupUserAD()
                    Commun.Journal("Création et envoi du Raport MicroSesame réussi", False)
                Catch ex As Exception
                    Commun.Journal("ERREUR:Creation du rapport MS" & ex.Message, True)
                End Try
            End If
        End If

        Dim debutCreationFichier As DateTime = Now()
        'creer les comptes des utilisateurs qui sont dans le fichier ForceCreationDeCompte.txt sous la forme <EmployéeID>,<Short_name_destination>
        exceptionCreationCompte()

        Dim ctrlCreationFichier As Boolean = CreationFichiers()
        Dim finCreationFichier As DateTime = Now()
        If ctrlCreationFichier = True Then
            Dim durationCF As TimeSpan = finCreationFichier - debutCreationFichier
            Dim dureeCF As String = String.Format("{0:00}h {1:00}m {2:00}s", durationCF.Hours, durationCF.Minutes, durationCF.Seconds)
            Dim dureeTolerable As Integer = Convert.ToInt32(ini.ReadValue("MODIFAUTO", "DureeMaxJson"))

            If DateDiff(DateInterval.Minute, DebutCreationFichier, finCreationFichier) > dureeTolerable Then
                Commun.SendEmail("administrateur@igbmc.fr", ini.ReadValue("MODIFAUTO", "mailDureeMaxJson"), "Duree de creation de fichier JSON superieure à " & dureeTolerable.ToString & " minutes", "La durée de creation du fichier JSON a travers IGBMC services a été anormalement longue: " & dureeCF)
            End If

            Commun.Journal("Création du fichier listepersoJson.txt réussie en : " & dureeCF, False)

            GestionDesFichiers()

            'ne pas sauter GestionDestinationsDepartements, construction du dictionnaire equipeinfo
            gestion.GestionDestinationsDepartements()

            Dim usersRaw As Dictionary(Of String, UserRaw) = LoadUsersFromFile("c:\temp\listepersoJson.txt")
            Dim dnParEmployeeId As Dictionary(Of String, String) = ChargerIndexDnParEmployeeId()
            Dim badgesParEmployeeId As Dictionary(Of String, String()) = jsonMS.ChargerBadgesParEmployeeId()

            Dim adUsersByEmployeeId As Dictionary(Of String, UtilisateurADIndex) = ChargerIndexUtilisateursAD()

            usersRH.Clear()
            For Each kvp As KeyValuePair(Of String, UserRaw) In usersRaw
                usersRH.Add(ConvertToUtilisateurRH(kvp.Value, dnParEmployeeId, badgesParEmployeeId))
            Next

            ModifDonneesAD(usersRH, adUsersByEmployeeId)

            If Environment.MachineName = "SERV-AD1" Then
                gestion.GestionReactiveDesactiveComptesInterne(adUsersByEmployeeId)
            End If

        End If


        GestionComptesExternes()

        Gestion.CompleterDatesContratManquantesComptesDesactivesEtSortis()

        'La gestion des AttributsDT doit imperativement intervenir apres GestionReactiveDesactiveComptesInterne
        Gestion.GestionAttributsDT()

        Gestion.GestionSuppressionProfilsItinerantsEtDossiersRedirigés()

        AttributionStrategieMDP()
        gestion.CtrlGroupAdmins()
        gestion.UpdateComptesProvisoires()
        gestion.ControleOUUtilisateurs()


        If Hour(Now) = 1 Or Hour(Now) = 2 Then
            'Gestion de l'expiration des mot de passe des comptes adm

            Commun.Journal("Gestion de l'expiration des comptes/ mots de passe", False)
            ExpirationMDP()

            Commun.Journal("Debut de la gestion de l'envoi des mails de cloture de compte", False)
            EnvoiMailCompteExpireXjour(30)
            EnvoiMailCompteExpireXjour(15)
            EnvoiMailCompteExpireXjour(7)
            EnvoiMailCompteExpireXjour(2)
            EnvoiMailCompteExpireXjour(1)
            Commun.Journal("Fin de la gestion de l'envoi des mails de cloture de compte", False)


            'Changement tous les mois pair, le premier du mois, des mots de passe des comptes prestataires de l'imagerie
            If System.DateTime.Now.ToString("dd") = "01" And EstPair(System.DateTime.Now.ToString("MM")) = True Then
                ChangePasswordAccountPrestaImagerie()
            End If

            UpdateFichierHistoAlias()
            Supprime.SupprimeMailbox()
            Supprime.DeleteOldPST()
        End If

        Try
            Commun.Journal("Lancement de la synchronisation de PaperCut avec l'AD", False)
            Pws.ForceSyncPeperCutAD("serv-printtools.igbmc.u-strasbg.fr")
            Commun.Journal("Fin de la synchronisation de PaperCut avec l'AD", False)
        Catch ex As Exception
            Commun.Journal(vbTab & "ERREUR : Lancement de la synchronisation de PaperCut avec l'AD : " & ex.Message, True)
        End Try


        If Commun.controlSendMail = True Then
            sendJournalError()
        End If

        Try
            Shell(cheminMAJZoneInfo)
        Catch
            Commun.Journal("ERREUR: Lancement de MAJZoneInfo", True)
        End Try

        If File.Exists(nomFichierRapportMS) Then
            File.Delete(nomFichierRapportMS)
        End If

        Dim finScript As DateTime = Now()
        Dim durationScript As TimeSpan = finScript - debutScript
        Dim dureeScript As String = String.Format("{0:00}h {1:00}m {2:00}s", durationScript.Hours, durationScript.Minutes, durationScript.Seconds)

        Commun.Journal("Fin de traitement en : " & dureeScript)
        'Commun.Journal("____________________________________________________________________________________________________")

        ini.WriteValue("MODIFAUTO", "lastExec", Now.ToString("dd/MM/yyyy HH:mm:ss"))
    End Sub
    Public Sub sendJournalError()
        Commun.SendEmail(AdminScriptLogin & "@igbmc.fr", "steph@igbmc.fr", "ModifAuto.NET : Rapport d'erreur", Commun.journalECHECMail)
    End Sub

    Public Sub GestionDesFichiers()
        Dim cheminPartage As String = ini.ReadValue("MODIFAUTO", "CheminPartage")
        ' COPIE DES FICHIERS 
        Try

            File.Copy("c:\temp\eq.txt", CheminPartage & "\todo\eq.txt", True)
            File.Copy("c:\temp\listep.txt", cheminPartage & "\todo\listep.txt", True)
            If withJson = "json" Then
                File.Copy("c:\temp\listepersoJson.txt", cheminPartage & "\todo\listepersoJson.txt", True)
                File.Copy("c:\temp\listepersoJson.txt", cheminPartage & "\cmpttmp\listepersoJson" & Now.ToString("dd") & "-" & Now.ToString("MM") & "-" & Now.ToString("yy") & "-" & Now.ToString("HH") & "h.txt", True)
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : COPIE : Copie du fichier : " & ex.Message, True)
        End Try

        Sleep(5000)

        Try

            'Kill("c:\temp\listepersoJson.txt")

        Catch
            Commun.Journal("ERREUR : suppression du fichier temporaire listepersoJson.txt", True)
        End Try

        Try
            'Kill("c:\temp\listep.txt")
        Catch
            Commun.Journal("ERREUR : suppression du fichier temporaire listep.txt", True)
        End Try
        Try
            'Kill("c:\temp\eq.txt")
        Catch
            Commun.Journal("ERREUR : suppression du fichier temporaire eq.txt", True)
        End Try
    End Sub

    ''' <summary>
    ''' Met à jour les comptes Active Directory à partir des utilisateurs RH normalisés.
    ''' </summary>
    ''' <param name="usersRH">
    ''' Collection d'utilisateurs RH déjà normalisés pour la comparaison avec l'AD.
    ''' </param>
    ''' <summary>
    ''' Met à jour les comptes Active Directory à partir des utilisateurs RH normalisés.
    ''' </summary>
    ''' <param name="usersRH">
    ''' Collection d'utilisateurs RH déjà normalisés pour la comparaison avec l'AD.
    ''' </param>

    Public Sub ModifDonneesAD(usersRH As IEnumerable(Of UtilisateurRH), adUsersByEmployeeId As Dictionary(Of String, UtilisateurADIndex))
        Commun.Journal("Debut des Modification des Utilisateurs", False)

        Dim ctrlMailOOrienteurs As Boolean = False
        Dim corpmailOOrienteurs As String = ""

        Dim tabPhoto(,) As String = Json.GetMyIGBMC()
        Commun.Journal(vbTab & "Photos Modifiées Récupérées", False)

        Dim usersRHList As List(Of UtilisateurRH) = usersRH.ToList()

        Dim max As Integer = usersRHList.Count
        Dim n As Integer = 0



        For Each userRH As UtilisateurRH In usersRHList
            n += 1
            Commun.AfficherBarre("Traitement des utilisateurs", n, max, False)

            Dim userAD As UtilisateurADIndex = Nothing
            If adUsersByEmployeeId.ContainsKey(userRH.employeeID_id) Then
                userAD = adUsersByEmployeeId(userRH.employeeID_id)
            End If
            Dim changementPhotoAFaire As Boolean = (Commun.MultiIndexOf(tabPhoto, userRH.employeeID_id, 0) <> -1)
            ChargerJpegPhotoDepuisTabPhoto(userRH, tabPhoto)
            PreparerUidNumber(userAD, userRH)
            PreparerDatesDesactivationSuppression(userRH)

            Dim changementsAD As List(Of ChangementAttributAD) = UtilisateurADDiffereDeRH(userAD, userRH, changementPhotoAFaire)
            If changementsAD.Count = 0 Then
                Continue For
            End If

            Dim changementsSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim changementsMap As New Dictionary(Of String, ChangementAttributAD)(StringComparer.OrdinalIgnoreCase)

            For Each changement As ChangementAttributAD In changementsAD
                changementsSet.Add(changement.Attribut)
                changementsMap(changement.Attribut) = changement
            Next

            Commun.Journal(vbTab & vbTab & "Changements détectés pour " & userRH.login_samAccountName, False)
            If userAD Is Nothing Then
                Commun.Journal(vbTab & vbTab & "ERREUR : Modification des données : " & userRH.login_samAccountName & " : utilisateur absent de l'index AD", True)
                Continue For
            End If

            Using objuser As New DirectoryEntry("LDAP://" & Commun.LdapPath(userAD.distinguishedName), Commun.admin, Commun.passwd, auth)
                Try
                    Dim prop As String = "jpegPhoto"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPPropertyByte(objuser, prop, userRH.jpegPhoto_jpegPhoto)

                        If userRH.jpegPhoto_jpegPhoto IsNot Nothing Then
                            Dim imageByteTh As Byte() = Thumbn.CreateThumb1(userRH.jpegPhoto_jpegPhoto, userRH.employeeID_id)
                            Commun.SetADLDAPPropertyByte(objuser, "thumbnailPhoto", imageByteTh)
                        Else
                            Commun.SetADLDAPPropertyByte(objuser, "thumbnailPhoto", Nothing)
                        End If

                        Commun.AppliquerChangement(objuser)
                        userAD.jpegPhoto = userRH.jpegPhoto_jpegPhoto
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""jpegPhoto"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "sn"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.nom_sn)
                        Commun.AppliquerChangement(objuser)
                        userAD.sn = userRH.nom_sn
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""sn"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "givenName"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.prenom_givenName)
                        Commun.AppliquerChangement(objuser)
                        userAD.givenName = userRH.prenom_givenName
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""givenName"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "physicalDeliveryOfficeName"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.physicalDeliveryOfficeName_bureaux)
                        Commun.AppliquerChangement(objuser)
                        userAD.physicalDeliveryOfficeName = userRH.physicalDeliveryOfficeName_bureaux
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""physicalDeliveryOfficeName"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "extensionAttribute2"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.extensionAttribute2_genre)
                        Commun.AppliquerChangement(objuser)
                        userAD.extensionAttribute2 = userRH.extensionAttribute2_genre
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""extensionAttribute2"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "division"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.division_organism)
                        Commun.AppliquerChangement(objuser)
                        userAD.division = userRH.division_organism
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""division"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "sAMAccountName"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.login_samAccountName)
                        Commun.AppliquerChangement(objuser)
                        userAD.samAccountName = userRH.login_samAccountName
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""sAMAccountName"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "mail"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.mail_mailPrincipal)
                        Commun.AppliquerChangement(objuser)
                        userAD.mail = userRH.mail_mailPrincipal
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""mail"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "mailNickname"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.login_samAccountName)
                        Commun.AppliquerChangement(objuser)
                        userAD.mailNickname = userRH.login_samAccountName
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""mailNickname"" : {ex.Message}", True)
                End Try

                Dim telephoneMobile As String = userRH.mobile_telPortable
                Dim telephonePrincipal As String = userRH.telephoneNumber_telPrincipal
                Dim tabOtherphoneFichier As String() = userRH.otherTelephone_telSecondaire

                Try
                    Dim prop As String = "otherTelephone"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPPropertyMulti(objuser, prop, tabOtherphoneFichier)
                        Commun.AppliquerChangement(objuser)
                        userAD.otherTelephone = If(tabOtherphoneFichier IsNot Nothing, tabOtherphoneFichier, Array.Empty(Of String)())
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""otherTelephone"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "mobile"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, telephoneMobile)
                        Commun.AppliquerChangement(objuser)
                        userAD.mobile = If(telephoneMobile IsNot Nothing, telephoneMobile, "")
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""mobile"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "telephoneNumber"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, telephonePrincipal)
                        Commun.AppliquerChangement(objuser)
                        userAD.telephoneNumber = If(telephonePrincipal IsNot Nothing, telephonePrincipal, "")
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""telephoneNumber"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "company"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, "IGBMC")
                        Commun.AppliquerChangement(objuser)
                        userAD.company = "IGBMC"
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""company"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "department"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.department_destinationNomLong)
                        Commun.AppliquerChangement(objuser)
                        userAD.department = userRH.department_destinationNomLong
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""department"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "departmentNumber"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.departmentNumber_destinationNomCourt)
                        Commun.AppliquerChangement(objuser)
                        userAD.departmentNumber = userRH.departmentNumber_destinationNomCourt
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""departmentNumber"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "title"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.title_unité)
                        Commun.AppliquerChangement(objuser)
                        userAD.title = userRH.title_unité
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""title"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "extensionAttribute1"
                    If changementsSet.Contains(prop) Then
                        If objuser.Parent.Path = "LDAP://" & Commun.LdapPath(OUUtilisateursActifs) Then
                            EnvoyerMailAPSiNecessaire(userRH, userAD)

                            Dim oldDate As String = userAD.extensionAttribute1
                            Dim dateCreation As Date = If(userAD.whenCreated.HasValue, userAD.whenCreated.Value, Now)
                            Dim tmp As Integer = DateDiff(DateInterval.Hour, dateCreation, Now)

                            If tmp > 3 Then
                                If String.IsNullOrEmpty(oldDate) Then oldDate = "Aucune"

                                Dim newDate As String = userRH.extensionAttribute1_finDeContrat
                                If String.IsNullOrEmpty(newDate) Then newDate = "Aucune"

                                Dim idMS As String = jsonMS.GetIdMS(userRH.employeeID_id)
                                jsonMS.SetMSEndValidity(idMS, newDate)

                                If sendMailOO = True Then
                                    Try
                                        Dim habilitationsExpirées As String = jsonMS.GetMSAccreditation(idMS, newDate)

                                        AjouterNotificationOfficierOrienteur(
                                        userRH,
                                        oldDate,
                                        newDate,
                                        habilitationsExpirées,
                                        ctrlMailOOrienteurs,
                                        corpmailOOrienteurs
                                    )
                                    Catch ex As DirectoryNotFoundException
                                        Commun.Journal(vbTab & vbTab & $"ERREUR : Lecture des habilitations expirées pour les Officiers orienteurs : {userRH.login_samAccountName} : {ex.Message}", True)
                                    End Try
                                End If
                            End If

                            Commun.SetADLDAPProperty(objuser, prop, userRH.extensionAttribute1_finDeContrat)
                            Commun.AppliquerChangement(objuser)
                            userAD.extensionAttribute1 = userRH.extensionAttribute1_finDeContrat
                            Dim ch As ChangementAttributAD = changementsMap(prop)
                            Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                        End If
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""extensionAttribute1"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "accountDeletionDate"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.accountDeletionDate_finDeContratPlus3Mois)
                        Commun.AppliquerChangement(objuser)
                        userAD.accountDeletionDate = userRH.accountDeletionDate_finDeContratPlus3Mois

                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""accountDeletionDate"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "accountDeactivationDT"
                    If changementsSet.Contains(prop) Then
                        If userRH.accountDeactivationDT_finDeContrat.HasValue Then
                            objuser.Properties(prop).Value = userRH.accountDeactivationDT_finDeContrat.Value.Date
                        Else
                            objuser.Properties(prop).Clear()
                        End If

                        Commun.AppliquerChangement(objuser)
                        userAD.accountDeactivationDT = userRH.accountDeactivationDT_finDeContrat

                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""accountDeactivationDT"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "accountDeletionDT"
                    If changementsSet.Contains(prop) Then
                        If userRH.accountDeletionDT_finDeContratPlus3Mois.HasValue Then
                            objuser.Properties(prop).Value = userRH.accountDeletionDT_finDeContratPlus3Mois.Value.Date
                        Else
                            objuser.Properties(prop).Clear()
                        End If

                        Commun.AppliquerChangement(objuser)
                        userAD.accountDeletionDT = userRH.accountDeletionDT_finDeContratPlus3Mois

                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""accountDeletionDT"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "uidNumber"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.uidNumber_uidNumber)
                        Commun.AppliquerChangement(objuser)
                        userAD.uidNumber = userRH.uidNumber_uidNumber
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""uidNumber"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "uid"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.uid_uid)
                        Commun.AppliquerChangement(objuser)
                        userAD.uid = userRH.uid_uid
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""uid"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "gidNumber"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.gidNumberAD_gidNumber)
                        Commun.AppliquerChangement(objuser)
                        userAD.gidNumber = userRH.gidNumberAD_gidNumber
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""gidNumber"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "unixHomeDirectory"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.unixHomeDirectoryAD_unixHomeDirectoryAD)
                        Commun.AppliquerChangement(objuser)
                        userAD.unixHomeDirectory = userRH.unixHomeDirectoryAD_unixHomeDirectoryAD
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""unixHomeDirectory"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "loginShell"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, "/bin/bash")
                        Commun.AppliquerChangement(objuser)
                        userAD.loginShell = "/bin/bash"
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""loginShell"" : {ex.Message}", True)
                End Try
                Try
                    Dim prop As String = "serialNumber"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPPropertyMulti(objuser, prop, userRH.serialNumber_serialNumber)
                        Commun.AppliquerChangement(objuser)
                        userAD.serialNumber = If(userRH.serialNumber_serialNumber IsNot Nothing, userRH.serialNumber_serialNumber, Array.Empty(Of String)())
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""serialNumber"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "employeeNumber"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.employeeNumber_employeeNumber)
                        Commun.AppliquerChangement(objuser)
                        userAD.employeeNumber = userRH.employeeNumber_employeeNumber
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""employeeNumber"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "manager"
                    If changementsSet.Contains(prop) Then
                        If userRH.login_samAccountName = ini.ReadValue("MODIFAUTO", "LoginDirecteur") Then
                            Commun.SetADLDAPProperty(objuser, prop, "")
                            Commun.AppliquerChangement(objuser)
                            userAD.manager = ""
                        Else
                            Commun.SetADLDAPProperty(objuser, prop, userRH.DNNplus1)
                            Commun.AppliquerChangement(objuser)
                            userAD.manager = userRH.DNNplus1
                        End If

                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""manager"" : {ex.Message}", True)
                End Try
                Try
                    Dim prop As String = "memberOf"
                    If changementsSet.Contains(prop) Then
                        SynchroniserGroupesUtilisateur(userAD, userRH)

                        userAD.memberOf_groupesGerables = If(userRH.memberOf_groupeRH, New String() {})
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""memberOf"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "displayName"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.displayName)
                        Commun.AppliquerChangement(objuser)
                        userAD.displayName = userRH.displayName
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""displayName"" : {ex.Message}", True)
                End Try

                Try
                    Dim prop As String = "displayNamePrintable"
                    If changementsSet.Contains(prop) Then
                        Commun.SetADLDAPProperty(objuser, prop, userRH.displayNamePrintable)
                        Commun.AppliquerChangement(objuser)
                        userAD.displayNamePrintable = userRH.displayNamePrintable
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""displayNamePrintable"" : {ex.Message}", True)
                End Try

                If changementsSet.Contains("cn") Then
                    Try
                        modifAliasMail(objuser, userRH.prenom_givenName, userRH.nom_sn)
                    Catch ex As Exception
                        Commun.Journal(vbTab & "ERREUR : Modification : Ajout d'Alias : " & userRH.login_samAccountName & " : " & ex.Message, True)
                    End Try

                    Try
                        objuser.Rename("CN=" & userRH.cn)
                        userAD.cn = userRH.cn
                        Dim ch As ChangementAttributAD = changementsMap("cn")
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    Catch ex As Exception
                        Commun.Journal(vbTab & "ERREUR : Modification : Renomer l'objet utilisateur : " & userRH.login_samAccountName & " : " & ex.Message, True)
                    End Try
                End If

                Try
                    Dim prop As String = "msExchExtensionAttribute16"
                    If changementsSet.Contains(prop) Then
                        Dim nouvelleValeur As String = MettreAJourMsExchExtensionAttribute16SansCommit(userRH)
                        Commun.SetADLDAPProperty(objuser, prop, nouvelleValeur)
                        Commun.AppliquerChangement(objuser)
                        userAD.msExchExtensionAttribute16 = nouvelleValeur
                        Dim ch As ChangementAttributAD = changementsMap(prop)
                        Commun.Journal(vbTab & vbTab & vbTab & $"Modification de l'attribut ""{ch.Attribut}"" Réussi: {userRH.login_samAccountName} : [{ch.AncienneValeur}] -> [{ch.NouvelleValeur}]", False)
                    End If
                Catch ex As Exception
                    Commun.Journal(vbTab & vbTab & $"ERREUR : Modification de l'attribut : ""msExchExtensionAttribute16"" : {ex.Message}", True)
                End Try
            End Using
        Next

        If sendMailOO = True AndAlso ctrlMailOOrienteurs = True Then
            corpmailOOrienteurs &= vbCrLf & "</body></html>"
            Commun.SendEmail("administrateur@igbmc.fr", "officiersorienteurs@igbmc.fr", "Changement de Date de fin de contrat", corpmailOOrienteurs)
            Commun.Journal(vbTab & "Envoi d'un mail aux Officiers orienteurs pour un changement de fin de contrat")
        End If

        Commun.Journal("Gestion des modifications utilisateurs réussie", False)
    End Sub
    Private Sub PreparerUidNumber(userAD As UtilisateurADIndex, userRH As UtilisateurRH)
        If userAD IsNot Nothing AndAlso String.IsNullOrWhiteSpace(userAD.uidNumber) Then
            userRH.uidNumber_uidNumber = Commun.UIDNumberMini()
        End If
    End Sub

    Private Sub ChargerJpegPhotoDepuisTabPhoto(userRH As UtilisateurRH, tabPhoto(,) As String)
        Dim rechPhoto As Integer = Commun.MultiIndexOf(tabPhoto, userRH.employeeID_id, 0)

        If rechPhoto = -1 Then
            'Commun.Journal("PHOTO RH ABSENTE : " & userRH.login_samAccountName & " | employeeID=" & userRH.employeeID_id, True)
            userRH.jpegPhoto_jpegPhoto = Nothing
            Exit Sub
        End If

        'Commun.Journal("PHOTO RH TROUVEE INDEX : " & userRH.login_samAccountName & " | employeeID=" & userRH.employeeID_id & " | url/id=" & tabPhoto(1, rechPhoto), False)

        userRH.jpegPhoto_jpegPhoto = Json.LirePhotoDepuisMyIGBMC(tabPhoto(1, rechPhoto))

        If userRH.jpegPhoto_jpegPhoto Is Nothing Then
            'Commun.Journal("PHOTO RH LUE = Nothing : " & userRH.login_samAccountName & " | employeeID=" & userRH.employeeID_id & " | url/id=" & tabPhoto(1, rechPhoto), True)
        Else
            'Commun.Journal("PHOTO RH OK : " & userRH.login_samAccountName & " | bytes=" & userRH.jpegPhoto_jpegPhoto.Length, False)
        End If
    End Sub
    Private Sub PreparerDatesDesactivationSuppression(userRH As UtilisateurRH)
        Dim culture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.InvariantCulture
        Dim style As System.Globalization.DateTimeStyles = System.Globalization.DateTimeStyles.None

        Dim dateContrat As Date
        Dim dateException As Date
        Dim dateBase As Date

        Dim contratValide As Boolean =
        Date.TryParseExact(userRH.extensionAttribute1_finDeContrat,
                           "dd/MM/yyyy",
                           culture,
                           style,
                           dateContrat)

        If Not contratValide Then
            userRH.accountDeactivationDT_finDeContrat = Nothing
            userRH.accountDeletionDate_finDeContratPlus3Mois = ""
            userRH.accountDeletionDT_finDeContratPlus3Mois = Nothing
            Exit Sub
        End If

        dateBase = dateContrat

        Dim finException As String = Gestion.exceptionUser(userRH.login_samAccountName)
        If finException <> "False" Then
            If Date.TryParseExact(finException,
                              "dd/MM/yyyy",
                              culture,
                              style,
                              dateException) Then
                If dateException > dateContrat Then
                    dateBase = dateException
                End If
            End If
        End If

        userRH.accountDeactivationDT_finDeContrat = dateBase.Date
        userRH.accountDeletionDate_finDeContratPlus3Mois = dateBase.AddMonths(3).ToString("dd/MM/yyyy")
        userRH.accountDeletionDT_finDeContratPlus3Mois = dateBase.AddMonths(3).Date
    End Sub

    Private Function TrouverDestinationsPourUtilisateur(userRH As UtilisateurRH) As List(Of DestinationInfo)
        Dim resultat As New List(Of DestinationInfo)
        Dim dejaAjoutees As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If userRH Is Nothing Then
            Return resultat
        End If

        For Each destinationUser As String In If(userRH.destinationsRH, New String() {})
            If String.IsNullOrWhiteSpace(destinationUser) Then
                Continue For
            End If

            For Each destination As DestinationInfo In DicoDestinationsRH.Values
                If destination Is Nothing Then
                    Continue For
                End If

                If String.Equals(destination.nom_court_dest, destinationUser, StringComparison.OrdinalIgnoreCase) Then
                    If Not dejaAjoutees.Contains(destination.id_dest) Then
                        resultat.Add(destination)
                        dejaAjoutees.Add(destination.id_dest)
                    End If
                End If
            Next
        Next

        If resultat.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(userRH.departmentNumber_destinationNomCourt) Then
            For Each destination As DestinationInfo In DicoDestinationsRH.Values
                If destination Is Nothing Then
                    Continue For
                End If

                If String.Equals(destination.nom_court_dest, userRH.departmentNumber_destinationNomCourt, StringComparison.OrdinalIgnoreCase) Then
                    If Not dejaAjoutees.Contains(destination.id_dest) Then
                        resultat.Add(destination)
                        dejaAjoutees.Add(destination.id_dest)
                    End If
                End If
            Next
        End If

        Return resultat
    End Function


    Private Function ConstruireGroupesRH(userRH As UtilisateurRH) As String()
        Dim groupes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If userRH Is Nothing Then
            Return groupes.ToArray()
        End If

        Dim destinations As List(Of DestinationInfo) = TrouverDestinationsPourUtilisateur(userRH)

        For Each destination As DestinationInfo In destinations
            If destination Is Nothing Then
                Continue For
            End If

            If Not String.IsNullOrWhiteSpace(destination.dn_dest_ad) Then
                groupes.Add(destination.dn_dest_ad)
            End If

            If Not String.IsNullOrWhiteSpace(destination.id_equipeinfo) AndAlso DicoEquipesInfoRefRH.ContainsKey(destination.id_equipeinfo) Then
                Dim equipe As EquipeInfo = DicoEquipesInfoRefRH(destination.id_equipeinfo)
                If equipe IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(equipe.dn_equipeinfo_eq_ad) Then
                    groupes.Add(equipe.dn_equipeinfo_eq_ad)
                End If
            End If

            If Not String.IsNullOrWhiteSpace(destination.id_dept) AndAlso DicoDepartementsRH.ContainsKey(destination.id_dept) Then
                Dim departement As DepartementInfo = DicoDepartementsRH(destination.id_dept)
                If departement IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(departement.dn_dept_ad) Then
                    groupes.Add(departement.dn_dept_ad)
                End If
            End If
        Next

        Dim listeDiff As String = LCase(Trim(userRH.listesDiffusions))
        If listeDiff <> "" AndAlso DicoGroupesDiffusionRH.ContainsKey(listeDiff) Then
            groupes.Add(DicoGroupesDiffusionRH(listeDiff))
        End If

        Return groupes.ToArray()
    End Function

    Private Sub SynchroniserGroupesUtilisateur(userAD As UtilisateurADIndex, userRH As UtilisateurRH)
        Dim groupesAD As String() = If(userAD.memberOf_groupesGerables, New String() {})
        Dim groupesRH As String() = If(userRH.memberOf_groupeRH, New String() {})

        Dim groupesADSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each groupeAD As String In groupesAD
            If Not String.IsNullOrWhiteSpace(groupeAD) Then
                groupesADSet.Add(groupeAD)
            End If
        Next

        Dim groupesRHSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each groupeRH As String In groupesRH
            If Not String.IsNullOrWhiteSpace(groupeRH) Then
                groupesRHSet.Add(groupeRH)
            End If
        Next

        For Each groupeRH As String In groupesRHSet
            If Not groupesADSet.Contains(groupeRH) Then
                Commun.AddRemoveADGroup(userAD.distinguishedName, groupeRH, "Add")
                Commun.Journal(vbTab & vbTab & vbTab & "Ajout groupe : " & userRH.login_samAccountName & " => " & NomGroupeDepuisDn(groupeRH), False)
            End If
        Next

        For Each groupeAD As String In groupesADSet
            If Not groupesRHSet.Contains(groupeAD) Then
                Commun.AddRemoveADGroup(userAD.distinguishedName, groupeAD, "Remove")
                Commun.Journal(vbTab & vbTab & vbTab & "Retrait groupe : " & userRH.login_samAccountName & " => " & NomGroupeDepuisDn(groupeAD), False)
            End If
        Next
    End Sub
    Private Function NomGroupeDepuisDn(dnGroupe As String) As String
        If String.IsNullOrWhiteSpace(dnGroupe) Then
            Return ""
        End If

        If dnGroupe.StartsWith("CN=", StringComparison.OrdinalIgnoreCase) Then
            Dim finCN As Integer = dnGroupe.IndexOf(","c)
            If finCN > 3 Then
                Return dnGroupe.Substring(3, finCN - 3)
            End If

            Return dnGroupe.Substring(3)
        End If

        Return dnGroupe
    End Function
    Private Function FiltrerGroupesGerables(groupes As String()) As String()
        Dim resultat As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each groupe As String In If(groupes, New String() {})
            If String.IsNullOrWhiteSpace(groupe) Then
                Continue For
            End If

            Dim nomGroupe As String = NomGroupeDepuisDn(groupe)
            If nomGroupe.EndsWith(" grp", StringComparison.OrdinalIgnoreCase) _
            OrElse nomGroupe.EndsWith("_eq", StringComparison.OrdinalIgnoreCase) _
            OrElse (nomGroupe.StartsWith("Dpt_", StringComparison.OrdinalIgnoreCase) AndAlso Not nomGroupe.StartsWith("Dpt_PI", StringComparison.OrdinalIgnoreCase)) _
            OrElse String.Equals(nomGroupe, "phd", StringComparison.OrdinalIgnoreCase) _
            OrElse String.Equals(nomGroupe, "postdoc", StringComparison.OrdinalIgnoreCase) Then
                resultat.Add(groupe)
            End If
        Next

        Return resultat.ToArray()
    End Function
    Private Sub EnvoyerMailAPSiNecessaire(
    ByVal userRH As UtilisateurRH,
    ByVal userAD As UtilisateurADIndex
)
        If userRH.extensionAttribute1_finDeContrat <> "" AndAlso If(userAD Is Nothing, "", userAD.extensionAttribute1) = "" Then
            Exit Sub
        End If

        Dim ctrlEnvoiMailAP As Boolean = GetContractsLenght(userRH.employeeID_id, False)
        If ctrlEnvoiMailAP = False Then
            Exit Sub
        End If

        Dim corpmailAssistentsPrévention As String =
        vbCrLf & "Nom : " & userRH.prenom_givenName & " " & userRH.nom_sn &
        vbCrLf & "Identifiant GDPI : " & userRH.employeeID_id &
        vbCrLf & "Service : " & userRH.department_destinationNomLong &
        vbCrLf & "Unité : " & userRH.title_unité &
        vbCrLf & "Mail : " & userRH.login_samAccountName & "@igbmc.fr" &
        vbCrLf & "Date de fin de contrat : " & userRH.extensionAttribute1_finDeContrat

        Commun.SendEmail("administrateur@igbmc.fr", "assistants-de-prevention@igbmc.fr;Bcc:steph@igbmc.fr", "(Mail automatique) Nouvel entrant", corpmailAssistentsPrévention)
    End Sub


    Public Function ChargerIndexUtilisateursAD() As Dictionary(Of String, UtilisateurADIndex)

        Dim adUsersByEmployeeId As New Dictionary(Of String, UtilisateurADIndex)(StringComparer.OrdinalIgnoreCase)

        Using ldap As New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateurs), Commun.admin, Commun.passwd, auth)
            Using searcher As New DirectorySearcher(ldap)

                searcher.Filter = "(&(objectClass=user)(employeeID=*))"
                searcher.SearchScope = SearchScope.Subtree
                searcher.PageSize = 1000
                searcher.PropertiesToLoad.Clear()

                searcher.PropertiesToLoad.Add("employeeID")
                searcher.PropertiesToLoad.Add("distinguishedName")
                searcher.PropertiesToLoad.Add("departmentNumber")
                searcher.PropertiesToLoad.Add("sn")
                searcher.PropertiesToLoad.Add("givenName")
                searcher.PropertiesToLoad.Add("mail")
                searcher.PropertiesToLoad.Add("mailNickname")
                searcher.PropertiesToLoad.Add("physicalDeliveryOfficeName")
                searcher.PropertiesToLoad.Add("telephoneNumber")
                searcher.PropertiesToLoad.Add("mobile")
                searcher.PropertiesToLoad.Add("manager")
                searcher.PropertiesToLoad.Add("division")
                searcher.PropertiesToLoad.Add("department")
                searcher.PropertiesToLoad.Add("title")
                searcher.PropertiesToLoad.Add("extensionAttribute1")
                searcher.PropertiesToLoad.Add("extensionAttribute2")
                searcher.PropertiesToLoad.Add("displayName")
                searcher.PropertiesToLoad.Add("displayNamePrintable")
                searcher.PropertiesToLoad.Add("company")
                searcher.PropertiesToLoad.Add("otherTelephone")
                searcher.PropertiesToLoad.Add("whenCreated")
                searcher.PropertiesToLoad.Add("uidNumber")
                searcher.PropertiesToLoad.Add("sAMAccountName")
                searcher.PropertiesToLoad.Add("loginShell")
                searcher.PropertiesToLoad.Add("uid")
                searcher.PropertiesToLoad.Add("gidNumber")
                searcher.PropertiesToLoad.Add("unixHomeDirectory")
                searcher.PropertiesToLoad.Add("cn")
                searcher.PropertiesToLoad.Add("msExchExtensionAttribute16")
                searcher.PropertiesToLoad.Add("memberOf")
                searcher.PropertiesToLoad.Add("employeeNumber")
                searcher.PropertiesToLoad.Add("serialNumber")
                searcher.PropertiesToLoad.Add("accountActivationDT")
                searcher.PropertiesToLoad.Add("accountDeactivationDT")
                searcher.PropertiesToLoad.Add("accountDeletionDT")
                searcher.PropertiesToLoad.Add("accountDeletionDate")
                searcher.PropertiesToLoad.Add("description")
                searcher.PropertiesToLoad.Add("comment")
                searcher.PropertiesToLoad.Add("jpegPhoto")
                searcher.PropertiesToLoad.Add("memberOf")

                Using results As SearchResultCollection = searcher.FindAll()
                    For Each r As SearchResult In results

                        Dim employeeId As String = LireProp(r, "employeeID")

                        If employeeId = "" Then
                            Continue For
                        End If

                        If Not adUsersByEmployeeId.ContainsKey(employeeId) Then
                            Dim u As New UtilisateurADIndex()

                            u.employeeID = employeeId
                            u.distinguishedName = LireProp(r, "distinguishedName")
                            u.departmentNumber = LireProp(r, "departmentNumber")
                            u.sn = LireProp(r, "sn")
                            u.givenName = LireProp(r, "givenName")
                            u.mail = LireProp(r, "mail")
                            u.mailNickname = LireProp(r, "mailNickname")
                            u.physicalDeliveryOfficeName = LireProp(r, "physicalDeliveryOfficeName")
                            u.telephoneNumber = LireProp(r, "telephoneNumber")
                            u.mobile = LireProp(r, "mobile")
                            u.manager = LireProp(r, "manager")
                            u.division = LireProp(r, "division")
                            u.department = LireProp(r, "department")
                            u.title = LireProp(r, "title")
                            u.extensionAttribute1 = LireProp(r, "extensionAttribute1")
                            u.extensionAttribute2 = LireProp(r, "extensionAttribute2")
                            u.displayName = LireProp(r, "displayName")
                            u.displayNamePrintable = LireProp(r, "displayNamePrintable")
                            u.company = LireProp(r, "company")
                            u.otherTelephone = LirePropMulti(r, "otherTelephone")
                            u.whenCreated = LirePropDate(r, "whenCreated")
                            u.uidNumber = LireProp(r, "uidNumber")
                            u.samAccountName = LireProp(r, "sAMAccountName")
                            u.loginShell = LireProp(r, "loginShell")
                            u.uid = LireProp(r, "uid")
                            u.gidNumber = LireProp(r, "gidNumber")
                            u.unixHomeDirectory = LireProp(r, "unixHomeDirectory")
                            u.cn = LireProp(r, "cn")
                            u.msExchExtensionAttribute16 = LireProp(r, "msExchExtensionAttribute16")
                            u.memberOf = LirePropMulti(r, "memberOf")
                            u.memberOf_groupesGerables = FiltrerGroupesGerables(u.memberOf)
                            u.employeeNumber = LireProp(r, "employeeNumber")
                            u.serialNumber = LirePropMulti(r, "serialNumber")
                            u.accountActivationDT = LirePropDate(r, "accountActivationDT")
                            u.accountDeactivationDT = LirePropDate(r, "accountDeactivationDT")
                            u.accountDeletionDT = LirePropDate(r, "accountDeletionDT")
                            u.accountDeletionDate = LireProp(r, "accountDeletionDate")
                            u.description = LireProp(r, "description")
                            u.comment = LireProp(r, "comment")
                            u.jpegPhoto = LirePropByte(r, "jpegPhoto")

                            adUsersByEmployeeId.Add(employeeId, u)
                        End If

                    Next
                End Using

            End Using
        End Using

        Return adUsersByEmployeeId

    End Function

    Private Function MettreAJourMsExchExtensionAttribute16SansCommit(userRH As UtilisateurRH) As String
        Dim aliasmailAD As String = LCase(BuildAliasMail(userRH.prenom_givenName, userRH.nom_sn))
        Dim aliasMailpourIGBMCSERVICES As String = Replace(aliasmailAD, "@igbmc.fr", "")

        Json.SendJson("login=" & userRH.login_samAccountName &
                  "&domain=%40igbmc.fr&alias=" & aliasMailpourIGBMCSERVICES,
                  "persons/" & userRH.employeeID_id & "/email", "AD", "PUT")

        Return aliasmailAD
    End Function

    Private Function LireProp(r As SearchResult, nomProp As String) As String
        If r Is Nothing Then Return ""
        If Not r.Properties.Contains(nomProp) Then Return ""
        If r.Properties(nomProp).Count = 0 Then Return ""

        Dim valeur As Object = r.Properties(nomProp)(0)
        If valeur Is Nothing Then Return ""

        Return CStr(valeur).Trim()
    End Function
    Private Function LirePropByte(r As SearchResult, prop As String) As Byte()
        If r Is Nothing Then Return Nothing
        If Not r.Properties.Contains(prop) Then Return Nothing
        If r.Properties(prop).Count = 0 Then Return Nothing
        If r.Properties(prop)(0) Is Nothing Then Return Nothing

        Return DirectCast(r.Properties(prop)(0), Byte())
    End Function

    Private Function LirePropMulti(r As SearchResult, nomProp As String) As String()
        If r Is Nothing Then Return New String() {}
        If Not r.Properties.Contains(nomProp) Then Return New String() {}
        If r.Properties(nomProp).Count = 0 Then Return New String() {}

        Dim valeurs(r.Properties(nomProp).Count - 1) As String

        For i As Integer = 0 To r.Properties(nomProp).Count - 1
            If r.Properties(nomProp)(i) IsNot Nothing Then
                valeurs(i) = CStr(r.Properties(nomProp)(i)).Trim()
            Else
                valeurs(i) = ""
            End If
        Next

        Return valeurs
    End Function

    Private Function LirePropDate(r As SearchResult, nomProp As String) As Date?
        If r Is Nothing Then Return Nothing
        If Not r.Properties.Contains(nomProp) Then Return Nothing
        If r.Properties(nomProp).Count = 0 Then Return Nothing

        Dim valeur As Object = r.Properties(nomProp)(0)
        If valeur Is Nothing Then Return Nothing

        If TypeOf valeur Is Date Then
            Return CType(valeur, Date)
        End If

        Dim s As String = CStr(valeur).Trim()
        If s = "" Then Return Nothing

        Dim dt As DateTime
        If DateTime.TryParseExact(
        s,
        "yyyyMMddHHmmss'.0Z'",
        Globalization.CultureInfo.InvariantCulture,
        Globalization.DateTimeStyles.AssumeUniversal Or Globalization.DateTimeStyles.AdjustToUniversal,
        dt) Then

            Return dt
        End If

        Return Nothing
    End Function
    ''' <summary>
    ''' Construit l'alias mail long attendu au format prénom.nom@igbmc.fr.
    ''' </summary>
    ''' <param name="prenom">
    ''' Prénom de l'utilisateur.
    ''' </param>
    ''' <param name="nom">
    ''' Nom de l'utilisateur.
    ''' </param>
    ''' <returns>
    ''' Retourne l'alias mail long normalisé.
    ''' </returns>
    Private Function BuildAliasMail(prenom As String, nom As String) As String
        Dim aliasConstruit As String = LCase(prenom & "." & nom)
        aliasConstruit = Replace(aliasConstruit, " ", "-")
        aliasConstruit = Replace(aliasConstruit, "'", "")
        Return aliasConstruit & "@igbmc.fr"
    End Function

    ''' <summary>
    ''' Compare les attributs métier entre l'utilisateur AD indexé et l'utilisateur RH normalisé.
    ''' </summary>
    ''' <param name="adUser">
    ''' Utilisateur lu depuis l'index AD.
    ''' </param>
    ''' <param name="userRH">
    ''' Utilisateur RH normalisé, contenant les valeurs attendues côté métier.
    ''' </param>
    ''' <returns>
    ''' Retourne une chaîne contenant la liste des attributs différents, séparés par des virgules.
    ''' Retourne une chaîne vide si aucune différence n'est détectée.
    ''' Retourne <c>ADUserMissing</c> si l'utilisateur n'existe pas dans l'index AD.
    ''' </returns>
    ''' <remarks>
    ''' Cette fonction sert de point central de comparaison métier entre l'état attendu
    ''' (<c>UtilisateurRH</c>) et l'état connu dans l'AD (<c>UtilisateurADIndex</c>).
    ''' Elle ne doit comparer que des attributs métier simples, et non relire directement
    ''' les propriétés live de <c>DirectoryEntry</c>.
    ''' </remarks>

    Private Function UtilisateurADDiffereDeRH(adUser As UtilisateurADIndex, userRH As UtilisateurRH, changementPhotoAFaire As Boolean) As List(Of ChangementAttributAD)
        Dim changements As New List(Of ChangementAttributAD)

        If adUser Is Nothing Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "ADUserMissing",
            .AncienneValeur = "",
            .NouvelleValeur = userRH.employeeID_id
        })
            Return changements
        End If

        If adUser.employeeID <> userRH.employeeID_id Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "employeeID",
            .AncienneValeur = adUser.employeeID,
            .NouvelleValeur = userRH.employeeID_id
        })
        End If

        If adUser.departmentNumber <> userRH.departmentNumber_destinationNomCourt Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "departmentNumber",
            .AncienneValeur = adUser.departmentNumber,
            .NouvelleValeur = userRH.departmentNumber_destinationNomCourt
        })
        End If

        If adUser.sn <> userRH.nom_sn Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "sn",
            .AncienneValeur = adUser.sn,
            .NouvelleValeur = userRH.nom_sn
        })
        End If

        If adUser.givenName <> userRH.prenom_givenName Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "givenName",
            .AncienneValeur = adUser.givenName,
            .NouvelleValeur = userRH.prenom_givenName
        })
        End If

        If adUser.mail <> userRH.mail_mailPrincipal Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "mail",
            .AncienneValeur = adUser.mail,
            .NouvelleValeur = userRH.mail_mailPrincipal
        })
        End If

        If adUser.mailNickname <> userRH.login_samAccountName Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "mailNickname",
            .AncienneValeur = adUser.mailNickname,
            .NouvelleValeur = userRH.login_samAccountName
        })
        End If

        If adUser.physicalDeliveryOfficeName <> userRH.physicalDeliveryOfficeName_bureaux Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "physicalDeliveryOfficeName",
            .AncienneValeur = adUser.physicalDeliveryOfficeName,
            .NouvelleValeur = userRH.physicalDeliveryOfficeName_bureaux
        })
        End If

        If adUser.telephoneNumber <> userRH.telephoneNumber_telPrincipal Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "telephoneNumber",
            .AncienneValeur = adUser.telephoneNumber,
            .NouvelleValeur = userRH.telephoneNumber_telPrincipal
        })
        End If

        If Join(TrierTableau(adUser.otherTelephone), ";") <> Join(TrierTableau(userRH.otherTelephone_telSecondaire), ";") Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "otherTelephone",
            .AncienneValeur = Join(If(adUser.otherTelephone, New String() {}), ";"),
            .NouvelleValeur = Join(If(userRH.otherTelephone_telSecondaire, New String() {}), ";")
        })
        End If

        If adUser.mobile <> userRH.mobile_telPortable Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "mobile",
            .AncienneValeur = adUser.mobile,
            .NouvelleValeur = userRH.mobile_telPortable
        })
        End If

        If adUser.manager <> userRH.DNNplus1 Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "manager",
            .AncienneValeur = adUser.manager,
            .NouvelleValeur = userRH.DNNplus1
        })
        End If

        If adUser.division <> userRH.division_organism Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "division",
            .AncienneValeur = adUser.division,
            .NouvelleValeur = userRH.division_organism
        })
        End If

        If adUser.department <> userRH.department_destinationNomLong Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "department",
            .AncienneValeur = adUser.department,
            .NouvelleValeur = userRH.department_destinationNomLong
        })
        End If

        If adUser.title <> userRH.title_unité Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "title",
            .AncienneValeur = adUser.title,
            .NouvelleValeur = userRH.title_unité
        })
        End If

        If adUser.extensionAttribute1 <> userRH.extensionAttribute1_finDeContrat Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "extensionAttribute1",
            .AncienneValeur = adUser.extensionAttribute1,
            .NouvelleValeur = userRH.extensionAttribute1_finDeContrat
        })
        End If

        If adUser.extensionAttribute2 <> userRH.extensionAttribute2_genre Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "extensionAttribute2",
            .AncienneValeur = adUser.extensionAttribute2,
            .NouvelleValeur = userRH.extensionAttribute2_genre
        })
        End If

        If adUser.displayName <> userRH.displayName Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "displayName",
            .AncienneValeur = adUser.displayName,
            .NouvelleValeur = userRH.displayName
        })
        End If

        If adUser.displayNamePrintable <> userRH.displayNamePrintable Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "displayNamePrintable",
            .AncienneValeur = adUser.displayNamePrintable,
            .NouvelleValeur = userRH.displayNamePrintable
        })
        End If

        If adUser.company <> "IGBMC" Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "company",
            .AncienneValeur = adUser.company,
            .NouvelleValeur = "IGBMC"
        })
        End If

        If adUser.samAccountName <> userRH.login_samAccountName Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "sAMAccountName",
            .AncienneValeur = adUser.samAccountName,
            .NouvelleValeur = userRH.login_samAccountName
        })
        End If

        If adUser.cn <> userRH.cn Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "cn",
            .AncienneValeur = adUser.cn,
            .NouvelleValeur = userRH.cn
        })
        End If

        If adUser.msExchExtensionAttribute16 <> LCase(BuildAliasMail(userRH.prenom_givenName, userRH.nom_sn)) Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "msExchExtensionAttribute16",
            .AncienneValeur = adUser.msExchExtensionAttribute16,
            .NouvelleValeur = LCase(BuildAliasMail(userRH.prenom_givenName, userRH.nom_sn))
        })
        End If

        If Join(TrierTableau(adUser.memberOf_groupesGerables), ";") <> Join(TrierTableau(userRH.memberOf_groupeRH), ";") Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "memberOf",
            .AncienneValeur = Join(If(adUser.memberOf_groupesGerables, New String() {}), ";"),
            .NouvelleValeur = Join(If(userRH.memberOf_groupeRH, New String() {}), ";")
        })
        End If

        If adUser.uid <> userRH.uid_uid Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "uid",
            .AncienneValeur = adUser.uid,
            .NouvelleValeur = userRH.uid_uid
        })
        End If

        If adUser.gidNumber <> userRH.gidNumberAD_gidNumber Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "gidNumber",
            .AncienneValeur = adUser.gidNumber,
            .NouvelleValeur = userRH.gidNumberAD_gidNumber
        })
        End If

        If adUser.unixHomeDirectory <> userRH.unixHomeDirectoryAD_unixHomeDirectoryAD Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "unixHomeDirectory",
            .AncienneValeur = adUser.unixHomeDirectory,
            .NouvelleValeur = userRH.unixHomeDirectoryAD_unixHomeDirectoryAD
        })
        End If

        If String.IsNullOrWhiteSpace(adUser.loginShell) Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "loginShell",
            .AncienneValeur = adUser.loginShell,
            .NouvelleValeur = "/bin/bash"
        })
        End If

        If adUser.employeeNumber <> userRH.employeeNumber_employeeNumber Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "employeeNumber",
            .AncienneValeur = adUser.employeeNumber,
            .NouvelleValeur = userRH.employeeNumber_employeeNumber
        })
        End If

        If Join(TrierTableau(adUser.serialNumber), ";") <> Join(TrierTableau(userRH.serialNumber_serialNumber), ";") Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "serialNumber",
            .AncienneValeur = Join(If(adUser.serialNumber, New String() {}), ";"),
            .NouvelleValeur = Join(If(userRH.serialNumber_serialNumber, New String() {}), ";")
        })
        End If

        If String.IsNullOrWhiteSpace(adUser.uidNumber) Then
            userRH.uidNumber_uidNumber = Commun.UIDNumberMini()
            changements.Add(New ChangementAttributAD With {
            .Attribut = "uidNumber",
            .AncienneValeur = adUser.uidNumber,
            .NouvelleValeur = userRH.uidNumber_uidNumber
        })
        End If

        If changementPhotoAFaire AndAlso Not PhotosSontEgales(adUser.jpegPhoto, userRH.jpegPhoto_jpegPhoto) Then
            changements.Add(New ChangementAttributAD With {
            .Attribut = "jpegPhoto",
            .AncienneValeur = If(adUser.jpegPhoto Is Nothing, "", "[photo AD]"),
            .NouvelleValeur = If(userRH.jpegPhoto_jpegPhoto Is Nothing, "", "[photo RH]")
        })
        End If

        Dim extensionAttribute1Renseigne As Boolean =
        Not String.IsNullOrWhiteSpace(adUser.extensionAttribute1) OrElse
        Not String.IsNullOrWhiteSpace(userRH.extensionAttribute1_finDeContrat)

        If DatesDifferentes(adUser.accountDeactivationDT, userRH.accountDeactivationDT_finDeContrat) Then
            If userRH.accountDeactivationDT_finDeContrat.HasValue OrElse Not extensionAttribute1Renseigne Then
                changements.Add(New ChangementAttributAD With {
                .Attribut = "accountDeactivationDT",
                .AncienneValeur = If(adUser.accountDeactivationDT.HasValue, adUser.accountDeactivationDT.Value.ToString("dd/MM/yyyy"), ""),
                .NouvelleValeur = If(userRH.accountDeactivationDT_finDeContrat.HasValue, userRH.accountDeactivationDT_finDeContrat.Value.ToString("dd/MM/yyyy"), "")
            })
            End If
        End If

        If DatesDifferentes(adUser.accountDeletionDT, userRH.accountDeletionDT_finDeContratPlus3Mois) Then
            If userRH.accountDeletionDT_finDeContratPlus3Mois.HasValue OrElse Not extensionAttribute1Renseigne Then
                changements.Add(New ChangementAttributAD With {
                .Attribut = "accountDeletionDT",
                .AncienneValeur = If(adUser.accountDeletionDT.HasValue, adUser.accountDeletionDT.Value.ToString("dd/MM/yyyy"), ""),
                .NouvelleValeur = If(userRH.accountDeletionDT_finDeContratPlus3Mois.HasValue, userRH.accountDeletionDT_finDeContratPlus3Mois.Value.ToString("dd/MM/yyyy"), "")
            })
            End If
        End If

        If adUser.accountDeletionDate <> userRH.accountDeletionDate_finDeContratPlus3Mois Then
            If Not String.IsNullOrWhiteSpace(userRH.accountDeletionDate_finDeContratPlus3Mois) OrElse Not extensionAttribute1Renseigne Then
                changements.Add(New ChangementAttributAD With {
                .Attribut = "accountDeletionDate",
                .AncienneValeur = adUser.accountDeletionDate,
                .NouvelleValeur = userRH.accountDeletionDate_finDeContratPlus3Mois
            })
            End If
        End If

        Return changements
    End Function
    Private Function DatesDifferentes(date1 As Date?, date2 As Date?) As Boolean
        If date1.HasValue <> date2.HasValue Then
            Return True
        End If

        If Not date1.HasValue Then
            Return False
        End If

        Return date1.Value.Date <> date2.Value.Date
    End Function

    Private Function PhotosSontEgales(photo1 As Byte(), photo2 As Byte()) As Boolean
        If photo1 Is Nothing AndAlso photo2 Is Nothing Then Return True
        If photo1 Is Nothing OrElse photo2 Is Nothing Then Return False
        If photo1.Length <> photo2.Length Then Return False

        For i As Integer = 0 To photo1.Length - 1
            If photo1(i) <> photo2(i) Then Return False
        Next

        Return True
    End Function

    Public Function TrierTableau(valeurs As String()) As String()
        If valeurs Is Nothing Then
            Return New String() {}
        End If

        Dim copie As String() = CType(valeurs.Clone(), String())

        Array.Sort(copie, StringComparer.OrdinalIgnoreCase)

        Return copie
    End Function
    ''' <summary>
    ''' Construit les données téléphoniques métier à partir d'un utilisateur RH brut.
    ''' </summary>
    ''' <param name="userRaw">
    ''' Utilisateur RH brut.
    ''' </param>
    ''' <param name="userRH">
    ''' Utilisateur RH normalisé à compléter.
    ''' </param>
    Private Sub ConstruireDonneesTelephoniques(userRaw As UserRaw, userRH As UtilisateurRH)
        Dim telephonesNormalises As String = NormalizePhoneNumbers(userRaw.TelUser)

        userRH.mobile_telPortable = GetMobileTelephone(telephonesNormalises)
        userRH.telephoneNumber_telPrincipal = GetTelephonePrincipalHorsMobile(telephonesNormalises)
        userRH.otherTelephone_telSecondaire = GetOtherTelephonesHorsMobile(telephonesNormalises)
    End Sub
    Private Function NormaliserTelephonePourComparaison(numero As String) As String
        If String.IsNullOrWhiteSpace(numero) Then Return ""

        Dim n As String = numero.Trim()

        n = n.Replace(" ", "")
        n = n.Replace(".", "")
        n = n.Replace("-", "")
        n = n.Replace("(", "")
        n = n.Replace(")", "")

        Return n
    End Function
    Private Function EstTelephoneMobile(numero As String) As Boolean
        If String.IsNullOrWhiteSpace(numero) Then Return False

        Dim n As String = NormaliserTelephonePourComparaison(numero)

        Return n.StartsWith("06") _
        OrElse n.StartsWith("07") _
        OrElse n.StartsWith("+336") _
        OrElse n.StartsWith("+337") _
        OrElse n.StartsWith("+33 6") _
        OrElse n.StartsWith("+33 7")
    End Function
    Private Function DecouperTelephones(telephonesNormalises As String) As List(Of String)
        Dim resultat As New List(Of String)

        If String.IsNullOrWhiteSpace(telephonesNormalises) Then
            Return resultat
        End If

        For Each tel As String In telephonesNormalises.Split(";"c)
            Dim valeur As String = tel.Trim()

            If valeur <> "" Then
                resultat.Add(valeur)
            End If
        Next

        Return resultat
    End Function
    Private Function GetMobileTelephone(telephonesNormalises As String) As String
        Dim telephones As List(Of String) = DecouperTelephones(telephonesNormalises)

        For Each tel As String In telephones
            If EstTelephoneMobile(tel) Then
                Return tel
            End If
        Next

        Return ""
    End Function
    Private Function GetOtherTelephonesHorsMobile(telephonesNormalises As String) As String()
        Dim telephones As List(Of String) = DecouperTelephones(telephonesNormalises)
        Dim resultat As New List(Of String)
        Dim telephonePrincipalTrouve As Boolean = False

        For Each tel As String In telephones
            If Not EstTelephoneMobile(tel) Then
                If Not telephonePrincipalTrouve Then
                    telephonePrincipalTrouve = True
                Else
                    resultat.Add(tel)
                End If
            End If
        Next

        Return resultat.ToArray()
    End Function
    '''' <summary>
    '''' Extrait les téléphones secondaires depuis la chaîne normalisée des téléphones.
    '''' </summary>
    '''' <param name="telephones">
    '''' Chaîne des téléphones séparés par des point-virgules.
    '''' </param>
    '''' <returns>
    '''' Retourne les numéros secondaires sous forme de tableau.
    '''' </returns>
    'Private Function GetOtherTelephones(telephones As String) As String()
    '    If String.IsNullOrWhiteSpace(telephones) Then Return New String() {}

    '    Dim tabPhoneFichier As String() = Split(telephones, ";")
    '    If tabPhoneFichier Is Nothing OrElse tabPhoneFichier.Length <= 1 Then
    '        Return New String() {}
    '    End If

    '    Dim result As New List(Of String)

    '    For i As Integer = 1 To UBound(tabPhoneFichier)
    '        If String.IsNullOrWhiteSpace(tabPhoneFichier(i)) Then
    '            Continue For
    '        End If

    '        If InStr(tabPhoneFichier(i), "/") > 0 Then
    '            Dim tmpTab As String() = Split(tabPhoneFichier(i), "/")
    '            For Each tel As String In tmpTab
    '                If Not String.IsNullOrWhiteSpace(tel) Then
    '                    result.Add(tel.Trim())
    '                End If
    '            Next
    '        Else
    '            result.Add(tabPhoneFichier(i).Trim())
    '        End If
    '    Next

    '    Return result.ToArray()
    'End Function

    Private Function GetTelephonePrincipalHorsMobile(telephonesNormalises As String) As String
        Dim telephones As List(Of String) = DecouperTelephones(telephonesNormalises)

        For Each tel As String In telephones
            If Not EstTelephoneMobile(tel) Then
                Return tel
            End If
        Next

        Return ""
    End Function
    ''' <summary>
    ''' Applique la règle spécifique Xivo sur la valeur métier de l'attribut ipPhone.
    ''' </summary>
    ''' <param name="userRH">
    ''' Utilisateur RH normalisé.
    ''' </param>
    ''' <param name="listExtensionsXivo">
    ''' Contenu brut retourné par l'API Xivo pour les lignes SIP.
    ''' </param>
    ''' <returns>
    ''' Retourne la valeur finale attendue pour l'attribut <c>ipPhone</c>.
    ''' </returns>
    ''' <remarks>
    ''' La valeur de base d'ipPhone doit déjà avoir été construite dans <c>UtilisateurRH.ipPhone</c>.
    ''' Cette fonction n'applique que la règle spécifique dépendante de Xivo.
    ''' </remarks>
    'Private Function BuildIpPhoneValueAvecXivo(userRH As UtilisateurRH, listExtensionsXivo As String) As String
    '    Dim ipPhone As String = userRH.ipPhone

    '    If String.IsNullOrWhiteSpace(ipPhone) Then
    '        Return ""
    '    End If

    '    If String.IsNullOrWhiteSpace(listExtensionsXivo) Then
    '        Return ipPhone
    '    End If

    '    If Len(ipPhone) <> 4 Then
    '        Return ipPhone
    '    End If

    '    Try
    '        Dim root As JObject = JObject.Parse(listExtensionsXivo)
    '        Dim items As JToken = root("items")

    '        If items Is Nothing OrElse items.Type <> JTokenType.Array Then
    '            Return ipPhone
    '        End If

    '        Dim nomComplet As String = LCase(userRH.prenom_givenName & " " & userRH.nom_sn)

    '        For Each item As JToken In items
    '            Dim callerid As String = item("callerid")?.ToString()

    '            If String.IsNullOrWhiteSpace(callerid) Then
    '                Continue For
    '            End If

    '            If LCase(callerid).Contains("""" & nomComplet & """") Then
    '                Return ""
    '            End If
    '        Next

    '    Catch ex As Exception
    '        Throw New Exception("Erreur lors de l'analyse du JSON Xivo : " & ex.Message)
    '    End Try

    '    Return ipPhone
    'End Function

    ''' <summary>
    ''' Convertit un <see cref="UserRaw"/> en <c>UtilisateurRH</c> pour rester compatible
    ''' avec le reste du code métier qui manipule encore cette classe historique.
    ''' </summary>
    ''' <param name="user">
    ''' Utilisateur brut issu du dictionnaire chargé depuis le fichier RH.
    ''' </param>
    ''' <returns>
    ''' Retourne une instance de <c>UtilisateurRH</c> alimentée à partir de <see cref="UserRaw"/>.
    ''' </returns>
    ''' <remarks>
    ''' Cette méthode sert de pont entre le nouveau modèle <see cref="UserRaw"/> et l'ancien code
    ''' qui attend encore un objet <c>UtilisateurRH</c>.
    ''' </remarks>
    Private Function ConvertToUtilisateurRH(user As UserRaw, dnParEmployeeId As Dictionary(Of String, String), badgesParEmployeeId As Dictionary(Of String, String())) As UtilisateurRH
        Dim userRH As New UtilisateurRH()

        userRH.nom_sn = user.lastname
        userRH.prenom_givenName = user.firstname
        userRH.login_samAccountName = user.login
        userRH.departmentNumber_destinationNomCourt = user.Dest_short_name
        userRH.department_destinationNomLong = user.dest_name
        userRH.tempsTravail = user.time_rate
        userRH.physicalDeliveryOfficeName_bureaux = user.BatimentUser

        ConstruireDonneesTelephoniques(user, userRH)

        userRH.equipeInfo = user.equipeUser

        Dim idEquipeInfo As String = ObtenirIdEquipeInfoDepuisNomCourt(userRH.equipeInfo)

        If idEquipeInfo <> "" AndAlso DicoEquipesInfoRefRH.ContainsKey(idEquipeInfo) Then
            userRH.equipeInfoDico = DicoEquipesInfoRefRH(idEquipeInfo)
        Else
            userRH.equipeInfoDico = Nothing
        End If

        userRH.manager_IDNplus1 = user.Nplus1ID

        If user.Nplus1ID <> "" AndAlso dnParEmployeeId.ContainsKey(user.Nplus1ID) Then
            userRH.DNNplus1 = dnParEmployeeId(user.Nplus1ID)
        Else
            userRH.DNNplus1 = ""
        End If
        userRH.employeeID_id = user.IDuser
        userRH.msExchExtensionAttribute16_aliasMailLong = user.aliasMail
        userRH.listesDiffusions = user.ld
        userRH.division_organism = user.organism
        userRH.extensionAttribute1_finDeContrat = user.finContrat
        userRH.extensionAttribute2_genre = user.genre
        userRH.title_unité = user.uniteName

        Dim badges As String() = Nothing

        If badgesParEmployeeId IsNot Nothing AndAlso badgesParEmployeeId.ContainsKey(userRH.employeeID_id) Then
            badges = badgesParEmployeeId(userRH.employeeID_id)
        End If

        userRH.serialNumber_serialNumber = TrierTableau(badges)

        If userRH.serialNumber_serialNumber IsNot Nothing AndAlso userRH.serialNumber_serialNumber.Length = 1 Then
            userRH.employeeNumber_employeeNumber = userRH.serialNumber_serialNumber(0)
        Else
            userRH.employeeNumber_employeeNumber = ""
        End If

        userRH.memberOf_groupeRH = ConstruireGroupesRH(userRH)
        userRH.destinationsRH = If(user.destinationsRH, New String() {})
        userRH.memberOf_groupeRH = ConstruireGroupesRH(userRH)

        Return userRH
    End Function

    '''' <summary>
    '''' Construit la valeur attendue pour l'attribut AD <c>ipPhone</c>.
    '''' </summary>
    '''' <param name="telephones">
    '''' Chaîne des téléphones séparés par des point-virgules.
    '''' </param>
    '''' <returns>
    '''' Retourne la valeur calculée pour <c>ipPhone</c>.
    '''' </returns>
    'Private Function BuildIpPhoneValue(telephones As String) As String
    '    Dim ipPhone As String = GetTelephonePrincipal(telephones)

    '    ipPhone = Replace(ipPhone, "(+3336948) ", "")
    '    ipPhone = Replace(ipPhone, "(+3338865) ", "")
    '    ipPhone = Replace(ipPhone, ")", "")
    '    ipPhone = Replace(ipPhone, "+33 ", "0")

    '    Return ipPhone
    'End Function
    Private Function ChargerIndexDnParEmployeeId() As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Using ldap As New DirectoryEntry("LDAP://" & Commun.LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
            Using searcher As New DirectorySearcher(ldap)
                searcher.Filter = "(&(objectClass=user)(employeeID=*))"
                searcher.SearchScope = SearchScope.Subtree
                searcher.PageSize = 1000
                searcher.PropertiesToLoad.Clear()
                searcher.PropertiesToLoad.Add("employeeID")
                searcher.PropertiesToLoad.Add("distinguishedName")

                Using results As SearchResultCollection = searcher.FindAll()
                    For Each r As SearchResult In results
                        Dim employeeId As String = LireProp(r, "employeeID")
                        Dim dn As String = LireProp(r, "distinguishedName")

                        If employeeId <> "" AndAlso Not result.ContainsKey(employeeId) Then
                            result.Add(employeeId, dn)
                        End If
                    Next
                End Using
            End Using
        End Using

        Return result
    End Function

    ''' <summary>
    ''' Prépare les fichiers de travail, charge la liste des utilisateurs,
    ''' fusionne les doublons multi-équipe et génère le fichier final mono-équipe.
    ''' </summary>
    ''' <returns>
    ''' Retourne <c>True</c> si le traitement s'est terminé correctement ; sinon <c>False</c>.
    ''' </returns>
    ''' <remarks>
    ''' Le comportement dépend de la variable globale <c>withJson</c>.
    ''' En mode <c>json</c>, les données sont régénérées depuis igbmcservices.
    ''' En mode <c>debug</c>, des fichiers de référence sont copiés dans <c>c:\temp</c>.
    ''' </remarks>
    Public Function CreationFichiers() As Boolean
        Dim result As Boolean = False

        ' Première étape : selon le mode choisi, on prépare les fichiers source.
        If withJson = "json" Then
            Commun.Journal("Debut de recuperation des données de igbmcservices")

            ' Appel HTTP vers igbmcservices pour récupérer les destinations de type équipe.
            Dim dataJsonDestinationTrue As String = Json.SendJson("", "destinations?is_team=true", "AD", "GET")
            If dataJsonDestinationTrue Is Nothing Then End

            Dim tabDestinationsJsonIsTrue = Json.DeserializeJson(dataJsonDestinationTrue, "destinations")

            ' Cette méthode existante génère ensuite les fichiers texte exploités par le reste du traitement.
            creationFichierParJson(tabDestinationsJsonIsTrue)

        ElseIf withJson = "debug" Then
            Commun.Journal("Mode DEBUG : recuperation des fichiers de données")

            ' En mode debug, on recopie simplement des fichiers de référence dans c:\temp.
            File.Copy(ini.ReadValue("MODIFAUTO", "CheminPartage") & "\todo\eq.txt", "c:\temp\eq.txt", True)
            File.Copy(ini.ReadValue("MODIFAUTO", "CheminPartage") & "\todo\listepersoJson.txt", "c:\temp\listepersoJson.txt", True)

        ElseIf withJson = "temp" Then
            ' Rien à faire ici pour l'instant.

        Else
            End
        End If

        Try
            ' Charge tous les utilisateurs depuis le fichier source.
            ' Le dictionnaire est indexé par IDuser pour fusionner automatiquement les doublons.
            Dim users As Dictionary(Of String, UserRaw) = LoadUsersFromFile("c:\temp\listepersoJson.txt")
            Dim referentielDestinationsDepartements As ReferentielDestinationsDepartements = LoadReferentielDestinationsDepartementsFromFile("c:\temp\eq.txt")

            If referentielDestinationsDepartements Is Nothing Then
                Throw New Exception("Referentiel destinations/departements non chargé")
            End If

            If referentielDestinationsDepartements.Destinations Is Nothing OrElse referentielDestinationsDepartements.Destinations.Count = 0 Then
                Throw New Exception("Aucune destination chargée")
            End If

            If referentielDestinationsDepartements.Departements Is Nothing OrElse referentielDestinationsDepartements.Departements.Count = 0 Then
                Throw New Exception("Aucun departement chargé")
            End If

            ViderDictionnaires()

            DicoDestinationsRH = referentielDestinationsDepartements.Destinations
            DicoDepartementsRH = referentielDestinationsDepartements.Departements
            DicoEquipesInfoRefRH = referentielDestinationsDepartements.EquipesInfo

            DicoGroupesDiffusionRH("phd") = Commun.TransformeSAMACCOUNTenCN("phd")
            DicoGroupesDiffusionRH("postdoc") = Commun.TransformeSAMACCOUNTenCN("postdoc")

            If users Is Nothing OrElse users.Count = 0 Then
                Throw New Exception("Aucun utilisateur chargé")
            End If

            ' Écrit le fichier final mono-équipe, après normalisation des bâtiments et téléphones.
            SaveUsersToFile("c:\temp\listep.txt", users)

            result = True

        Catch ex As Exception
            Commun.Journal("ERREUR : Creation des fichiers et des tableaux : " & ex.Message, True)
            Commun.SendEmail(AdminScriptLogin & "@igbmc.fr", "steph@igbmc.fr", "ModifAuto.NET : Rapport d'erreur", Commun.journalECHECMail)
            Return result
        End Try

        Return result
    End Function

    Private Function LoadReferentielDestinationsDepartementsFromFile(path As String) As ReferentielDestinationsDepartements
        Dim destinations As New Dictionary(Of String, DestinationInfo)
        Dim departements As New Dictionary(Of String, DepartementInfo)
        Dim equipesInfo As New Dictionary(Of String, EquipeInfo)

        FileOpen(1, path, OpenMode.Input)
        Try
            While Not EOF(1)
                Dim ligne As String = LineInput(1)

                If String.IsNullOrWhiteSpace(ligne) OrElse InStr(ligne, ",") = 0 Then
                    Continue While
                End If

                Dim temp() As String = Split(ligne, ",")
                If temp.Length < 7 Then
                    Continue While
                End If

                Dim nomCourtDest As String = temp(0).Trim()
                Dim nomLongDest As String = temp(1).Trim()
                Dim idDest As String = temp(2).Trim()
                Dim loginResponsable As String = temp(3).Trim()
                Dim nomCourtDept As String = temp(4).Trim()
                Dim nomLongDept As String = temp(5).Trim()
                Dim idDept As String = temp(6).Trim()

                Dim equipeInfoNomCourt As String = ""
                Dim equipeInfoNomLong As String = ""
                Dim idEquipeInfo As String = ""
                Dim dnDestAd As String = ""
                Dim dnEquipeInfoEqAd As String = ""

                ModEquipeDestinationDepartement.ChargerInfosEquipeDepuisDestinationNomCourt(nomCourtDest,
                                                            equipeInfoNomCourt,
                                                            equipeInfoNomLong,
                                                            idEquipeInfo,
                                                            dnDestAd,
                                                            dnEquipeInfoEqAd)

                If String.IsNullOrWhiteSpace(idEquipeInfo) Then
                    idEquipeInfo = "externe"
                End If

                If Not destinations.ContainsKey(idDest) Then
                    destinations.Add(idDest, New DestinationInfo With {
                        .id_dest = idDest,
                        .nom_court_dest = nomCourtDest,
                        .nom_long_dest = nomLongDest,
                        .id_equipeinfo = idEquipeInfo,
                        .id_dept = idDept,
                        .login_responsable_dest = loginResponsable,
                        .dn_dest_ad = dnDestAd
                    })
                End If

                If Not departements.ContainsKey(idDept) Then
                    departements.Add(idDept, New DepartementInfo With {
                        .id_dept = idDept,
                        .nom_court_dept = nomCourtDept,
                        .nom_long_dept = nomLongDept,
                        .dn_dept_ad = ""
                    })
                End If

                If Not String.IsNullOrWhiteSpace(idEquipeInfo) Then
                    If Not equipesInfo.ContainsKey(idEquipeInfo) Then
                        equipesInfo.Add(idEquipeInfo, New EquipeInfo With {
                            .id_equipeinfo = idEquipeInfo,
                            .nom_court_equipeinfo = equipeInfoNomCourt,
                            .nom_long_equipeinfo = equipeInfoNomLong,
                            .dn_equipeinfo_eq_ad = dnEquipeInfoEqAd
                        })
                    End If
                End If
            End While
        Finally
            FileClose(1)
        End Try

        Dim dicoDnDepartementsAd As Dictionary(Of String, String) = ModEquipeDestinationDepartement.ChargerDnDepartementsDepuisAd()

        For Each kvp As KeyValuePair(Of String, DepartementInfo) In departements
            Dim idDept As String = kvp.Key
            If dicoDnDepartementsAd.ContainsKey(idDept) Then
                kvp.Value.dn_dept_ad = dicoDnDepartementsAd(idDept)
            End If
        Next

        Return New ReferentielDestinationsDepartements With {
            .Destinations = destinations,
            .Departements = departements,
            .EquipesInfo = equipesInfo
        }
    End Function


    ''' <summary>
    ''' Lit le fichier source des utilisateurs et construit un dictionnaire indexé par identifiant utilisateur.
    ''' </summary>
    ''' <param name="filePath">
    ''' Chemin complet du fichier source à lire.
    ''' </param>
    ''' <returns>
    ''' Retourne un dictionnaire contenant un seul <see cref="UserRaw"/> par <c>IDuser</c>.
    ''' </returns>
    ''' <remarks>
    ''' Si un même utilisateur apparaît plusieurs fois, les lignes sont fusionnées via <c>MergeUser</c>
    ''' pour conserver l'équipe principale.
    ''' </remarks>
    Private Function LoadUsersFromFile(filePath As String) As Dictionary(Of String, UserRaw)
        ' Le dictionnaire contient un seul enregistrement final par ID utilisateur.
        ' StringComparer.OrdinalIgnoreCase évite qu'un même ID avec une casse différente soit vu comme 2 utilisateurs.
        Dim users As New Dictionary(Of String, UserRaw)(StringComparer.OrdinalIgnoreCase)

        FileOpen(1, filePath, OpenMode.Input)

        Try
            While Not EOF(1)
                Dim ligneP As String = LineInput(1)

                ' Transforme la ligne CSV en objet UserRaw.
                Dim user As UserRaw = ParseUserRaw(ligneP)
                'If user Is Nothing Then
                '    Commun.Journal("ParseUserRaw = Nothing | ligne = " & ligneP, True)
                '    Continue While
                'End If
                ' Si la ligne est invalide ou incomplète, on l'ignore.
                If user Is Nothing Then
                    Continue While
                End If

                ' Sans IDuser, on ne peut pas identifier ni fusionner l'utilisateur.
                If String.IsNullOrWhiteSpace(user.IDuser) Then
                    Continue While
                End If

                ' Si l'utilisateur existe déjà, cela signifie qu'il apparaît sur plusieurs équipes.
                ' On applique alors la règle métier de fusion pour conserver l'équipe principale.
                If users.ContainsKey(user.IDuser) Then
                    users(user.IDuser) = MergeUser(users(user.IDuser), user)
                Else
                    ' Premier passage pour cet utilisateur : on l'ajoute tel quel.
                    users.Add(user.IDuser, user)
                End If
            End While
        Finally
            FileClose(1)
        End Try

        Return users
    End Function

    ''' <summary>
    ''' Convertit une ligne texte du fichier source en objet <see cref="UserRaw"/>.
    ''' </summary>
    ''' <param name="ligneP">
    ''' Ligne brute issue du fichier <c>listepersoJson.txt</c>.
    ''' </param>
    ''' <returns>
    ''' Retourne un objet <see cref="UserRaw"/> si la ligne contient au moins 18 colonnes ; sinon <c>Nothing</c>.
    ''' </returns>
    ''' <remarks>
    ''' Cette implémentation repose sur un simple <c>Split</c> sur la virgule.
    ''' Elle reste donc sensible aux champs qui contiendraient eux-mêmes des virgules.
    ''' </remarks>
    Private Function ParseUserRaw(ligneP As String) As UserRaw
        If String.IsNullOrWhiteSpace(ligneP) Then
            Return Nothing
        End If

        Dim tabP As String() = Split(ligneP, ","c)

        If tabP Is Nothing OrElse tabP.Length < 18 Then
            Return Nothing
        End If

        Dim user As New UserRaw

        user.lastname = Trim(tabP(0))
        user.firstname = Trim(tabP(1))
        user.login = LCase(Trim(tabP(2)))
        user.Dest_short_name = Trim(tabP(3))
        user.dest_name = Trim(tabP(4))

        Dim tr As Integer = 0
        If Not String.IsNullOrWhiteSpace(tabP(5)) Then
            Integer.TryParse(Split(tabP(5), "."c)(0), tr)
        End If
        user.time_rate = tr

        user.BatimentUser = Trim(tabP(6))
        user.TelUser = Trim(tabP(7))
        user.equipeUser = Trim(tabP(8))
        user.Nplus1ID = Trim(tabP(9))
        user.IDuser = Trim(tabP(10))
        user.aliasMail = Trim(tabP(11))
        user.ld = LCase(Trim(tabP(12)))
        user.organism = Trim(tabP(13))
        user.finContrat = Trim(tabP(14))
        user.genre = Trim(tabP(15))
        user.diffusionPhotoInterne = Trim(tabP(16))
        user.uniteName = Trim(tabP(17))

        If String.IsNullOrWhiteSpace(user.Dest_short_name) Then
            user.destinationsRH = New String() {}
        Else
            user.destinationsRH = New String() {user.Dest_short_name}
        End If

        Return user
    End Function

    ''' <summary>
    ''' Fusionne deux occurrences d'un même utilisateur afin de conserver la ligne métier la plus pertinente.
    ''' </summary>
    ''' <param name="existingUser">
    ''' Utilisateur déjà présent dans le dictionnaire.
    ''' </param>
    ''' <param name="newUser">
    ''' Nouvelle occurrence du même utilisateur lue dans le fichier source.
    ''' </param>
    ''' <returns>
    ''' Retourne l'utilisateur fusionné à conserver dans le dictionnaire.
    ''' </returns>
    ''' <remarks>
    ''' La règle principale consiste à garder la ligne avec le plus grand <c>time_rate</c>.
    ''' Si la ligne prioritaire ne contient pas de <c>Nplus1ID</c>, l'ancien chef est conservé.
    ''' </remarks>
    Private Function MergeUser(existingUser As UserRaw, newUser As UserRaw) As UserRaw
        ' Fusion des destinations avant toute chose, pour ne rien perdre.
        Dim destinations As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each d As String In If(existingUser.destinationsRH, New String() {})
            If Not String.IsNullOrWhiteSpace(d) Then
                destinations.Add(d.Trim())
            End If
        Next

        For Each d As String In If(newUser.destinationsRH, New String() {})
            If Not String.IsNullOrWhiteSpace(d) Then
                destinations.Add(d.Trim())
            End If
        Next

        If Not String.IsNullOrWhiteSpace(existingUser.Dest_short_name) Then
            destinations.Add(existingUser.Dest_short_name.Trim())
        End If

        If Not String.IsNullOrWhiteSpace(newUser.Dest_short_name) Then
            destinations.Add(newUser.Dest_short_name.Trim())
        End If

        ' Règle principale : on conserve la ligne avec le plus grand time_rate.
        If newUser.time_rate > existingUser.time_rate Then
            Dim oldManager As String = existingUser.Nplus1ID

            existingUser.lastname = newUser.lastname
            existingUser.firstname = newUser.firstname
            existingUser.login = newUser.login
            existingUser.Dest_short_name = newUser.Dest_short_name
            existingUser.dest_name = newUser.dest_name
            existingUser.time_rate = newUser.time_rate
            existingUser.BatimentUser = newUser.BatimentUser
            existingUser.TelUser = newUser.TelUser
            existingUser.equipeUser = newUser.equipeUser
            existingUser.IDuser = newUser.IDuser
            existingUser.aliasMail = newUser.aliasMail
            existingUser.ld = newUser.ld
            existingUser.organism = newUser.organism
            existingUser.finContrat = newUser.finContrat
            existingUser.genre = newUser.genre
            existingUser.diffusionPhotoInterne = newUser.diffusionPhotoInterne
            existingUser.uniteName = newUser.uniteName

            If newUser.Nplus1ID <> "" Then
                existingUser.Nplus1ID = newUser.Nplus1ID
            Else
                existingUser.Nplus1ID = oldManager
            End If
        Else
            If existingUser.Nplus1ID = "" AndAlso newUser.Nplus1ID <> "" Then
                existingUser.Nplus1ID = newUser.Nplus1ID
            End If
        End If

        ' On réinjecte les destinations fusionnées à la fin.
        existingUser.destinationsRH = destinations.ToArray()

        Return existingUser
    End Function

    ''' <summary>
    ''' Écrit le fichier final mono-équipe à partir du dictionnaire d'utilisateurs fusionnés.
    ''' </summary>
    ''' <param name="filePath">
    ''' Chemin complet du fichier de sortie à générer.
    ''' </param>
    ''' <param name="users">
    ''' Dictionnaire des utilisateurs à écrire.
    ''' </param>
    ''' <remarks>
    ''' Avant écriture, certains champs sont normalisés, notamment le bâtiment et les numéros de téléphone.
    ''' Le format de sortie reste compatible avec le format historique attendu par le reste de l'application.
    ''' </remarks>
    Private Sub SaveUsersToFile(filePath As String, users As Dictionary(Of String, UserRaw))
        FileOpen(1, filePath, OpenMode.Output)

        Try
            For Each kvp As KeyValuePair(Of String, UserRaw) In users
                Dim user As UserRaw = kvp.Value

                ' Normalise certaines valeurs avant écriture dans le fichier final.
                user.BatimentUser = NormalizeBuilding(user.BatimentUser)
                user.TelUser = NormalizePhoneNumbers(user.TelUser)

                ' On conserve ici le format historique du fichier de sortie.
                PrintLine(1,
                    user.lastname & "," &
                    user.firstname & "," &
                    user.login & "," &
                    user.Dest_short_name & "," &
                    user.dest_name & "," &
                    user.time_rate & "," &
                    user.BatimentUser & "," &
                    user.TelUser & "," &
                    user.equipeUser & "," &
                    user.Nplus1ID & "," &
                    user.IDuser & "," &
                    user.aliasMail)
            Next
        Finally
            FileClose(1)
        End Try
    End Sub

    ''' <summary>
    ''' Remplace certains libellés de bâtiments par une version courte et homogène.
    ''' </summary>
    ''' <param name="building">
    ''' Libellé du bâtiment à normaliser.
    ''' </param>
    ''' <returns>
    ''' Retourne le libellé normalisé.
    ''' </returns>
    ''' <remarks>
    ''' Cette méthode applique uniquement des remplacements textuels simples.
    ''' </remarks>
    Private Function NormalizeBuilding(building As String) As String
        ' Uniformise certains noms de bâtiments pour garder un libellé court et cohérent.
        building = Replace(building, "BATIMENT ICS", "ICS")
        building = Replace(building, "C.E.B.G.S", "CEBGS")
        building = Replace(building, "E.S.B.S.", "ESBS")
        Return building
    End Function

    ''' <summary>
    ''' Normalise une liste de numéros de téléphone séparés par des point-virgules.
    ''' </summary>
    ''' <param name="phoneRaw">
    ''' Chaîne brute contenant zéro, un ou plusieurs numéros de téléphone.
    ''' </param>
    ''' <returns>
    ''' Retourne la chaîne des numéros après normalisation.
    ''' </returns>
    ''' <remarks>
    ''' Les numéros internes sur 4 chiffres sont complétés avec un préfixe externe.
    ''' Les numéros français sur 10 chiffres commençant par 0 sont convertis en format <c>+33</c>.
    ''' Le marqueur <c>----</c> est converti en valeur vide.
    ''' </remarks>
    Private Function NormalizePhoneNumbers(phoneRaw As String) As String
        ' Le fichier peut contenir plusieurs téléphones séparés par des ';'.
        Dim tabTelTemp = Split(phoneRaw, ";")
        Dim result As New List(Of String)

        For Each tel As String In tabTelTemp
            If tel <> "----" Then
                ' Cas d'un numéro interne sur 4 chiffres.
                If Len(tel) = 4 Then
                    Dim prefixTel As String = ""

                    ' Selon le préfixe interne, on reconstitue le numéro externe correspondant.
                    If Strings.Left(tel, 2) = "32" OrElse Strings.Left(tel, 2) = "33" OrElse Strings.Left(tel, 2) = "34" OrElse Strings.Left(tel, 2) = "35" OrElse Strings.Left(tel, 2) = "56" OrElse Strings.Left(tel, 2) = "57" Then
                        prefixTel = "(+3338865) "
                    End If

                    If Strings.Left(tel, 2) = "50" OrElse Strings.Left(tel, 2) = "51" OrElse Strings.Left(tel, 2) = "52" Then
                        prefixTel = "(+3336948) "
                    End If

                    result.Add(prefixTel & tel)

                    ' Cas d'un numéro français sur 10 chiffres commençant par 0.
                ElseIf Len(tel) = 10 AndAlso Strings.Left(tel, 1) = "0" Then
                    ' On remplace le 0 initial par l'indicatif +33.
                    result.Add("+33 " & tel.Substring(1))
                Else
                    ' Sinon on laisse la valeur telle quelle.
                    result.Add(tel)
                End If
            Else
                ' Le marqueur ---- est transformé en champ vide dans le fichier final.
                result.Add("")
            End If
        Next

        ' Recompose la liste des numéros dans le même format que l'entrée.
        Return String.Join(";", result)
    End Function


    Public Sub GestionComptesExternes()
        Commun.Journal("Debut de la gestion des comptes Externes", False)

        Const ADS_UF_ACCOUNT_DISABLE = 2
        Dim i = 0
        Using objADExt As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateursExternes), Commun.admin, Commun.passwd, auth)
            Dim results As SearchResultCollection
            Using searcher As DirectorySearcher = New DirectorySearcher(objADExt)
                searcher.SearchScope = SearchScope.OneLevel
                'searcher.Filter = "(&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)(!(msExchRecipientTypeDetails=1))))"
                searcher.Filter = "(&((&(objectClass=user)((!(msExchRecipientTypeDetails=1)))))(!(sAMAccountName=gardiens)))"
                results = searcher.FindAll
            End Using
            For Each result As SearchResult In results
                Dim login1 As String = result.Properties("sAMAccountName")(0)
                Try

                    Using userExt As DirectoryEntry = result.GetDirectoryEntry
                        Dim login As String = userExt.Properties("sAMAccountName").Value
                        Commun.AddRemoveADGroup(login, "G_Externes", "Add")

                        ''ignorer les utilisateurs dont le contrat a expiré
                        Dim interval As Integer = DateDiff("d", userExt.Properties("accountDeactivationDT").Value, Now.Date)
                        If interval >= 1 Then
                            Continue For
                        End If

                        'si l 'utilisateur appartient au groupe "utilisa. du domaine"
                        If Commun.AppartientGroup(login, "Utilisa. du domaine") = True And Commun.AccountIsDisabled(userExt) = False Then
                            userExt.Properties("description").Value = "Desactivé le " & Strings.Left(CStr(Now), 10) & " : Utilisateurs Appartenant au groupe ""utilisateur du domaine"""
                            Commun.AppliquerChangement(userExt)
                            Commun.ReactiveDesactiveCompte(userExt, "desactive")
                        End If

                        'si une des 2 propriétés n'est pas definie pour l'utilisateur
                        If (Not userExt.Properties.Contains("mail") Or Not userExt.Properties.Contains("proxyaddresses")) Then 'Or Not userExt.Properties.Contains("homeMDB") Then
                            'Si la propriété "mail" est presente, on lance la commande Powershell pour definir l'utilisateur comme "utilisateur avec Messagerie" dans Exchange
                            If userExt.Properties.Contains("mail") Then
                                Dim mail As String = userExt.Properties("mail").Value
                                If Strings.InStr(mail, "@igbmc.fr") = 0 Then
                                    If Pws.commandePWSMailUser(userExt.Properties("samAccountName").Value, mail) = True Then
                                        userExt.Properties("description").Clear()
                                        Commun.AppliquerChangement(userExt)
                                        Commun.ReactiveDesactiveCompte(userExt, "active")
                                        Sleep(20000)
                                        Commun.Journal("GestionComptesExternes : User Externe : Utilisateur de messagerie créé : " & login)
                                    Else
                                        Using AD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
                                            Using searcherMailExist As DirectorySearcher = New DirectorySearcher(AD)

                                                'searcher.Filter = "(&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)(!(msExchRecipientTypeDetails=1))))"
                                                searcherMailExist.Filter = "(&(proxyAddresses=SMTP:" & mail & "))"
                                                Dim resultsMailExist As SearchResult = searcherMailExist.FindOne
                                                If Not result Is Nothing Then
                                                    Dim userWithMail As String = resultsMailExist.Path
                                                    Commun.Journal("ERREUR : GestionComptesExternes: User Externe : La creation de l'utilisateur de messagerie à échouée : " & login & " .L'adresse Externe existe deja sur cet objet : " & userWithMail & " : " & mail, True)
                                                End If
                                            End Using
                                        End Using
                                        Commun.Journal("ERREUR : GestionComptesExternes: User Externe : La creation de l'utilisateur de messagerie à échouée: " & login & " : " & mail, True)
                                    End If

                                Else
                                    Using ADuser As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateursActifs), Commun.admin, Commun.passwd, auth)
                                        Using searcherMail As DirectorySearcher = New DirectorySearcher(ADuser)
                                            searcherMail.SearchScope = SearchScope.OneLevel
                                            searcherMail.Filter = "((proxyAddresses=*" & mail & "))"
                                            Dim result1 As SearchResult = searcherMail.FindOne
                                            If result1 Is Nothing Then
                                                userExt.Properties("description").Value = "Compte externe lié a un utilisateur interne parti"
                                                userExt.Properties("mail").Value = Nothing
                                                Commun.AppliquerChangement(userExt)
                                                Commun.ReactiveDesactiveCompte(userExt, "desactive")
                                                Commun.Journal("GestionComptesExternes : User Externe : Compte externe désacivé lié a une adresse interne inexistante : " & login)
                                            End If
                                        End Using
                                    End Using
                                End If
                                'si la propiété "proxyaddresses" est definie, on recupere l'adresse principale pour definir la propriété "mail"
                            ElseIf userExt.Properties.Contains("proxyAddresses") Then
                                Dim currentAddresses = userExt.Properties("proxyAddresses").Value
                                'Dim adress As New List(Of String)
                                Dim addresseMailIdentifie As String = ""
                                For Each value In currentAddresses
                                    If Strings.Left(value, 5) = "SMTP:" Then
                                        addresseMailIdentifie = Replace(value, "SMTP:", "")
                                        Exit For
                                    End If
                                Next value
                                userExt.Properties("mail").Value = addresseMailIdentifie
                                userExt.Properties("description").Clear()
                                Commun.AppliquerChangement(userExt)
                                Commun.ReactiveDesactiveCompte(userExt, "active")
                                Commun.Journal("GestionComptesExternes : User Externe : Mise a jour de l'adresse mail : " & login)
                            Else
                                'MsgBox(userExt.Properties("displayname").Value)
                                If Commun.AccountIsDisabled(userExt) = False Then
                                    userExt.Properties("description").Value = "Desactivé le " & Strings.Left(CStr(Now), 10) & " : Aucune Adresse mail de contact"
                                    Commun.AppliquerChangement(userExt)
                                    Commun.ReactiveDesactiveCompte(userExt, "desactive")
                                    Commun.Journal("GestionComptesExternes : User Externe : Compte externe désacivé aucune adresse mail de contact : " & login)
                                End If
                            End If

                        ElseIf Commun.AccountIsDisabled(userExt) = True Then
                            userExt.Properties("description").Clear()
                            Commun.AppliquerChangement(userExt)
                            Commun.ReactiveDesactiveCompte(userExt, "active")
                        End If

                        'si l'attribut mail est different de TargetAddress
                        If "SMTP:" & userExt.Properties("mail").Value <> userExt.Properties("TargetAddress").Value Then
                            Dim newMail As String = "SMTP:" & userExt.Properties("mail").Value
                            Dim oldMail As String = userExt.Properties("TargetAddress").Value
                            Dim currentAddresses = userExt.Properties("proxyAddresses").Value


                            Dim index As Integer = Array.IndexOf(currentAddresses, oldMail)
                            currentAddresses(index) = newMail
                            userExt.Properties("proxyAddresses").Value = currentAddresses
                            userExt.Properties("TargetAddress").Value = newMail

                            Commun.AppliquerChangement(userExt)

                        End If
fermerUsing:
                    End Using
                Catch ex As Exception
                    Commun.Journal("ERREUR : GestionComptesExternes: User Externe : " & login1, True)
                End Try
            Next
        End Using


        Commun.Journal("Fin de la gestion des comptes externes")
    End Sub



    Public Function MailCloture(ByVal prenom As String, ByVal dateDeSuppressionPrevue As String, dateFindeContrat As String) As String

        Dim nbrJourRestant As Integer = DateDiff("d", Now, Convert.ToDateTime(dateDeSuppressionPrevue))

        Return "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" &
        "<html>" &
        "<head>" &
        "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" &
        "<title>ARRET DU COMPTE</title>" &
        "</head>" &
        "<body>" &
         vbCrLf & "English version below...<BR><BR>" &
             vbCrLf & vbCrLf & "Bonjour " & prenom & ", <BR><BR>" &
             vbCrLf & vbCrLf & "Votre contrat de travail vous liant à l’IGBMC s’est terminé le " & dateFindeContrat & " à minuit.<BR><BR>" &
             vbCrLf & vbCrLf & "En conformité avec la Politique de Sécurité des Systèmes d’Information de l’IGBMC, <B>votre compte informatique à l’IGBMC a été automatiquement désactivé.</B><BR>" &
             vbCrLf & "Si un nouveau contrat est enregistré dans un délai de " & nbrJourRestant & " jours, votre compte sera automatiquement réactivé.<BR><BR>" &
             vbCrLf & vbCrLf & "Votre boite mail restera accessible pendant une période de " & nbrJourRestant & " jours à compter d’aujourd’hui avant d'être définitivement supprimée.<BR>" &
             vbCrLf & "Vous pouvez notamment vous y connecter au travers du webmail de l’IGBMC : <a href=""https://igbmcmail.igbmc.fr"">https://igbmcmail.igbmc.fr</a><BR>" &
             vbCrLf & "Vous pouvez demander une archive complète de votre boite mail avant sa suppression définitive en envoyant un email à hepldesk@igbmc.fr.<BR><BR>" &
             vbCrLf & vbCrLf & "Vous n'avez à présent plus accès aux ressources informatiques de l’IGBMC (espaces de stockage, ressources de calcul, postes de travail, etc.).<BR>" &
             vbCrLf & "Les données de vos espaces de stokage (Seafile, Mendel, Space2) ont été mises à la disposition de votre chef d'équipe. Si vous n’avez pas pris le temps de les récupérer, merci de le contacter.<BR><BR>" &
             vbCrLf & vbCrLf & "<B>Attention :<BR>" &
             vbCrLf & "<FONT color=""red"">Votre compte informatique sera définitivement supprimé le " & dateDeSuppressionPrevue & ".</FONT><BR>" &
             vbCrLf & "Après cette date, plus aucune récupération de données ou réactivation ne sera possible. Votre compte ainsi que votre boite mail seront définitivement détruits.</B><BR><BR>" &
             vbCrLf & vbCrLf & "Si la clôture de votre compte informatique vous semble anormale," &
             " merci de contacter votre gestionnaire au sein du service des ressources humaines afin de demander la mise à jour de votre dossier.<BR><BR>" &
             vbCrLf & vbCrLf & "Le service Informatique.<BR><BR><BR>" &
             "<hr>" &
             vbCrLf & vbCrLf & vbCrLf & "Dear " & prenom & ", <BR><BR>" &
             vbCrLf & vbCrLf & "Your contract with the IGBMC ended at midnight on " & dateFindeContrat & ".<BR><BR>" &
             vbCrLf & vbCrLf & "In accordance with the IGBMC Information Systems Security Policy, <B>your IT account at the IGBMC has been automatically disabled.</B><BR>" &
             vbCrLf & "If a new contract is registered within " & nbrJourRestant & " days, your account will be automatically reactivated.<BR><BR>" &
             vbCrLf & vbCrLf & "Your mailbox will remain accessible for a period of " & nbrJourRestant & " days starting today before being permanently deleted.<BR>" &
             vbCrLf & "You can connect to it via the IGBMC webmail: <a href=""https://igbmcmail.igbmc.fr"">https://igbmcmail.igbmc.fr</a><BR><BR>" &
             vbCrLf & "You can request a complete archive of your mailbox before it is permanently deleted by sending an e-mail to hepldesk@igbmc.fr.<BR><BR>" &
             vbCrLf & vbCrLf & "You no longer have access to the IT resources of the IGBMC (storage spaces, computing resources, workstations, etc.).<BR>" &
             vbCrLf & "The data in your storage areas (Seafile, Mendel, Space2) have been made available to your team leader. If you have not taken the time to retrieve them, please contact her/him.<BR><BR>" &
             vbCrLf & vbCrLf & "<B>Warning :<BR>" &
             vbCrLf & "<FONT color=""red"">Your computer account will be permanently deleted on " & dateDeSuppressionPrevue & ".</FONT><BR>" &
             vbCrLf & "After this date, no more data recovery or account reactivation will be possible. Your IT account and your mailbox will be permanently destroyed.</B><BR><BR>" &
             vbCrLf & vbCrLf & "If the closure of your computer account appears abnormal," &
             " please contact your manager in the human resources department to request the update of your contract’s details.<BR><BR>" &
             vbCrLf & vbCrLf & "The IT department.<BR>" &
             vbCrLf & "</body>" &
             "</html>"


    End Function

    Public Function UserMembreDeDestination(ByVal username As String) As String()
        Dim appartientA As String()
        Dim Entry As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
        Dim Searcher As New DirectorySearcher(Entry)
        Searcher.SearchScope = DirectoryServices.SearchScope.Subtree
        'Searcher.Filter = "(&(objectcategory=user)(sAMAccountName=" & username & "))"
        Searcher.Filter = "(&(objectcategory=group)(member=" & Commun.TransformeSAMACCOUNTenCN(username) & ")(name=* grp))"
        Dim res As SearchResultCollection = Searcher.FindAll
        For Each grp As SearchResult In res
            Dim grpName As String = grp.Properties("cn")(0).ToString
            appartientA.Add(grpName)
        Next grp
        Searcher.Dispose()
        Searcher = Nothing
        Return appartientA
        Entry.Close()
        Entry.Dispose()
        Entry = Nothing
        appartientA = Nothing
    End Function

    Public Function EquipeComptableFichierUser(ByVal login As String) As String()
        Try
            Dim srJ As StreamReader

            srJ = New StreamReader("c:\temp\listepersoJson.txt", System.Text.Encoding.Default)

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
    Public Function chercheDepartementparRapporDestination(ByVal destination As String) As String
        Try

            Using srJ As StreamReader = New StreamReader("c:\temp\eq.txt", System.Text.Encoding.Default)

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
                Return departement
            End Using

        Catch ex As Exception
            Commun.Journal("ERREUR : Recuperation du departement : " & destination & " : " & ex.Message, True)
        End Try
    End Function
    Public Sub CompararaisonAjoutRetraitDestinationsDepartement(ByVal usrDir As DirectoryEntry)
        Dim login As String = usrDir.Properties("sAMAccountName").Value
        Dim EquipeUserAD As String() = UserMembreDeDestination(login)
        Dim EquipeUserFichier As String() = EquipeComptableFichierUser(login)
        Dim DepartementUserAdSR As SearchResultCollection = Commun.SearchFilterAll(New DirectoryEntry("LDAP://" & Commun.LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth), "(&(objectcategory=group)(member=" & Commun.TransformeSAMACCOUNTenCN(login) & ")(name=Dpt_*)(!(name=Dpt_PI*)))", SearchScope.Subtree, "name")
        Dim DepartementUserAd As String() = {}
        For Each res As SearchResult In DepartementUserAdSR
            DepartementUserAd.Add(res.Properties("name")(0).ToString)
        Next
        Try

            'comparaison des 2 tableaux entre eux

            If EquipeUserAD Is Nothing Then
                EquipeUserAD = {} 'EquipeUserFichier
            End If

            Dim DepartementUserFile As String() = {}
            If EquipeUserFichier Is Nothing Then
                EquipeUserFichier = {} 'EquipeUserAD
            Else
                For Each grp In EquipeUserFichier
                    DepartementUserFile.Add(chercheDepartementparRapporDestination(grp))
                Next
            End If

            Dim addToAD As String() = EquipeUserFichier.Except(EquipeUserAD).ToArray

            Dim remFromAD As String() = EquipeUserAD.Except(EquipeUserFichier).ToArray



            'Ajout de l'utilisateur aux groupes-destinations du fichier json auquels il n'appartient pas dans l'ad
            'L'ajout au departement se fait un peu plus bas
            For Each grp As String In addToAD
                If Commun.AppartientGroup(login, "G_Domain_DisableOpenSession") = False Then 'And Commun.AppartientGroup(login, grp) = False Then
                    Commun.AddRemoveADGroup(login, grp, "Add")
                    Commun.Journal("Succes : CompararaisonAjoutRetraitDestinationsDepartement : " & login & " ajouté dans " & grp)
                End If
            Next

            'Retrait de l'utilisateur des groupes-destinations de l'AD auquels il n'appartient plus d'apres le fichier Json

            For Each grp As String In remFromAD
                'If Commun.AppartientGroup(login, grp) = True Then
                Commun.AddRemoveADGroup(login, grp, "Remove")
                Commun.Journal("Succes : CompararaisonAjoutRetraitDestinationsDepartement : " & login & " retiré de " & grp)
                'End If
            Next

            'Retrait de l'utilisateur des departements de l'AD auquels il n'appartient plus d'apres le fichier Json
            For Each dptGrp In DepartementUserAd
                Dim aaa = Array.IndexOf(DepartementUserFile, dptGrp)
                If aaa = -1 Then
                    Commun.AddRemoveADGroup(login, dptGrp, "Remove")
                    Commun.Journal("Succes : CompararaisonAjoutRetraitDestinationsDepartement : " & login & " retiré de " & dptGrp)
                End If
            Next

            'Ajout de l'utilisateur au groupe du departement dont ses destinations dependent
            'Specifiquement dans le cas d'un nouvel utilisateur qui est ajouté a sa destination au moment de la création de son compte (autocompte)
            'Modifauto ne l'ajoutera pas a la destination donc pas au département
            If Commun.AppartientGroup(login, "G_Domain_DisableOpenSession") = False Then
                For Each dest As String In EquipeUserFichier
                    Dim dptGrp As String = chercheDepartementparRapporDestination(dest)
                    If dptGrp <> "" Then
                        If Array.IndexOf(DepartementUserAd, dptGrp) = -1 Then
                            Commun.AddRemoveADGroup(login, dptGrp, "Add")
                            Commun.Journal("Succes : CompararaisonAjoutRetraitDestinationsDepartement : " & login & " ajouté dans " & dptGrp)
                        End If
                    End If
                Next
            End If

            EquipeUserAD = Nothing
            EquipeUserFichier = Nothing
        Catch ex As Exception
            Commun.Journal("ERREUR : CompararaisonAjoutRetraitDestinationsDepartement : " & login & " : " & ex.Message, True)
        End Try
    End Sub

    Public Sub modifAliasMail(ByVal objuser As DirectoryEntry, ByVal prenom As String, ByVal nom As String)

        Dim ctrlChangement As Boolean = False
        'On met le userID de la fonction Commun.DetermineAliasLibre afin d'optenir directement un alias dispo sans rechercher le userID dans l'historique alias 
        Dim aliasMail As String = Commun.DetermineAliasLibre(prenom, nom, "0")(0)


        Try

            Dim currentAddresses = objuser.Properties("proxyAddresses").Value
            Dim adress As New List(Of String)


            For Each value In currentAddresses
                Dim tempValue = Strings.Left(value, 5) & LCase(Strings.Right(value, Len(value) - 5))
                adress.Add(tempValue)
            Next value
            'Definit tous les format d'adresse possible avec le nouveau nom
            Dim nvlAdresse1 As String = "smtp:" & aliasMail & "@igbmc.fr"
            Dim nvlAdresse2 As String = "SMTP:" & aliasMail & "@igbmc.fr"

            Dim iof1 As Integer = adress.IndexOf(nvlAdresse1)
            Dim iof2 As Integer = adress.IndexOf(nvlAdresse2)

            If iof1 = -1 And iof2 = -1 Then
                adress.Add(nvlAdresse1)
                objuser.Properties("proxyAddresses").Value = adress.ToArray()
                Commun.AppliquerChangement(objuser)
                ctrlChangement = True
            End If

            currentAddresses = Nothing
            adress.Clear()
            adress = Nothing

            'Si il y a eu des changements, enregistrement de l'alias mail dans l'historique
            If ctrlChangement = True Then
                Commun.ajoutAliasFichierHisto(aliasMail, objuser.Properties("employeeID").Value)
            End If

        Catch ex As Exception
            Commun.Journal("ERREUR : Modification alias mail sur AD : " & objuser.Properties("sAMAccountName").Value & " : " & ex.Message, True)
        End Try

    End Sub

    Public Sub AttributionStrategieMDP()
        Commun.Journal("Verification des strategies de mot de passe", False)
        'Cas adminInfo
        Dim tabresults As String()
        Dim tabPoste As String()
        Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("OU=AdmInfo,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
            Dim results As SearchResultCollection = Commun.SearchFilterAll(objAD, "(&(objectClass=user)(!(distinguishedName=CN=userprog,OU=AdmInfo,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr)))", SearchScope.Subtree)
            If Not results Is Nothing Then
                For Each result As SearchResult In results
                    tabresults.Add(Replace(result.Path, "LDAP://" & Commun.LdapServerPrefix(), ""))
                Next
            End If
        End Using

        updateGroupeWithArray("G_SMDPM_Admins", tabresults)

        'cas UsersAdm
        Dim tabresults1 As String()
        Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("OU=UsersAdm,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
            Dim results As SearchResultCollection = Commun.SearchFilterAll(objAD, "(&(objectClass=user)(!(distinguishedName=CN=userprog,OU=AdmInfo,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr)))", SearchScope.Subtree)
            If Not results Is Nothing Then
                For Each result As SearchResult In results
                    tabresults1.Add(Replace(result.Path, "LDAP://" & Commun.LdapServerPrefix(), ""))
                Next
            End If
        End Using


        updateGroupeWithArray("G_SMDPM_UsersAdm", tabresults1)

        'cas AdminsPostes
        Dim tabresults3 As String()
        Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("OU=AdminsPostes,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
            Dim results As SearchResultCollection = Commun.SearchFilterAll(objAD, "(&(objectClass=user)(!(distinguishedName=CN=userprog,OU=AdmInfo,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr)))", SearchScope.Subtree)
            If Not results Is Nothing Then
                For Each result As SearchResult In results
                    tabresults3.Add(Replace(result.Path, "LDAP://" & Commun.LdapServerPrefix(), ""))
                Next
            End If
        End Using


        updateGroupeWithArray("G_SMDPM_AdminsPostes", tabresults3)

        'cas Users-Admins
        Dim tabresults2 As String()
        Using objGroupAdminPoste As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("OU=Admins,OU=Postes,OU=Micro,OU=Groupes,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
            Dim results As SearchResultCollection = Commun.SearchFilterAll(objGroupAdminPoste, "(&(objectCategory=group)(member=*))", SearchScope.Subtree)
            Dim dejaMembreStrategie As String() = Commun.MembresDuGroupe("G_SMDPM_Users-Admins")
            If Not results Is Nothing Then
                For Each result As SearchResult In results
                    Dim nomDuGroup As String = Commun.TransformeSAMACCOUNTenCN(Replace(result.Path, "LDAP://" & Commun.LdapServerPrefix(), ""))
                    Dim membresGroup As String() = Commun.MembresDuGroupe(nomDuGroup)
                    If membresGroup Is Nothing Then Continue For
                    Dim nomDuPoste As String = Replace((Replace(nomDuGroup, "DL_", "")), "_Admins", "")
                    Dim posteAjout As Boolean = 0
                    For Each user In membresGroup
                        If user Like "CN=*,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr" And user <> "CN=Stephane CERDAN,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr" Then
                            If posteAjout = False Then
                                tabPoste.Add(Commun.TransformeSAMACCOUNTenCN(nomDuPoste & "$"))
                                posteAjout = True
                            End If
                            If Array.IndexOf(dejaMembreStrategie, user) = -1 Then
                                DisablePasswordNeverExpiresETPasswordLastSet(user)
                            End If
                            tabresults2.Add(user)
                        End If
                    Next
                Next

            End If
        End Using

        'Mise en place de la strategie UAC, pour les postes par groupe
        Using groupePoste As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(Commun.TransformeSAMACCOUNTenCN("PostesAvecAdmin")), Commun.admin, Commun.passwd, auth)
            groupePoste.Properties("member").Value = tabPoste
            Commun.AppliquerChangement(groupePoste)
        End Using


        updateGroupeWithArray("G_SMDPM_Users-Admins", tabresults2)


        Commun.Journal("Verification des strategies de mot de passe terminée", False)

    End Sub
    Public Sub DisablePasswordNeverExpiresETPasswordLastSet(ByVal usrPath As String)
        Using user As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(usrPath), Commun.admin, Commun.passwd, auth)
            Const NON_EXPIRE_FLAG = &H10000
            Dim Val As Integer = user.Properties("userAccountControl").Value
            'Defini le flag ADS_UF_DONT_EXPIRE_PASSWD sur "non coché"
            user.Properties("userAccountControl").Value = Val And Not NON_EXPIRE_FLAG
            Commun.AppliquerChangement(user)
            'Defini le dernier changement de mot de passe a aujourd'hui
            user.Properties("pwdLastSet").Value = 0
            Commun.AppliquerChangement(user)
            user.Properties("pwdLastSet").Value = -1
            Commun.AppliquerChangement(user)
        End Using
    End Sub

    Public Sub updateGroupeWithArray(ByVal nomGroupe As String, ByVal arr As String())
        Dim tabGroup As String() = Commun.MembresDuGroupe(nomGroupe)
        Dim absentDeArr As String() = tabGroup
        Dim absentDeTabGroup1 As String() = arr
        If Not tabGroup Is Nothing And Not arr Is Nothing Then
            absentDeArr = tabGroup.Except(arr).ToArray()
            absentDeTabGroup1 = arr.Except(tabGroup).ToArray()
        End If

        If Not absentDeTabGroup1 Is Nothing Then
            For Each user In absentDeTabGroup1
                Commun.AddRemoveADGroup(user, nomGroupe, "Add")
            Next
        End If
        If Not absentDeArr Is Nothing Then
            For Each user In absentDeArr
                Commun.AddRemoveADGroup(user, nomGroupe, "Remove")
            Next
        End If
    End Sub
    Public Sub ExpirationMDP()
        Dim samaccountname As String
        Try

            Using OUAdmins As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
                Using OUAdminsearcher As DirectorySearcher = New DirectorySearcher(OUAdmins)
                    'OUAdminsearcher.Filter = "(&(objectClass=user)(|((memberof=" & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_Admins") & ")(memberof=" & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_Users-Admins") & "))))"
                    OUAdminsearcher.Filter = "(&(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)" _
                                            & "(|((memberof=" & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_Admins") & ")" _
                                            & "(memberof = " & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_Users-Admins") & ")" _
                                            & "(memberof=" & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_UsersAdm") & ")" _
                                            & "(memberof=" & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_CS") & "))))"
                    OUAdminsearcher.PropertiesToLoad.Add("pwdLastSet")
                    OUAdminsearcher.PropertiesToLoad.Add("userPrincipalName")
                    OUAdminsearcher.PropertiesToLoad.Add("Description")
                    OUAdminsearcher.PropertiesToLoad.Add("mail")
                    OUAdminsearcher.PropertiesToLoad.Add("wWWHomePage")
                    Dim results As SearchResultCollection = OUAdminsearcher.FindAll()
                    For Each resultUser As SearchResult In results
                        samaccountname = Replace(Replace(resultUser.Properties("userPrincipalName")(0), "@igbmc.fr", ""), "@igbmc.u-strasbg.fr", "")
                        'si le compte est userprog continuer sans traiter
                        If samaccountname = "userprog" Then Continue For

                        Dim rappelComplexite As String = ""
                        Dim type As String = ""
                        Dim sujet As String
                        Dim groupe As String = ""
                        Dim StrategieMDP As DirectoryEntry
                        If Commun.AppartientGroup(samaccountname, "G_SMDPM_Admins") = True Then
                            StrategieMDP = New DirectoryEntry("LDAP://" & Commun.LdapPath("CN=Strategie_MDPM_Admins,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
                            type = "adminInfo"
                            sujet = "Expiration de Votre mot de passe d'administration"
                            groupe = "G_SMDPM_Admins"
                        ElseIf Commun.AppartientGroup(samaccountname, "G_SMDPM_UsersAdm") = True Then
                            StrategieMDP = New DirectoryEntry("LDAP://" & Commun.LdapPath("CN=Strategie_MDPM_UsersAdm,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
                            type = "usersAdm"
                            sujet = "Expiration de Votre mot de passe d'administration"
                            groupe = "G_SMDPM_UsersAdm"
                        ElseIf Commun.AppartientGroup(samaccountname, "G_SMDPM_Users-Admins") = True Then
                            StrategieMDP = New DirectoryEntry("LDAP://" & Commun.LdapPath("CN=Strategie_MDPM_Users-Admins,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
                            type = "users-admin"
                            sujet = "Expiration de Votre mot de passe"
                            groupe = "G_SMDPM_Users-Admins"
                        ElseIf Commun.AppartientGroup(samaccountname, "G_SMDPM_AdminsPostes") = True Then
                            StrategieMDP = New DirectoryEntry("LDAP://" & Commun.LdapPath("CN=Strategie_MDPM_AdminsPostes,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
                            type = "adminsPostes"
                            sujet = "Expiration de Votre mot de passe d'administration"
                            groupe = "G_SMDPM_AdminsPostes"
                        ElseIf Commun.AppartientGroup(samaccountname, "G_SMDPM_CS") = True Then
                            StrategieMDP = New DirectoryEntry("LDAP://" & Commun.LdapPath("CN=Strategie_MDMP_CompteService,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
                            type = "Compte de service"
                            sujet = "Expiration du mot de passe de votre compte de service " & samaccountname
                            groupe = "G_SMDPM_CS"
                        Else
                            Commun.Journal("ERREUR : Determiner strategie de mot de passe : " & samaccountname, True)
                            GoTo sortie
                        End If

                        Dim nbrChar As String = StrategieMDP.Properties("msDS-MinimumPasswordLength").Value.ToString
                        Dim historyLenght As String = StrategieMDP.Properties("msDS-PasswordHistoryLength").Value.ToString
                        Dim MaxPWDageJourPolicy As Integer = ConvertAttribute(StrategieMDP.Properties("msDS-MaximumPasswordAge").Value)


                        StrategieMDP.Close()
                        StrategieMDP.Dispose()
                        StrategieMDP = Nothing

                        rappelComplexite = "Votre mot de passe doit etre changé tous les " & MaxPWDageJourPolicy & " jours." & vbCrLf & "Il ne peut pas etre le meme que les " & historyLenght & " précédents." & vbCrLf & "Votre mot de passe doit contenir au moins :" & vbCrLf & vbTab & "- " & nbrChar & " caractères"
                        rappelComplexite = rappelComplexite & vbCrLf & vbTab & "- 1 Majuscule" & vbCrLf & vbTab & "- 1 Minuscule" & vbCrLf & vbTab & "- 1 Chiffre" & vbCrLf & vbTab & "- 1 Caractère spécial (non-alphabétique)"

                        Dim lastSetPWD As DateTime = Format(New DateTime(1601, 1, 2).AddTicks(resultUser.Properties("pwdLastSet")(0)), "dd/MM/yyyy")
                        Dim expirationPWDDate As DateTime = lastSetPWD.AddDays(MaxPWDageJourPolicy - 1)

                        Using userADM As DirectoryEntry = resultUser.GetDirectoryEntry
                            Dim expirationPWDDateTxt As String = expirationPWDDate.ToString("dd/MM/yyy")
                            If type = "usersAdm" Or type = "adminInfo" Then
                                If userADM.Properties("physicalDeliveryOfficeName").Value <> "Expire le : " & expirationPWDDateTxt And resultUser.Properties("pwdLastSet")(0) <> 0 Then
                                    userADM.Properties("physicalDeliveryOfficeName").Value = "Expire le : " & expirationPWDDateTxt
                                    Commun.AppliquerChangement(userADM)
                                End If
                            End If
                        End Using

                        Dim aujourdhui As DateTime = Format(Date.Now.AddDays(0), "dd/MM/yyyy")
                        Dim demain As DateTime = Format(Date.Now.AddDays(1), "dd/MM/yyyy")
                        Dim expiration30 As DateTime = Format(Date.Now.AddDays(30), "dd/MM/yyyy")
                        Dim expiration7 As DateTime = Format(Date.Now.AddDays(7), "dd/MM/yyyy")
                        Dim expiration3 As DateTime = Format(Date.Now.AddDays(3), "dd/MM/yyyy")
                        Dim expiration2 As DateTime = Format(Date.Now.AddDays(2), "dd/MM/yyyy")

                        Dim Email As String

                        If type = "adminInfo" Or type = "usersAdm" Then
                            Email = Strings.Left(samaccountname, Len(samaccountname) - 3) & "@igbmc.fr"
                        ElseIf type = "Compte de service" Then
                            Email = resultUser.Properties("wWWHomePage")(0)
                        Else
                            Email = resultUser.Properties("mail")(0)
                        End If



                        Dim ctrl As Boolean = False
                        If type = "adminInfo" Then
                            ctrl = (aujourdhui = expirationPWDDate Or demain = expirationPWDDate Or expiration7 = expirationPWDDate Or expiration3 = expirationPWDDate Or expiration2 = expirationPWDDate)
                        ElseIf type = "Compte de service" Then
                            ctrl = (aujourdhui = expirationPWDDate Or demain = expirationPWDDate Or expiration30 = expirationPWDDate Or expiration7 = expirationPWDDate Or expiration3 = expirationPWDDate Or expiration2 = expirationPWDDate Or Now() > expirationPWDDate)
                        Else
                            ctrl = (aujourdhui = expirationPWDDate Or demain = expirationPWDDate Or expiration30 = expirationPWDDate Or expiration7 = expirationPWDDate Or expiration3 = expirationPWDDate Or expiration2 = expirationPWDDate)
                        End If

                        If ctrl = True Then
                            Dim corpMail As String = MailTemplatePasswdExpire(type, samaccountname, expirationPWDDate, MaxPWDageJourPolicy, historyLenght, nbrChar)
                            Commun.SendEmail("serviceinfo@igbmc.fr", Email & ";Cc:serviceinfo@igbmc.fr", sujet, corpMail)
                            Commun.Journal("Mot de passe ADM : mail envoyé a : " & Email, False)
                        End If
                    Next resultUser

sortie:
                End Using
            End Using
            Commun.Journal("Gestion des mots de passe des comptes ADM terminée avec succes", False)
        Catch ex As Exception
            Commun.Journal("ERREUR : Gestion de password des comptes ADM : " & samaccountname & " : " & ex.Message, True)
        End Try

    End Sub
    Public Function MailTemplatePasswdExpire(ByVal type As String, ByVal samaccountname As String, ByVal expirationPWD As String, ByVal MaxPWDageJourPolicy As String, ByVal historyLenght As String, ByVal nbrChar As String) As String
        Dim corpMail As String
        If type = "adminInfo" Or type = "usersAdm" Or type = "adminsPostes" Then
            Dim rappelComplexite As String = "Votre mot de passe doit etre changé tous les " & MaxPWDageJourPolicy & " jours." & vbCrLf & "Il ne peut pas etre le meme que les " & historyLenght & " précédents." & vbCrLf & "Votre mot de passe doit contenir au moins :" & vbCrLf & vbTab & "- " & nbrChar & " caractères"
            rappelComplexite = rappelComplexite & vbCrLf & vbTab & "- 1 Majuscule" & vbCrLf & vbTab & "- 1 Minuscule" & vbCrLf & vbTab & "- 1 Chiffre" & vbCrLf & vbTab & "- 1 Caractère spécial (non-alphabétique)"

            corpMail = "Le mot de passe de votre compte administrateur (" & samaccountname & ") va expirer le " & expirationPWD & "." & vbCrLf _
                                                        & "Pensez à le changer en ouvrant une session sur un ordinateur du domaine ou en vous connectant ici : https://password.igbmc.fr" & vbCrLf & vbCrLf _
                                                        & rappelComplexite & vbCrLf & vbCrLf & vbCrLf & "Le service Informatique" & vbCrLf & "(Email généré automatiquement)"
        End If

        If type = "Compte de service" Then
            Dim rappelComplexite As String = "Le mot de passe de votre compte de service doit etre changé tous les " & MaxPWDageJourPolicy & " jours." & vbCrLf & "Il ne peut pas etre le meme que les " & historyLenght & " précédents." & vbCrLf & "Votre mot de passe doit contenir au moins :" & vbCrLf & vbTab & "- " & nbrChar & " caractères"
            rappelComplexite = rappelComplexite & vbCrLf & vbTab & "- 1 Majuscule" & vbCrLf & vbTab & "- 1 Minuscule" & vbCrLf & vbTab & "- 1 Chiffre" & vbCrLf & vbTab & "- 1 Caractère spécial (non-alphabétique)"

            corpMail = "Le mot de passe de votre compte de service (" & samaccountname & ") va expirer ou a expiré le " & expirationPWD & "." & vbCrLf _
                                                        & "Pensez à le changer en vous connectant ici : https://password.igbmc.fr," & vbCrLf & vbCrLf _
                                                        & "ainsi que dans l'application dans laquelle il est utilisé" & vbCrLf & vbCrLf _
                                                        & rappelComplexite & vbCrLf & vbCrLf & vbCrLf & "Le service Informatique" & vbCrLf & "(Email généré automatiquement)"
        End If

        If type = "users-admin" Then

            corpMail = "Le mot de passe de votre compte informatique IGBMC (" & samaccountname & ") va expirer le " & expirationPWD & "." & vbCrLf _
                                    & "Votre mot de passe expire tous les " & MaxPWDageJourPolicy & " jours car votre compte informatique dispose des droits d'administration sur un ou plusieurs postes de travail de l'IGBMC." & vbCrLf & vbCrLf _
                                    & "Pour éviter toute perturbation dans l'accès à votre ordinateur ainsi qu'aux applications de l'IGBMC, nous vous recommandons de le changer AVANT qu'il ne soit expiré : " & vbCrLf & vbCrLf _
                                    & "Si vous utilisez un ordinateur portable, assurez que celui-ci soit connecté au réseau interne ou au VPN de l'IGBMC AVANT de changer votre mot de passe." & vbCrLf & vbCrLf _
                                    & vbTab & "- Si vous utilisez un ordinateur sous Windows, rendez-vous dans le menu Windows, choisissez ""Sécurité de Windows"" puis cliquez sur ""Modifier un mot de passe""" & vbCrLf _
                                    & vbTab & "- Si vous utilisez un Mac, rendez-vous dans les Préférences Systèmes, choisissez ""Utilisateurs et groupes"" puis cliquez sur ""Modifier le mot de passe...""" & vbCrLf & vbCrLf _
                                    & "Votre mot de passe doit comporter au moins " & nbrChar & " caractères dont au moins 1 majuscule, 1 minuscule, 1 chiffre et un caractère spécial (hors & , et |)" & vbCrLf & vbCrLf _
                                    & "Une fois votre mot de passe changé, vous serez peut-être amené à mettre à jour votre mot de passe dans certaines de vos applications, comme votre client email par exemple." & vbCrLf & vbCrLf _
                                    & "Si vous n'avez pas pu changer votre mot de passe avant son expiration, vous pouvez le mettre à jour en vous connectant au webmail de l'IGBMC depuis une simple connexion Internet (https://igbmcmail.igbmc.fr) ou en prenant contact avec le support informatique (03 88 65 35 53)." & vbCrLf & vbCrLf _
                                    & "###" & vbCrLf & vbCrLf _
                                    & "Your IGBMC computer account password (" & samaccountname & ") will expire on the " & expirationPWD & "." & vbCrLf _
                                    & "Your password expires every " & MaxPWDageJourPolicy & " days because your computer account has administrative rights on one Or more IGBMC workstations." & vbCrLf & vbCrLf _
                                    & "To avoid any disruption in the access to your computer as well as to the IGBMC applications, we recommend you to change it BEFORE it expires:" & vbCrLf & vbCrLf _
                                    & "If you are using a laptop, make sure it is connected to the internal network or IGBMC VPN BEFORE you change your password." & vbCrLf & vbCrLf _
                                    & vbTab & "- If you are using a Windows computer, go to the Windows menu, choose ""Windows Security"" and click on ""Change Password""" & vbCrLf _
                                    & vbTab & "- If you are using a Mac, go to System Preferences, choose ""Users And Groups"" And click ""Change Password...""" & vbCrLf & vbCrLf _
                                    & "Your password must be at least " & nbrChar & " characters including at least 1 upper case, 1 lower case, 1 number and one special character (excluding & , and |)." & vbCrLf & vbCrLf _
                                    & "Once you have changed your password, you may need to update your password in some of your applications, such as your email client for example." & vbCrLf & vbCrLf _
                                    & "If you have not been able to change your password before it expires, you can update it by connecting to the IGBMC webmail from a simple Internet connection (https://igbmcmail.igbmc.fr) or by contacting the IT support (03 88 65 35 53)."
        End If


        Return corpMail
    End Function
    '    
    Public Function ConvertAttribute(ByVal li As Object) As Integer
        Try
            Dim lngHigh = li.HighPart
            Dim lngLow = li.LowPart
            Dim lastLogon = (lngHigh * 2 ^ 32) - lngLow
            'Dim returnDateTime As DateTime = DateTime.FromFileTime(lastLogon)
            Dim interval As TimeSpan = New TimeSpan(lastLogon * -1)
            Dim returnDateTime = interval.Days

            Return returnDateTime
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Controle qu'un utilisateur est en exception valide
    ''' </summary>
    ''' <param name="login">Login de l'utilisateur a controler.</param>
    ''' <remarks>Retourne "False" s'il n'est pas en exception, et la date (au format texte) s'il est en exception</remarks>

    Public Function IndexOfMulti(ByVal tab As String(,), ByVal recherche As String, ByVal colonne As Integer) As Integer
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


    Public Function ShellSort(ByVal tab1 As String(),
                        Optional ByVal loBound As Long = -1,
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

    Public Sub creationFichierParJson(DestinationsJsonIsTrue)
        FileOpen(2, "c:\temp\eq.txt", OpenMode.Output)
        'Dim data As String = Json.SendJson("", "destinations?is_team=true", "AD", "GET")
        Dim tabResultJson As String()
        Dim lineJson As String = ""
        Dim usersDesactivations As System.Collections.Generic.HashSet(Of String) = ChargerUsersDesactivations()

        Dim max As Integer = UBound(DestinationsJsonIsTrue)
        For d = 0 To max
            Commun.AfficherBarre("récupération des equipes sur IGBMCServices", d, max)
            Dim IDdest As String = DestinationsJsonIsTrue(d).ID

            'If IDdest <> "189" Then Continue For

            Dim Dest_short_name As String = DestinationsJsonIsTrue(d).short_name
            Dim dest_name As String = DestinationsJsonIsTrue(d).name
            Dim department_id As String = DestinationsJsonIsTrue(d).department_id
            Dim nbrUsersDest As Integer = 0
            Dim team_id As String = DestinationsJsonIsTrue(d).team_group
            Dim group_id As String = DestinationsJsonIsTrue(d).group_id
            Dim entity_id As String = DestinationsJsonIsTrue(d).entity_id
            Dim leaderLoginDest As String = ""

            Dim equipeUser As String = ""
            equipeUser = Commun.RecupEquipeinfo(Dest_short_name)
            If equipeUser = "" Then
                equipeUser = "externe"
            End If




            Dim data As String = Json.SendJson("", "persons?present=true&extern=false&current_destination=true&destination=" & IDdest, "AD", "GET")
            Dim persons = Json.DeserializeJson(data, "persons")

            nbrUsersDest = UBound(persons) + 1



            'Creation du fichier des destinations
            If nbrUsersDest > 0 Then
                Dim destinationsJsonLeaders As String = Json.SendJson("", "destinations/" & IDdest & "/leaders", "AD", "GET")

                'Recherche du responsable de la destination dans la BDP avec recursivité
                leaderLoginDest = LeaderDest(entity_id, group_id, department_id, team_id, IDdest)

                'Ajout des departements au fichier des équipes
                Dim dataDepartments As String = Json.SendJson("", "departments/" & department_id, "AD", "GET")
                'deserialisation d'un objet Json "root"
                Dim departmentsJson As Json.departmentC = JsonConvert.DeserializeObject(Of Json.departmentC)(dataDepartments.ToString)
                Dim depart_Short_name As String = departmentsJson.short_name
                Dim depart_name As String = departmentsJson.name
                Dim depart_ID As String = departmentsJson.id

                PrintLine(2, Dest_short_name & "," & dest_name & "," & IDdest & "," & leaderLoginDest & "," & depart_Short_name & "," & depart_name & "," & depart_ID)

            End If

            For p = 0 To UBound(persons)
                Dim lastname As String = persons(p).lastname
                Dim firstname As String = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(LCase(persons(p).firstname))
                Dim login As String = persons(p).login
                Dim aliasMail As String = persons(p).email_alias
                Dim IDuser As String = persons(p).ID
                If usersDesactivations.Contains(IDuser.Trim()) Then
                    Continue For
                End If

                Dim dateEntree As Date = persons(p).entrance_date
                Dim dataContract As String = Json.SendJson("", "persons/" & IDuser & "/contracts?current_contract=true", "AD", "GET")

                Dim contracts = Json.DeserializeJson(dataContract, "contracts")
                Dim contractRunning As Boolean = ContratEnCours(contracts)
                Dim finContrat As String = DateDeFinDeContract(contracts, IDuser)

                'Dim location As Json.locationC = JsonConvert.DeserializeObject(Of Json.locationC)(locations(l).ToString)
                Dim unite As Json.uniteC = JsonConvert.DeserializeObject(Of Json.uniteC)(persons(p).unite.ToString())
                Dim uniteNametmp As String = unite.nom
                Dim uniteName As String

                Select Case uniteNametmp
                    Case "IGBMC RECHERCHE"
                        uniteName = "UMR 7104"

                    Case "PALME"
                        uniteName = "UAR 2060"

                    Case "BIOSTRUCTURE"
                        uniteName = "UAR 2061"

                    Case "PHEN-ICS"
                        uniteName = "UAR 2062"

                    Case Else
                        uniteName = ""
                End Select

                'Dim unité As String = persons(p).email_alias

                Dim diffusionPhotoInterne As Boolean = persons(p).allow_picture_internal_broadcast 'true/false
                Dim genre As String
                If persons(p).gender = "0" Then
                    genre = "Homme"
                Else
                    genre = "Femme"
                End If

                'recuperation de l'appartenence a la liste phd ou postdoc
                Dim ld As String = ListeDiffusion(IDuser)

                'ou que le compte n'existe pas dans l'AD
                Dim dateMini As Date = #1/1/2015#

                'Verification si l'utilsateur a un compte dans l'AD
                Dim compteAD As Boolean = True
                Dim CN As String = Commun.TransformeSAMACCOUNTenCN(login)
                If CN = "" Then
                    compteAD = False
                End If

                'Si l'utilisateur n'a pas de compte dans l'AD et qu'il est entré apres le 1er janvier 2015
                If compteAD = False And (dateMini < dateEntree And contractRunning = True) Then 'Or testException = True)
                    Dim crea As Creation = New Creation

                    'si le login n'est pas defini
                    If login = "" Then
                        login = crea.DetermineLogin(firstname, lastname, IDuser)
                    End If

                    crea.createCompte(lastname, firstname, Dest_short_name, IDuser, login, genre, finContrat, ld)

                    'quand le compte AD est créé, on defini la variable "compteAD" sur True
                    compteAD = True

                    'ecriture du fichier de creation de compte
                    Dim newUser As String = lastname & "," & firstname & "," & login & ",," & Dest_short_name & ",,,," & ld & "," & IDuser & "," & aliasMail & "," & genre
                    Dim sw As New StreamWriter(ini.ReadValue("MODIFAUTO", "CheminPartage") & "\todo\c" & Now.ToString("dd") & "-" & Now.ToString("MM") & "-" & Now.ToString("yy") & "-" & Now.ToString("HH") & "h.txt", True)
                    sw.WriteLine(newUser)
                    sw.Close()
                    sw.Dispose()
                End If

                'Si l'utilisateur a un contract en cours
                If contractRunning = True And compteAD = True Then
                    Dim locations As JArray = persons(p).locations
                    Dim loc_building_roomT(locations.Count - 1) As String
                    Dim loc_phoneT(locations.Count - 1) As String
                    For l = 0 To locations.Count - 1
                        'deserialisation d'un objet Json "root"
                        Dim location As Json.locationC = JsonConvert.DeserializeObject(Of Json.locationC)(locations(l).ToString)
                        loc_building_roomT(l) = Replace(location.building & " " & Strings.Left(location.room, InStr(location.room, " ") - 1), ",", ";")
                        loc_phoneT(l) = Replace(Replace(Replace(location.phone, " ", ""), ".", ""), ",", ";")
                    Next l
                    Dim BatimentUser As String = Join(loc_building_roomT, ";")

                    Dim TelUser As String = Join(loc_phoneT, ";")
                    Erase loc_phoneT, loc_building_roomT

                    Dim organism As String = Organisme(contracts)
                    Dim Nplus1ID As String = ChercherNPlus1deDestination(IDuser, IDdest)



                    '             0                 1               2                   3                   4               5               6                   7               8                  9              10              11              12          13                  14              15                    16                              17
                    lineJson = lastname & "," & firstname & "," & login & "," & Dest_short_name & "," & dest_name & "," & "100" & "," & BatimentUser & "," & TelUser & "," & equipeUser & "," & Nplus1ID & "," & IDuser & "," & aliasMail & "," & ld & "," & organism & "," & finContrat & "," & genre & "," & diffusionPhotoInterne.ToString & "," & uniteName
                    tabResultJson.Add(lineJson)
                End If
            Next p

        Next d
        FileClose(2)

        'inserer les externes dans le tableau

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


    End Sub
    Private Function ChargerUsersDesactivations() As System.Collections.Generic.HashSet(Of String)
        Dim resultat As New System.Collections.Generic.HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim cheminFichier As String = "\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\UsersDesactivations.txt"

        Try
            If Not File.Exists(cheminFichier) Then
                Commun.Journal("ATTENTION : fichier UsersDesactivations.txt introuvable : " & cheminFichier, True)
                Return resultat
            End If

            Using reader As New StreamReader(cheminFichier, System.Text.Encoding.Default)
                Do While Not reader.EndOfStream
                    Dim ligne As String = reader.ReadLine()
                    If ligne Is Nothing Then Continue Do

                    ligne = ligne.Trim()
                    If ligne = "" OrElse ligne.StartsWith("#") OrElse ligne.StartsWith(";") Then Continue Do

                    resultat.Add(ligne)
                Loop
            End Using

        Catch ex As Exception
            Commun.Journal("ERREUR : lecture UsersDesactivations.txt : " & ex.Message, True)
        End Try

        Return resultat
    End Function
    Public Sub exceptionCreationCompte()
        Dim lines() As String = IO.File.ReadAllLines("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\ForceCreationDeCompte.txt")
        For Each line As String In lines

            Dim dest_short_name As String = Split(line, ",")(1)

            Dim dataException As String = Json.SendJson("", "persons/" & Split(line, ",")(0), "AD", "GET")
            Dim persons = JsonConvert.DeserializeObject(Of Json.personC)(dataException)
            'si le login n'est pas defini
            Dim crea As New Creation

            Dim lastname As String = persons.lastname
            Dim firstname As String = persons.firstname
            Dim login As String = persons.login
            Dim IDuser As String = persons.id
            Dim genre As String
            If persons.gender = "0" Then
                genre = "Homme"
            Else
                genre = "Femme"
            End If
            Dim finContrat As String

            If login = "" Then
                login = crea.DetermineLogin(firstname, lastname, IDuser)
            End If
            crea.createCompte(lastname, firstname, dest_short_name, IDuser, login, genre, finContrat, "")

        Next
        'nettoyage du fichier ForceCreationDeCompte.txt
        Try
            System.IO.File.WriteAllText("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\ForceCreationDeCompte.txt", "")
        Catch
            Commun.Journal("ERREUR : nettoyage du fichier ForceCreationDeCompte.txt", True)
        End Try
    End Sub
    Public Function ConvertOrganism(ByVal orgID As String) As String
        Dim result As String = ""

        Dim orgInteger As Integer = Convert.ToInt32(orgID)
        Select Case orgInteger
            Case 2
                result = "CNRS"
            Case 3
                result = "INSERM"
            Case 4
                result = "ADEREGEM"
            Case 5
                result = "UNISTRA"
            Case 6
                result = "ARC"
            Case 7
                result = "LIGUE"
            Case 8
                result = "BMS"
            Case 9
                result = "COLLEGE"
            Case 11
                result = "UNISTRA-LAB"
            Case 12
                result = "GIE"
            Case 13
                result = "RECHERCHE"
            Case 14
                result = "HUS"
            Case 15
                result = "DIVERS"
        End Select

        Return result

    End Function

    Public Sub ChangePasswordAccountPrestaImagerie()
        Try
            Dim passwordPrestaImagerieAdm = RandomPassword.Generate(8)
            Dim passwordPrestaImagerieUsr = RandomPassword.Generate(8)
            Using userEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(Commun.TransformeSAMACCOUNTenCN("cs-prestaadm")), Commun.admin, Commun.passwd, auth)
                userEntry.Invoke("SetPassword", New Object() {passwordPrestaImagerieAdm})
                Commun.AppliquerChangement(userEntry)
            End Using
            Using userEntry1 = New DirectoryEntry("LDAP://" & Commun.LdapPath(Commun.TransformeSAMACCOUNTenCN("cs-prestausr")), Commun.admin, Commun.passwd, auth)
                userEntry1.Invoke("SetPassword", New Object() {passwordPrestaImagerieUsr})
                Commun.AppliquerChangement(userEntry1)
            End Using
            Dim textMail = "Bonjour," & vbCrLf & vbCrLf & "Les mots de passe pour les prestataires externes ont changés." & vbCrLf & vbCrLf & "Le mot de passe pour le compte ""cs-prestaadm"" est : " & vbTab & vbTab & passwordPrestaImagerieAdm & vbCrLf & "Le mot de passe pour le compte ""cs-prestausr"" est : " & vbTab & vbTab & passwordPrestaImagerieUsr
            Commun.SendEmail("noreply@igbmc.fr", "groupe-mic-photon@igbmc.fr;Bcc:serviceinfo@igbmc.fr", "Nouveaux mots de passe (" & Now.ToString("MMMM" & " " & Now.ToString("yyyy") & ")"), textMail)

            Commun.Journal("Changement des mots de passe des comptes prestataires de l'imagerie réussi", False)

        Catch ex As Exception
            Commun.Journal("ERREUR : ChangePasswordAccountPrestaImagerie : " & ex.Message, True)
        End Try
    End Sub

    Public Sub EnvoiMailCompteExpireXjour(ByVal j As Integer)

        j += 1
        Dim adresseMail As String = ""
        Dim ctrlMailEnvoye As Boolean = False
        Dim dateNowU As String = Now.Date.ToString("yyyyMMddHHmmss.sZ")

        Try
            Dim dateDeSuppressionPrevueUniversal As String = Now.Date.AddDays(j).ToString("yyyyMMddHHmmss.sZ")
            Dim dateDeSuppressionPrevueTxt As String = Now.Date.AddDays(j).ToString("dd/MM/yyyy")
            Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateursDesactives), Commun.admin, Commun.passwd, auth)
                Using searcher As DirectorySearcher = New DirectorySearcher(objAD)
                    searcher.Filter = "(&(objectClass=person)(objectClass=user)(accountDeletionDT=" & dateDeSuppressionPrevueUniversal & "))"

                    Dim results As SearchResultCollection = searcher.FindAll
                    For Each result As SearchResult In results
                        'si l'utilisateur est un externe, on ne lui envoie pas de mail
                        If InStr(result.Path, "(Externe") > 0 Then Continue For
                        Using userAD As DirectoryEntry = result.GetDirectoryEntry
                            adresseMail = userAD.Properties("mail")(0)
                            Dim prenom As String = userAD.Properties("givenName")(0)
                            Dim dateDefinDeContrat As String = userAD.Properties("extensionAttribute1").Value
                            Dim mail = MailCloture(prenom, dateDeSuppressionPrevueTxt, dateDefinDeContrat)

                            Commun.SendEmail("noreply@igbmc.fr", adresseMail & ";Bcc:serviceinfo@igbmc.fr", "ARRET DU COMPTE", mail)
                            Commun.Journal("Mail de fermeture de compte envoyé (-" & j - 1 & ") : " & adresseMail)
                            ctrlMailEnvoye = True
                        End Using
                    Next
                End Using
            End Using

            If ctrlMailEnvoye = True Then
                Commun.Journal("Envoi des mails de cloture de compte (j-" & j & ") terminé avec succes", False)
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : EnvoiMailCompteExpireXjour : Mail de fermeture de compte (-" & j & ") : " & adresseMail & " : " & ex.Message, True)
        End Try
    End Sub
    Public Function ChercherNPlus1deDestination(ByVal idUser As String, ByVal idDest As String) As String
        Dim dateactu As Date = Convert.ToDateTime(Now.ToString("yyyy-MM-dd"))


        Dim IDNPlus1 As String = ""
        Dim dataNplus1 As String = Json.SendJson("", "persons/" & idUser & "/n_plus_uns", "AD", "GET")

        Dim n_plus_uns = Json.DeserializeJson(dataNplus1, "nplusuns")

        For n = 0 To UBound(n_plus_uns)
            If n_plus_uns(n).destination_id = idDest Then
                IDNPlus1 = n_plus_uns(n).id

            End If
        Next n

        Return IDNPlus1

    End Function
    Public Function ContratEnCours(contracts) As Boolean

        ContratEnCours = False
        If contracts.length > 0 Then ContratEnCours = True

        Return ContratEnCours

    End Function
    Public Function Organisme(contracts) As String
        Dim result As String = ""

        For c = 0 To UBound(contracts)
            Dim organismeName As String = ConvertOrganism(contracts(c).organism_id)
            If InStr(result, organismeName) = 0 Then
                If result = "" Then
                    result = organismeName
                Else
                    result = result & "+" & organismeName
                End If
            End If
            '    End If
        Next c
        Return result
    End Function

    Public Function DateDeFinDeContract(ByVal contracts As Object, ByVal id As String, Optional ByVal accepterContratFutur As Boolean = False) As String

        Dim result As String = ""
        Try
            If contracts.length = 1 Then

                If contracts(0).end_date <> "" Then
                    Dim dateEndTxt As String = Strings.Left(contracts(0).end_date, 10)
                    Dim dateTab As String() = Split(dateEndTxt, "-")
                    result = dateTab(2) & "/" & dateTab(1) & "/" & dateTab(0)
                End If
            Else
                Dim Data As String = Json.SendJson("", "persons/" & id & "/contracts", "AD", "GET")
                Dim ResponseData = New JavaScriptSerializer().Deserialize(Of Object)(Data)
                contracts = ResponseData("contracts")

                Dim dateEnd As Date = #1/1/1900#
                Dim dateStart As Date = #1/1/1900#
                Dim dateSelected As Date = #1/1/1900#
                Dim dateFutureSelected As Date = #1/1/1900#

                For c = 0 To UBound(contracts)
                    Dim dateEndTxt As String = Strings.Left(ResponseData("contracts")(c)("end_date"), 10)
                    Dim dateStartTxt As String = Strings.Left(ResponseData("contracts")(c)("start_date"), 10)

                    If dateStartTxt <> Nothing Then
                        dateStart = (Convert.ToDateTime(dateStartTxt))
                    Else
                        Continue For
                    End If


                    If dateEndTxt <> Nothing Then
                        dateEnd = (Convert.ToDateTime(dateEndTxt))
                    Else
                        Return ""
                        Exit Function
                    End If

                    If dateStart <= Now Then
                        If dateEnd > dateSelected Then
                            dateSelected = dateEnd
                        End If
                    ElseIf accepterContratFutur AndAlso dateEnd > dateFutureSelected Then
                        dateFutureSelected = dateEnd
                    End If
                Next c

                If dateSelected <= #1/1/1900# AndAlso accepterContratFutur Then
                    dateSelected = dateFutureSelected
                End If

                result = dateSelected.ToString("dd/MM/yyyy")
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : DateDeFinDeContract : employeeID : " & id & " : " & ex.Message, True)
        End Try


        Return result

    End Function

    Public Sub UpdateFichierHistoAlias()

        Commun.Journal("Mise a jour du fichier d'historique des alias", False)

        'PARTIE ALIAS
        Dim tabAlias(,) As String = Commun.CreateTabHistoAliasLogin("alias")

        Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
            Using searcher As DirectorySearcher = New DirectorySearcher(objAD)
                searcher.Filter = "(&(proxyAddresses=*)(employeeID=*))"
                searcher.PageSize = 2000
                Dim results As SearchResultCollection = searcher.FindAll
                For Each result As SearchResult In results
                    'si l'objet est dans l'OU NAP on continue avec l'objet suivant
                    If InStr(result.Path, "OU=Nap,DC=igbmc,DC=u-strasbg,DC=fr") <> 0 Then Continue For
                    Using objADwithSMTPAddresses As DirectoryEntry = result.GetDirectoryEntry
                        'si l'attribut "SMTPAddresses" existe
                        If objADwithSMTPAddresses.Properties.Contains("proxyAddresses") Then
                            Dim aliasMail As String = ""
                            Dim addresses As New List(Of String)
                            Dim aliasPrecedent As String = ""
                            'Pour toutes les adresses contenue dans l'attribut "proxyAddresses"
                            Dim userID As String = ""

                            If objADwithSMTPAddresses.Properties.Contains("employeeID") Then
                                userID = objADwithSMTPAddresses.Properties("employeeID").Value
                            Else
                                userID = 0
                            End If
                            For Each address In objADwithSMTPAddresses.Properties("proxyAddresses").Value
                                Dim addressTemp As String = LCase(address.ToString)


                                'on ne recupere que les adresses mail basé sur @igbmc.fr (pour eviter de recuperer les adresses externes des contacts ou des utilisateurs de messagerie

                                If Strings.Left(addressTemp, 5) = "smtp:" And InStr(addressTemp, "@igbmc.fr") > 0 Then
                                    addressTemp = Replace(addressTemp, "smtp:", "")
                                    'on recupere l'alias sans l'@ et le domaine
                                    aliasMail = Strings.Left(addressTemp, InStr(addressTemp, "@") - 1)

                                    aliasPrecedent += "'" & aliasMail & "'"
                                End If
                            Next


                            If aliasPrecedent = "" Then Continue For

                            Dim listAlias = Replace(aliasPrecedent, "''", "','")

                            'on recherche l'indexOf de l'userID dans tabAlias
                            Dim index4 As Integer = Commun.MultiIndexOf(tabAlias, userID, 0)

                            'Si on ne trouve pas l'userId, on ajoute une ligne dans le tableau avec les alias
                            If index4 = -1 Then
                                Dim i As Integer = UBound(tabAlias, 2) + 1
                                ReDim Preserve tabAlias(1, i)
                                tabAlias(0, i) = userID
                                tabAlias(1, i) = listAlias
                            Else
                                'si on le trouve, on verifie que tous les alias de l'utilisateur sont inscrit
                                Dim tabAliasFichier As String() = Split(tabAlias(1, index4), ",")
                                Dim tabAliasAD As String() = Split(listAlias, ",")
                                Dim tabAliasEnPlus As String() = tabAliasAD.Except(tabAliasFichier).ToArray
                                If UBound(tabAliasEnPlus) = 0 Then
                                    For j = 0 To UBound(tabAliasEnPlus)
                                        tabAlias(1, index4) = tabAlias(1, index4) & "," & tabAliasEnPlus(j)
                                    Next j
                                End If
                            End If

                        End If
                    End Using
                Next
            End Using

        End Using

        'on ecrit le tableau modifié dans le fichier Alias.txt
        Using sw1 As New StreamWriter("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\Alias.txt", False)
            'on ecrit le tableau alias dans le fichier
            For j = 0 To UBound(tabAlias, 2)
                sw1.WriteLine(tabAlias(0, j) & "§" & tabAlias(1, j))
            Next
        End Using

        'PARTIE LOGIN
        'inscription des login dans le fichier login.txt
        Dim tabLogin(,) As String = Commun.CreateTabHistoAliasLogin("login")
        Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"))
            Using searcher As DirectorySearcher = New DirectorySearcher(objAD)
                searcher.Filter = "(&(sAMAccountName=*)(!(sAMAccountName=*$*)))"
                searcher.PageSize = 2000
                Dim results As SearchResultCollection = searcher.FindAll
                For Each result As SearchResult In results
                    'si l'objet est dans l'OU NAP on continue avec l'objet suivant
                    If InStr(result.Path, "OU=Nap,DC=igbmc,DC=u-strasbg,DC=fr") <> 0 Then Continue For
                    Using objADSAMAccount As DirectoryEntry = result.GetDirectoryEntry
                        Dim login As String = objADSAMAccount.Properties("sAMAccountName").Value
                        Dim userID As String = ""
                        If objADSAMAccount.Properties.Contains("employeeID") Then
                            userID = objADSAMAccount.Properties("employeeID").Value
                        Else
                            userID = 0
                        End If

                        'on recherche l'indexOf de l'userID dans tabAlias
                        Dim index3 As Integer = Commun.MultiIndexOf(tabLogin, userID, 0)

                        'Si on ne trouve pas l'userId, on ajoute une ligne dans le tableau avec les alias
                        If index3 = -1 Then
                            Dim i As Integer = UBound(tabLogin, 2) + 1
                            ReDim Preserve tabLogin(1, i)
                            tabLogin(0, i) = userID
                            tabLogin(1, i) = login
                        Else
                            tabLogin(1, index3) = login
                        End If

                    End Using
                Next result

                Using sw5 As New StreamWriter("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\Login.txt", False)
                    'on ecrit le tableau alias dans le fichier
                    For j = 0 To UBound(tabLogin, 2)
                        sw5.WriteLine(tabLogin(0, j) & "§" & tabLogin(1, j))
                    Next
                End Using
            End Using
        End Using



        Commun.Journal("Mise a jour du fichier d'historique des Alias terminée avec succes", False)
    End Sub
    Public Function EstPair(n As Long) As Boolean
        EstPair = (n Mod 2) = 0
    End Function


    Public Function LeaderDest(ByVal idEntity As String, ByVal idGroup As String, ByVal idDep As String, ByVal idTeam As String, ByVal idDest As String) As String
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
    Public Function ListeDiffusion(ByVal IDuser As String) As String
        Dim dataLD As String = Json.SendJson("", "persons/" & IDuser & "/mailing_lists", "AD", "GET")
        Dim responseDataLDs = New JavaScriptSerializer().Deserialize(Of Object)(dataLD)
        'Dim liste As String = responseDataLD("mailing_lists")
        Dim ld As String = "other"
        For Each responseDataLD In responseDataLDs("mailing_lists")
            ld = responseDataLD("name")
        Next


        Return ld
    End Function



    ''' <summary>
    ''' Détermine si un mail doit être envoyé aux assistants de prévention
    ''' en fonction de la durée cumulée des contrats d'une personne.
    ''' </summary>
    ''' <param name="id">Identifiant de la personne dans IGBMCSERVICES / GDPI.</param>
    ''' <param name="createdUser">
    ''' Indique le contexte du contrôle.
    ''' <c>True</c> pour une création d'utilisateur ;
    ''' <c>False</c> pour une mise à jour d'un utilisateur existant.
    ''' </param>
    ''' <returns>
    ''' <c>True</c> si un mail doit être envoyé aux assistants de prévention ;
    ''' sinon <c>False</c>.
    ''' </returns>
    ''' <remarks>
    ''' La durée des contrats est calculée en jours à partir des contrats récupérés via l'API.
    ''' Les contrats avec le statut <c>running</c> alimentent la durée courante ;
    ''' les autres alimentent la durée passée.
    ''' En création, un mail est envoyé si la durée du contrat courant est au moins égale à 90 jours.
    ''' En modification, un mail est envoyé si la durée cumulée franchit le seuil de 90 jours,
    ''' alors que la durée passée seule était inférieure à ce seuil.
    ''' En cas d'erreur, l'exception est journalisée.
    ''' </remarks>
    Function GetContractsLenght(ByVal id As String, Optional ByVal createdUser As Boolean = False) As Boolean
        Try
            Dim result As Boolean = False
            Dim dataContracts As String = Json.SendJson("", "persons/" & id & "/contracts", "AD", "GET")
            Dim responseContract = New JavaScriptSerializer().Deserialize(Of Object)(dataContracts)
            Dim dureeOld As Integer = 0
            Dim dureeCurrent As Integer = 0
            For Each contract In responseContract("contracts")
                Dim status As String = contract("status")

                Dim startDate As DateTime = contract("start_date")
                Dim endDateTxt = contract("end_date")
                If endDateTxt Is Nothing Then endDateTxt = "2100-01-01T00:00:00"
                Dim endDate As DateTime = endDateTxt

                If status <> "running" Then
                    dureeOld += DateDiff(DateInterval.Day, startDate, endDate)
                Else
                    dureeCurrent += DateDiff(DateInterval.Day, startDate, endDate)
                End If
            Next
            If createdUser = True Then
                If dureeCurrent >= 90 Then
                    'Send Mail
                    result = True
                End If
            Else
                If dureeOld < 90 And dureeCurrent + dureeOld >= 90 Then
                    'Send Mail
                    result = True
                End If
            End If
            Return result
        Catch ex As Exception
            Commun.Journal("ERREUR : Calcul de la durée des contrats : " & id & " : " & ex.Message, True)
        End Try

    End Function
End Module
