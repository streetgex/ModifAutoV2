Imports System.DirectoryServices
Imports System.IO
'Imports ActiveDs
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Collections.ObjectModel
Imports System.Web.Script.Serialization

Public Class Form1

    Shared withJson As String = "debug" 'Valeur possible : json debug temp(fichiers dans le dossier c:\temp)
    Public Shared tabPersoMonoEquipe As String(,)
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Public Shared tabExcepUser(,) As String
    Shared destinationsJsonIsTrue
    'Shared organismContractName As String
    Public Shared dossierArchivePST As String = "\\Space2\archives-pst\"
    Public Shared dossierPhotos As String = "\\Space2\photos-RH\"
    Public Shared nomFichierRapportMS As String = "c:\temp\MSrapport(" & Replace(Now.ToString("dd-MM-yyyy HH.mm"), "/", "-") & ").csv"

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

        Commun.fichierLog = "c:\temp\Log" & application.productname & ".log"
        Commun.Journal("Debut de traitement")


        If Environment.MachineName <> "SERV-AD1" Then

            'controlDoublonUIDNumber()
            'Dim aaa = Commun.UIDNumberMini
            'Pws.DeleteExportRequestCommun("zinka-5960-IGBMC")
            'Supprime.SupprimeMailbox()
            'AttributionStrategieMDP()
            'ExpirationMDP()

            'ctrlMS.recupUserAD()
            'ctrlMS.adddatedefin()
            'Thumbn.ComparePhoto("2503", New DirectoryEntry("LDAP://CN=Stephane CERDAN,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr"))
            'EcritureNumeroBadge(New DirectoryEntry("LDAP://CN=Stephane CERDAN,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr"))
            'Commun.SetADLDAPProperty("steph", "serialNumber", tabNumBadge,)

            'Dim dataContract As String = Json.SendJson("", "persons/" & 2503 & "/contracts", "AD", "GET") 'Json.SendJson("", "persons/" & 6647 & "/contracts", "AD", "GET")
            'Dim contractRunning As Boolean = ContratEnCours(dataContract)
            'Dim finContrat As String = DateDeFinDeContract(dataContract)
            'withJson = "json"


            'Supprime.SupprimeMailbox()

            'Dim ps As New PwShell
            'Dim results = ps.exec("Get-Mailbox steph")
            'ps.close()

            'Supprime.SupprimeMailbox()

            'Pws.CommandePWSCreatePSTMailbox("lang", "2230")
            'GestionComptesExternes()
            'Commun.ClearAttributDeactivationDeletionDate("misra")
            'ExpirationMDP()
            'Supprime.SupprimeMailbox()
            'gestion.GestionAttributsDT()
            'Creation.commandePWSMailbox("tunan", "DB-G31c")
            'Supprime.SupprimeMailbox()
            'Commun.CreateTabExceptionUser()
            'GestionComptesExternes()


        Else
            withJson = "json"
        End If

        'withJson = "debug"

        'Creation et envoi du fichier de controle MicroSesame
        If Hour(Now) = 3 Or Hour(Now) = 4 Then
            Try
                ctrlMS.recupUserAD()
                Commun.Journal("Création et envoi du Raport MicroSesame réussi", False)
            Catch ex As Exception
                Commun.Journal("ERREUR:Creation du rapport MS" & ex.Message, True)
            End Try
        End If

        Dim DebutCreationFichier As DateTime = Now()
        'creer les comptes des utilisateurs qui sont dans le fichier ForceCreationDeCompte.txt sous la forme <EmployéeID>,<Short_name_destination>
        exceptionCreationCompte()
        Dim ctrlCreationFichier As Boolean = CreationFichiers()
        Dim finCreationFichier As DateTime = Now()

        If ctrlCreationFichier = True Then
            Dim duration As TimeSpan = finCreationFichier - DebutCreationFichier
            Dim duree As String = String.Format("{0:00}h {1:00}m {2:00}s", duration.Hours, duration.Minutes, duration.Seconds)
            Dim dureeTolerable As Integer = Convert.ToInt32(RecupDataini.RecupVar("[DureeMaxJson]"))

            If DateDiff(DateInterval.Minute, DebutCreationFichier, finCreationFichier) > dureeTolerable Then
                Commun.SendEmail("administrateur@igbmc.fr", RecupDataini.RecupVar("[mailDureeMaxJson]"), "Duree de creation de fichier JSON superieure à " & dureeTolerable.ToString & " minutes", "La durée de creation du fichier JSON a travers IGBMC services a été anormalement longue: " & duree)
            End If

            Commun.Journal("Création du fichier listepersoJson.txt réussie en : " & duree, False)

            GestionDesFichiers()

            gestion.GestionDestinationsDepartements()
            modifDonneesAD()

            If Environment.MachineName = "SERV-AD1" Then
                gestion.GestionReactiveDesactiveComptesInterne()
            End If

        End If


        GestionComptesExternes()

        'La gestion des AttributsDT doit imperativement intervenir apres GestionReactiveDesactiveComptesInterne
        gestion.GestionAttributsDT()

        gestion.GestionSuppressionPIetDR()

        AttributionStrategieMDP()
        gestion.CtrlGroupAdmins()
        gestion.UpdateComptesProvisoires()
        gestion.ControleOUUtilisateurs()


        If Hour(Now) = 1 Or Hour(Now) = 2 Then
            'Gestion de l'expiration des mot de passe des comptes adm
            ExpirationMDP()

            Commun.Journal("Debut de la gestion de l'envoi des mails de cloture de compte", False)
            EnvoiMailCompteExpireXjour(30)
            EnvoiMailCompteExpireXjour(15)
            EnvoiMailCompteExpireXjour(7)
            EnvoiMailCompteExpireXjour(2)
            EnvoiMailCompteExpireXjour(1)
            Commun.Journal("Fin de la gestion de l'envoi des mails de cloture de compte", False)

            'SansContrat()

            'Changement tous les mois pair, le premier du mois, des mots de passe des comptes prestataires de l'imagerie
            If System.DateTime.Now.ToString("dd") = "01" And EstPair(System.DateTime.Now.ToString("MM")) = True Then
                ChangePasswordAccountPrestaImagerie()
            End If

            UpdateFichierHistoAlias()
            Supprime.SupprimeMailbox()

        End If


        If Commun.controlSendMail = True Then
            Commun.SendEmail(RecupDataini.RecupVar("[AdminScriptLogin]") & "@igbmc.fr", "steph@igbmc.fr", "ModifAuto.NET : Rapport d'erreur", Commun.journalECHECMail)
        End If

        Try
            Shell(RecupDataini.RecupVar("[cheminMAJZoneInfo]"))
        Catch
            Commun.Journal("ERREUR:Lancement de MAJZoneInfo", True)
        End Try

        If File.Exists(nomFichierRapportMS) Then
            File.Delete(nomFichierRapportMS)
        End If
        Commun.Journal("Fin de traitement")
        'CreationAuto.start()
    End Sub
    Shared Sub GestionDesFichiers()
        ' COPIE DES FICHIERS 
        Try

            File.Copy("c:\temp\eq.txt", RecupDataini.RecupVar("[CheminPartage]") & "\todo\eq.txt", True)
            File.Copy("c:\temp\listep.txt", RecupDataini.RecupVar("[CheminPartage]") & "\todo\listep.txt", True)
            If withJson = "json" Then
                File.Copy("c:\temp\listepersoJson.txt", RecupDataini.RecupVar("[CheminPartage]") & "\todo\listepersoJson.txt", True)
                File.Copy("c:\temp\listepersoJson.txt", RecupDataini.RecupVar("[CheminPartage]") & "\cmpttmp\listepersoJson" & Now.ToString("dd") & "-" & Now.ToString("MM") & "-" & Now.ToString("yy") & "-" & Now.ToString("HH") & "h.txt", True)
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
            Kill("c:\temp\listep.txt")
        Catch
            Commun.Journal("ERREUR : suppression du fichier temporaire listep.txt", True)
        End Try
        Try
            'Kill("c:\temp\eq.txt")
        Catch
            Commun.Journal("ERREUR : suppression du fichier temporaire eq.txt", True)
        End Try
    End Sub
    Shared Function TrouverCNUserAvecID(ByVal ID As String) As String

        Dim result As String = ""
        If ID <> "" Then
            Using Ldap As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                Using searcher As DirectorySearcher = New DirectorySearcher(Ldap)
                    searcher.Filter = "(&(objectClass=user) (EmployeeID=" & ID & "))"
                    Dim resultSearcher As SearchResult = searcher.FindOne()
                    result = Replace(resultSearcher.Path, "LDAP://", "")
                End Using
            End Using
        End If
        Return result
    End Function

    Shared Sub EcritureNumeroBadge(ByVal adUser As DirectoryEntry)

        Try
            Dim matricule As String = adUser.Properties("EmployeeID").Value.ToString
            Dim tabNumBadge As String() = jsonMS.AllBadgeNumber(matricule)

            Dim tabBadgeAD As String()
            Dim tabBadgeADTmp = adUser.Properties("serialNumber").Value

            If Not tabBadgeADTmp Is Nothing Then

                If tabBadgeADTmp.GetType.FullName = "System.String" Then
                    tabBadgeAD.Add(tabBadgeADTmp)
                Else
                    For Each obj In tabBadgeADTmp
                        If obj <> "" Then
                            tabBadgeAD.Add(obj)
                        End If
                    Next
                    Array.Sort(tabBadgeAD)
                End If
            End If

            'Dim aaa = tabBadgeADTmp.GetType

            If Not tabNumBadge Is Nothing Then
                Array.Sort(tabNumBadge)
                If tabNumBadge.Count = 1 Then
                    If adUser.Properties.Contains("employeeNumber") Then

                        If adUser.Properties("employeeNumber").Value <> tabNumBadge(0) Then
                            adUser.Properties("employeeNumber").Value = tabNumBadge(0)
                            adUser.CommitChanges()
                        End If
                    Else
                        Commun.SetADLDAPProperty(adUser, "employeeNumber", tabNumBadge(0))
                        adUser.CommitChanges()
                    End If
                End If
            End If

            If Join(tabNumBadge, ",") <> Join(tabBadgeAD, ",") Then
                Commun.SetADLDAPProperty(adUser, "serialNumber", Nothing)
                adUser.CommitChanges()



                If Not tabNumBadge Is Nothing Then
                    For Each badge As String In tabNumBadge
                        adUser.Properties("serialNumber").Add(badge)
                        Commun.Journal("Modification (ajout ou retrait) d'un numero de Badge MicroSesame : " & badge & " : " & adUser.Properties("cn").Value.ToString & " (" & matricule & ")", False)
                    Next

                    'If tabNumBadge.Count = 1 Then
                    '    If adUser.Properties("employeeNumber").Value.ToString <> tabNumBadge(0) Then
                    '        adUser.Properties("employeeNumber").Value = tabNumBadge(0)
                    '    End If
                    'End If
                End If
                adUser.CommitChanges()
            End If


        Catch ex As Exception
            Commun.Journal("ERREUR : ecriture du numero de Badge MicroSesame : " & ex.Message, True)
        End Try

    End Sub

    Shared Sub modifDonneesAD()
        'COMPARAISON DU PERSONNEL ENTRE HIER ET AUJOURDH'HUI ET CREATION DU FICHIER DE MODIFICATIONS

        Commun.Journal("Debut des Modification des Utilisateurs", False)


        Using Ldap As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)

            Dim ancienEQ As String = ""

            Dim test As String = ""

            Json.LoginXivo("https", "serv-xivo:9497")
            Dim listExtensionsXivo As String = ""
            Try
                'listExtensionsXivo = Json.MakeRequestXivo("GET", "https://serv-xivo:9486/1.1/extensions")
                listExtensionsXivo = LCase(Json.MakeRequestXivo("GET", "https://serv-xivo:9486/1.1/lines_sip"))
            Catch ex As Exception
                Commun.Journal("ERREUR : modifDonneesAD : Recuperation des données Xivo", True)
            End Try


            For n = 0 To UBound(tabPersoMonoEquipe, 2)

                'If tabPersoMonoEquipe(2, n) <> "seraphin" Then Continue For
                Dim usrNom As String = tabPersoMonoEquipe(0, n)
                Dim usrPrenom As String = tabPersoMonoEquipe(1, n)
                Dim usrLogin As String = tabPersoMonoEquipe(2, n)
                Dim usrDestNomCourt As String = tabPersoMonoEquipe(3, n)
                Dim usrDestNomLong As String = tabPersoMonoEquipe(4, n)
                Dim usrTempsTravail As Integer = Convert.ToInt32(tabPersoMonoEquipe(5, n))
                Dim usrBureaux As String = tabPersoMonoEquipe(6, n)
                Dim usrTelephones As String = tabPersoMonoEquipe(7, n)
                Dim usrEquipeInfo As String = tabPersoMonoEquipe(8, n)
                Dim IDNplus1 As String = tabPersoMonoEquipe(9, n)
                Dim CNNplus1 As String = ""
                CNNplus1 = TrouverCNUserAvecID(IDNplus1)
                Dim usrID As String = tabPersoMonoEquipe(10, n)
                Dim usrAliasMailLong As String = tabPersoMonoEquipe(11, n) 'avec @igbmc.fr
                Dim ld As String = tabPersoMonoEquipe(12, n)
                Dim organism As String = tabPersoMonoEquipe(13, n)
                Dim finDeContrat As String = tabPersoMonoEquipe(14, n) 'correspond au dernier jour de travail du contrat. reste vide si pas de date de fin connue
                Dim genre As String = tabPersoMonoEquipe(15, n)
                Dim diffusionPhoto As String = tabPersoMonoEquipe(16, n) 'true/false


                'Si l'utilistateur n'existe pas, on continue avec le suivant
                Dim result As SearchResult
                Using searcher As DirectorySearcher = New DirectorySearcher(Ldap)
                    searcher.Filter = "(&(objectClass=user) (EmployeeID=" & usrID & "))"
                    result = searcher.FindOne()
                End Using

                If result Is Nothing Then
                    Commun.Journal("ERREUR : Modification des données : " & usrLogin & " : L'Utilisateur n'existe pas", True)
                    Continue For
                End If

                Using objuser As DirectoryEntry = result.GetDirectoryEntry

                    result = Nothing

                    'recuperation de l'ancienne equipe comptable
                    If objuser.Properties("DepartmentNumber").Value <> "" Then
                        ancienEQ = objuser.Properties("DepartmentNumber").Value
                    End If
                    'On Error Resume Next
                    'test
                    If Thumbn.controlServeurZoneLabo = True Then
                        Dim testPhotoExist As Boolean = System.IO.File.Exists(Thumbn.pathPhoto & usrID & ".jpg")
                        If testPhotoExist = True Then
                            'Gestion des Thumbnails
                            'Thumbn.ComparePhoto(TransformeSAMACCOUNTenCN(usrLogin), usrID, objuser)
                            Thumbn.ComparePhoto(usrID, objuser)
                        Else
                            If Not objuser.Properties("jpegPhoto").Value Is Nothing Or Not objuser.Properties("thumbnailPhoto").Value Is Nothing Then
                                'Si pas de photo dans le dossier, on efface l'attribut AD
                                objuser.Properties("jpegPhoto").Value = Nothing
                                objuser.Properties("thumbnailPhoto").Value = Nothing
                                Commun.AppliquerChangement(objuser)
                                Commun.Journal("Nettoyage des attributs Photos Réussi : " & usrLogin)
                            End If
                        End If
                    End If

                    'Ecriture du Numero de badge extrait de MicroSesame
                    EcritureNumeroBadge(objuser)



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

                    'mise a jour du genre: Homme Femme
                    If Not objuser.Properties.Contains("extensionAttribute2") Then
                        objuser.Properties("extensionAttribute2").Add(genre)
                        Commun.AppliquerChangement(objuser)
                    Else
                        'MsgBox("Attribut ""extensionAttribute2"" existe")
                    End If
                    'Commun.AppliquerChangement(objuser)

                    'gestion de l'organisme
                    Try
                        If objuser.Properties("Division").Value <> organism Then
                            Commun.SetADLDAPProperty(objuser, "Division", organism)
                            Commun.AppliquerChangement(objuser)
                            Commun.Journal("Changement de l'organisme réussi : " & usrLogin)
                        End If
                    Catch ex As Exception
                        Commun.Journal("ERREUR : Modification de l'organisme : " & usrLogin & " : " & ex.Message, True)
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
                    If listExtensionsXivo <> "" Then

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
                                    If InStr(tabPhoneFichier(h), "/") > 0 Then
                                        Dim tmpTab As String() = Split(tabPhoneFichier(h), "/")
                                        For Each tel In tmpTab
                                            tabOtherphoneFichier.Add(tel)
                                        Next
                                    Else
                                        tabOtherphoneFichier.Add(tabPhoneFichier(h))
                                    End If

                                    'ReDim Preserve tabOtherphoneFichier(h - 1)
                                    'tabOtherphoneFichier(h - 1) = tabPhoneFichier(h)
                                Next h
                                'Erase tabOtherphoneFichier

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
                                'If InStr(listExtensionsXivo, """exten"": """ & IPphone & """") <> 0 Then 
                                If InStr(listExtensionsXivo, """callerid"": ""\""" & LCase(usrPrenom & " " & usrNom) & "\""") <> 0 Then
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
                                    Commun.Journal("Suppression de l'attribut ""IpPhone"" Réussi : " & usrLogin)
                                    Commun.AppliquerChangement(objuser)
                                Else
                                    Commun.SetADLDAPProperty(objuser, "IpPhone", IPphone)
                                    Commun.Journal("Modification de l'attribut ""IpPhone"" Réussi : " & usrLogin)
                                    Commun.AppliquerChangement(objuser)
                                End If
                            End If
                        Catch ex As Exception
                            Commun.Journal("ERREUR : Modification des données téléphoniques ""IPPhone : " & usrLogin & " : " & IPphone & " : " & ex.Message, True)
                        End Try
                    End If
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

                    'Gestion de la date de fin de contrat
                    Try
                        If objuser.Properties("extensionAttribute1").Value <> finDeContrat And objuser.Parent.Path = "LDAP://" & RecupDataini.RecupVar("[OUUtilisateursActifs]") Then
                            Commun.SetADLDAPProperty(objuser, "extensionAttribute1", finDeContrat)
                            Commun.AppliquerChangement(objuser)
                        End If
                    Catch
                        Commun.Journal("ERREUR : Modification : Date de fin de contrat : " & usrLogin, True)
                    End Try
                    'Données UNIX AD

                    If Not objuser.Properties.Contains("uidNumber") Then
                        Dim uidNumber As String = Commun.UIDNumberMini()
                        Commun.SetADLDAPProperty(objuser, "uidNumber", uidNumber)
                        Commun.AppliquerChangement(objuser)
                        Commun.Journal("Creation de l'attribut AD ""uidNumber"" réussi : " & usrLogin)
                    End If

                    Dim err As String = ""
                    Try
                        Dim uidAD As String = objuser.Properties("SAMAccountName").Value
                        'Dim uidNumberAD As String = objuser.Properties("uidnumber").Value
                        Dim gidNumberAD As String = Commun.FindAttribut(usrEquipeInfo & "_eq", "gidNumber")
                        Dim cheminZoneLabo As String = LCase(Commun.FindAttribut(usrEquipeInfo & "_eq", "url"))
                        Dim unixHomeDirectoryAD As String = LCase(Replace(Replace(cheminZoneLabo, "\\", "\"), "\", "/") & "/" & usrLogin)
                        If InStr(cheminZoneLabo, "labo4") > 0 Then
                            unixHomeDirectoryAD = Replace(cheminZoneLabo, "\\", "\")
                            unixHomeDirectoryAD = Replace(unixHomeDirectoryAD, "\", "/")
                            unixHomeDirectoryAD = Replace(unixHomeDirectoryAD, "-", "_")
                            unixHomeDirectoryAD = unixHomeDirectoryAD & "/" & usrEquipeInfo & "/" & usrLogin
                        End If
                        Dim loginShellAD As String = objuser.Properties("loginShell").Value

                        If objuser.Properties("uid").Value <> uidAD Then
                            err = "uid"
                            Commun.SetADLDAPProperty(objuser, "uid", uidAD)
                            Commun.AppliquerChangement(objuser)
                            Commun.Journal("Modification de l'attribut AD ""uid"" réussi : " & usrLogin)
                        End If

                        If objuser.Properties("gidNumber").Value <> gidNumberAD Then
                            err = "gidNumber"
                            Commun.SetADLDAPProperty(objuser, "gidNumber", gidNumberAD)
                            Commun.AppliquerChangement(objuser)
                            Commun.Journal("Modification de l'attribut AD ""gidNumber"" réussi : " & usrLogin)
                        End If


                        If objuser.Properties("unixHomeDirectory").Value <> unixHomeDirectoryAD Then
                            err = "unixHomeDirectory"
                            Commun.SetADLDAPProperty(objuser, "unixHomeDirectory", unixHomeDirectoryAD)
                            Commun.AppliquerChangement(objuser)
                            Commun.Journal("Modification de l'attribut AD ""unixHomeDirectory"" réussi : " & usrLogin)
                        End If

                        If Not objuser.Properties.Contains("loginShell") Then
                            Commun.SetADLDAPProperty(objuser, "loginShell", "/bin/bash")
                        Else
                            If objuser.Properties("loginShell").Value = "" Then
                                Commun.SetADLDAPProperty(objuser, "loginShell", "/bin/bash")
                            End If
                        End If

                    Catch ex As Exception
                        Commun.Journal("ERREUR : Modification des attributs Unix de l'utilisateur : attribut " & err & " : " & usrLogin & " : " & ex.Message, True)
                    End Try


                    'Chef d'équipe
                    Try

                        'Cas de Seraphin, pour eviter les boucles dans l'organigramme
                        If usrLogin = RecupDataini.RecupVar("[LoginDirecteur]") Then
                            If objuser.Properties.Contains("Manager") Then
                                Commun.SetADLDAPProperty(objuser, "Manager", "")
                                Commun.AppliquerChangement(objuser)
                            End If
                        Else
                            'Si le chef d'equipe n'est pas vide et qu'il est different de celui enregistré dans l'AD
                            If CNNplus1 <> "" And objuser.Properties("Manager").Value <> CNNplus1 Then
                                objuser.Properties("Manager").Value = CNNplus1
                                Commun.AppliquerChangement(objuser)
                                Commun.Journal("Modification de l'attribut ""Manager"" Réussi : " & usrLogin)
                            End If
                            'On Error GoTo 0
                        End If
                    Catch
                        Commun.Journal("ERREUR : Chef d'equipe de l'utilisateur : " & usrLogin & " : " & CNNplus1, True)
                    End Try

                    'gestion des appartenance aux groupes de destinations (grp)
                    CompararaisonAjoutRetraitDestinations(objuser)

                    'gestion des listes de diffusion phd et postdoc
                    Try
                        If ld = "phd" Or ld = "postdoc" Then
                            If Commun.AppartientGroup(usrLogin, ld) = False Then
                                Commun.AddRemoveADGroup(usrLogin, ld, "Add")
                                Commun.Journal("Ajout liste de diffusions " & ld & " : " & usrLogin)
                            End If
                        Else
                            If Commun.AppartientGroup(usrLogin, "phd") = True Then
                                Commun.AddRemoveADGroup(usrLogin, "phd", "Remove")
                                Commun.Journal("Retrait liste de diffusions PHD : " & usrLogin)
                            End If
                            If Commun.AppartientGroup(usrLogin, "postdoc") = True Then
                                Commun.AddRemoveADGroup(usrLogin, "postdoc", "Remove")
                                Commun.Journal("Retrait liste de diffusions POSTDOC : " & usrLogin)
                            End If
                        End If
                    Catch e As Exception
                        Commun.Journal("ERREUR : Ajout liste de diffusions PHD ou POSTDOC : " & e.Message & " : " & usrLogin & " : " & ld, True)
                    End Try

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
                            'Dim objDom As ActiveDs.IADsContainer = GetObject(objuser.Parent.Path)
                            'objDom.MoveHere(objuser.Path, "cn=" & objuser.Properties("displayName").Value)
                            'objDom = Nothing
                            objuser.Rename("CN=" & objuser.Properties("displayName").Value)
                        Catch ex As Exception
                            Commun.Journal("ERREUR : Modification : Renomer l'objet utilisateur : " & usrLogin & " : " & ex.Message, True)
                        End Try
                    End If

                    'Controle et Modification de l'alias sur IGBMCSERVICES
                    Dim aliasConstruit As String = LCase(usrPrenom & "." & usrNom)
                    aliasConstruit = Replace(aliasConstruit, " ", "-")
                    aliasConstruit = Replace(aliasConstruit, "'", "")
                    Dim aliasmailAD As String = aliasConstruit & "@igbmc.fr"
                    If aliasmailAD <> usrAliasMailLong Then
                        Try
                            Dim aliasMailpourIGBMCSERVICES As String = Replace(aliasmailAD, "@igbmc.fr", "")
                            Json.SendJson("login=" & objuser.Properties("SAMAccountName").Value & "&domain=%40igbmc.fr&alias=" & aliasMailpourIGBMCSERVICES, "persons/" & usrID & "/email", "AD", "PUT")
                            objuser.Properties("msExchExtensionAttribute16").Value = LCase(aliasmailAD)
                            Commun.AppliquerChangement(objuser)
                            Commun.Journal("Modification de l'alias Long Réussi : " & usrLogin & " : " & aliasmailAD)
                        Catch ex As Exception
                            Commun.Journal("ERREUR : Modification alias Mail sur IGBMCSERVICES : " & objuser.Properties("SAMAccountName").Value & " : " & ex.Message, True)
                        End Try
                    End If
                End Using


            Next n

        End Using
        Commun.Journal("Gestion des modifications utilisateurs réussie", False)
    End Sub


    Shared Function CreationFichiers()
        Dim result As Boolean = False
        If withJson = "json" Then
            Dim dataJsonDestinationTrue As String = Json.SendJson("", "destinations?is_team=true", "AD", "GET")
            If dataJsonDestinationTrue = Nothing Then End
            destinationsJsonIsTrue = Json.DeserializeJson(dataJsonDestinationTrue, "destinations")
            creationFichierParJson()
        ElseIf withJson = "debug" Then
            File.Copy(RecupDataini.RecupVar("[CheminPartage]") & "\todo\eq.txt", "c:\temp\eq.txt", True)
            File.Copy(RecupDataini.RecupVar("[CheminPartage]") & "\todo\listepersoJson.txt", "c:\temp\listepersoJson.txt", True)
        ElseIf withJson = "temp" Then

        Else
            End
        End If

        Try
            'DETERMINER L'EQUIPE PRINCIPALE

            FileOpen(1, "c:\temp\listepersoJson.txt", OpenMode.Input) 'fichier du personnel sans accent multi equipe

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
                    ReDim Preserve tabPersoMonoEquipe(16, g)
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
                    tabPersoMonoEquipe(12, g) = lignePerso(12)
                    tabPersoMonoEquipe(13, g) = lignePerso(13)
                    tabPersoMonoEquipe(14, g) = lignePerso(14)
                    tabPersoMonoEquipe(15, g) = lignePerso(15)
                    tabPersoMonoEquipe(16, g) = lignePerso(16)
                    lignePerso = Nothing

                    If tabPersoMonoEquipe Is Nothing Then Throw New Exception("tabPerso est vide")

                    If g <> 0 Then
                        'comparaison de 2 lignes du meme utilisateur (dans le cas d'equipe multiple) 
                        If tabPersoMonoEquipe(10, g) = tabPersoMonoEquipe(10, g - 1) Then
                            ''on ajoute les organismes ensemble s'il ne sont pas deja présent dans le champ
                            'If InStr(tabPersoMonoEquipe(13, g - 1), tabPersoMonoEquipe(13, g)) = 0 Then
                            '    tabPersoMonoEquipe(13, g - 1) = tabPersoMonoEquipe(13, g)
                            'End If

                            'Si le temps de travail dans une equipe, pour le meme utilisateur, est superieur a la ligne precedente, on reecrit le debut de la ligne precedente
                            If tabPersoMonoEquipe(5, g) > tabPersoMonoEquipe(5, g - 1) Then
                                tabPersoMonoEquipe(0, g - 1) = tabPersoMonoEquipe(0, g)
                                tabPersoMonoEquipe(1, g - 1) = tabPersoMonoEquipe(1, g)
                                tabPersoMonoEquipe(2, g - 1) = tabPersoMonoEquipe(2, g)
                                tabPersoMonoEquipe(3, g - 1) = tabPersoMonoEquipe(3, g)
                                tabPersoMonoEquipe(4, g - 1) = tabPersoMonoEquipe(4, g)
                                tabPersoMonoEquipe(5, g - 1) = tabPersoMonoEquipe(5, g)
                                tabPersoMonoEquipe(8, g - 1) = tabPersoMonoEquipe(8, g)
                                'Cas (Gilles Duval) faisant parti de 2 équipes, et etant chef de celle ou il travaille le plus => aucun chef ne resortait (bug igbmcservices)
                                If tabPersoMonoEquipe(9, g) <> "" Then
                                    tabPersoMonoEquipe(9, g - 1) = tabPersoMonoEquipe(9, g)
                                End If
                                tabPersoMonoEquipe(13, g - 1) = tabPersoMonoEquipe(13, g)
                                tabPersoMonoEquipe(14, g - 1) = tabPersoMonoEquipe(14, g)
                                g = g - 1
                            Else
                                If tabPersoMonoEquipe(9, g - 1) = "" Then
                                    tabPersoMonoEquipe(9, g - 1) = tabPersoMonoEquipe(9, g)
                                End If
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
                            telTemp = telTemp & ";" & tabTelTemp(j)
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
            result = True
        Catch ex As Exception
            Commun.Journal("ERREUR : Creation des fichiers et des tableaux : " & ex.Message, True)
            Commun.SendEmail(RecupDataini.RecupVar("[AdminScriptLogin]") & "@igbmc.fr", "steph@igbmc.fr", "ModifAuto.NET : Rapport d'erreur", Commun.journalECHECMail)
            Return result
            Exit Function
        End Try
        Return result
    End Function
    Shared Sub GestionComptesExternes()
        Commun.Journal("Debut de la gestion des comptes Externes", False)

        Const ADS_UF_ACCOUNT_DISABLE = 2
        Dim i = 0
        Using objADExt As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursExternes]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Dim results As SearchResultCollection
            Using searcher As DirectorySearcher = New DirectorySearcher(objADExt)
                searcher.SearchScope = SearchScope.OneLevel
                'searcher.Filter = "(&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)(!(msExchRecipientTypeDetails=1))))"
                searcher.Filter = "(&((&(objectClass=user)((!(msExchRecipientTypeDetails=1)))))(!(SAMAccountName=gardiens)))"
                results = searcher.FindAll
            End Using
            For Each result As SearchResult In results
                Dim login1 As String = result.Properties("SAMAccountName")(0)
                Try

                    Using userExt As DirectoryEntry = result.GetDirectoryEntry
                        'If Strings.InStr(userExt.Properties("Mail").Value, "@igbmc.fr") > 0 Then
                        '    MsgBox(userExt.Properties("DisplayName").Value)
                        '    Continue For
                        'End If
                        Dim login As String = userExt.Properties("SAMAccountName").Value
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
                        If (Not userExt.Properties.Contains("Mail") Or Not userExt.Properties.Contains("proxyaddresses")) Then 'Or Not userExt.Properties.Contains("homeMDB") Then
                            'Si la propriété "Mail" est presente, on lance la commande Powershell pour definir l'utilisateur comme "utilisateur avec Messagerie" dans Exchange
                            If userExt.Properties.Contains("Mail") Then
                                Dim mail As String = userExt.Properties("Mail").Value
                                If Strings.InStr(mail, "@igbmc.fr") = 0 Then
                                    If commandePWSMailUser(userExt.Properties("samAccountName").Value, mail) = True Then
                                        userExt.Properties("description").Clear()
                                        Commun.AppliquerChangement(userExt)
                                        Commun.ReactiveDesactiveCompte(userExt, "active")
                                        Sleep(20000)
                                        Commun.Journal("GestionComptesExternes : User Externe : Utilisateur de messagerie créé : " & login)
                                    Else
                                        Using AD As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
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
                                    Using ADuser As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursActifs]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                                        Using searcherMail As DirectorySearcher = New DirectorySearcher(ADuser)
                                            searcherMail.SearchScope = SearchScope.OneLevel
                                            searcherMail.Filter = "((proxyAddresses=*" & mail & "))"
                                            Dim result1 As SearchResult = searcherMail.FindOne
                                            If result1 Is Nothing Then
                                                userExt.Properties("description").Value = "Compte externe lié a un utilisateur interne parti"
                                                userExt.Properties("Mail").Value = Nothing
                                                Commun.AppliquerChangement(userExt)
                                                Commun.ReactiveDesactiveCompte(userExt, "desactive")
                                                Commun.Journal("GestionComptesExternes : User Externe : Compte externe désacivé lié a une adresse interne inexistante : " & login)
                                            End If
                                        End Using
                                    End Using
                                End If
                                'si la propiété "proxyaddresses" est definie, on recupere l'adresse principale pour definir la propriété "Mail"
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
                                userExt.Properties("Mail").Value = addresseMailIdentifie
                                userExt.Properties("description").Clear()
                                Commun.AppliquerChangement(userExt)
                                Commun.ReactiveDesactiveCompte(userExt, "active")
                                Commun.Journal("GestionComptesExternes : User Externe : Mise a jour de l'adresse mail : " & login)
                            Else
                                'MsgBox(userExt.Properties("displayname").Value)
                                If Commun.AccountIsDisabled(userExt) = False Then
                                    userExt.Properties("description").Value = "Desactivé le " & Strings.Left(CStr(Now), 10) & " : Aucune Adresse Mail de contact"
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
                        If "SMTP:" & userExt.Properties("Mail").Value <> userExt.Properties("TargetAddress").Value Then
                            Dim newMail As String = "SMTP:" & userExt.Properties("Mail").Value
                            Dim oldMail As String = userExt.Properties("TargetAddress").Value
                            Dim currentAddresses = userExt.Properties("proxyAddresses").Value


                            Dim index As Integer = Array.IndexOf(currentAddresses, oldMail)
                            currentAddresses(index) = newMail
                            userExt.Properties("proxyAddresses").Value = currentAddresses
                            userExt.Properties("TargetAddress").Value = newMail

                            userExt.CommitChanges()

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
    Shared Function commandePWSMailUser(ByVal aliasMail As String, ByVal externalEmail As String) As Boolean
        Dim result As Boolean = False
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
            pCredential = DirectCast(Nothing, PSCredential) 'New PSCredential("igbmc\steph", CreateSecurePasswordString("aaaaaa"))

            '-- set connection info
            pConnectionInfo = New WSManConnectionInfo(New Uri("http://serv-mbx12.igbmc.u-strasbg.fr/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", pCredential)
            pConnectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos
            '-- create remote runspace
            pRunspace = RunspaceFactory.CreateRunspace(pConnectionInfo)
            pRunspace.Open()

            '-- create powershell
            pShell = PowerShell.Create
            pShell.Runspace = pRunspace

            '-- create command
            pCommand = New PSCommand
            With pCommand
                .AddCommand("Enable-MailUser")
                .AddParameter("identity", aliasMail)
                .AddParameter("alias", aliasMail)
                .AddParameter("ExternalEmailAddress", externalEmail)
                .AddParameter("DomainController", ctrlDomain)
            End With

            '-- add command to powershell
            pShell.Commands = pCommand

            '-- invoke the powershell
            pResult = pShell.Invoke
            If pResult.Count = 1 Then result = True
            pRunspace.Close()
            pRunspace.Dispose()

        Catch e As Exception

        End Try
        Return result
    End Function



    Shared Function MailCloture(ByVal prenom As String, ByVal dateDeSuppressionPrevue As String) As String
        Return "Bonjour " & prenom & ", " & vbCrLf & vbCrLf & "Votre compte (mail, acces informatique, zone labo, profil,...) à l'IGBMC va etre definitivement supprimé le " _
            & dateDeSuppressionPrevue & "." & vbCrLf & "Les données de votre zone Labo ont été mises à la disposition de votre chef d'equipe. Pour les récupérer, veuillez le contacter." _
            & vbCrLf & "Veuillez récuperer les éléments de votre boite mail que vous souhaitez conserver, avant cette date." _
            & vbCrLf & "Apres cette date, plus aucune récuperation, ne sera possible, elles seront definitivement detruites." _
            & vbCrLf & vbCrLf & "En cas de besoin d'assistance, vous pouvez envoyer un mail à helpdesk@igbmc.fr" _
            & vbCrLf & vbCrLf & "Le service Informatique."
    End Function

    Shared Function MailCloture1(ByVal prenom As String, ByVal dateDeSuppressionPrevue As String, dateFindeContrat As String) As String

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
             vbCrLf & vbCrLf & "Votre boite mail restera accessible pendant une période de " & nbrJourRestant & " jours à compter d’aujourd’hui.<BR>" &
             vbCrLf & "Vous pouvez notamment vous y connecter au travers du webmail de l’IGBMC : <a href=""https://igbmcmail.igbmc.fr"">https://igbmcmail.igbmc.fr</a><BR><BR>" &
             vbCrLf & vbCrLf & "Vous n'avez à présent plus accès aux ressources informatiques de l’IGBMC (espaces de stockage, ressources de calcul, postes de travail, etc.).<BR>" &
             vbCrLf & "Les données de votre zone Labo ont été mises à la disposition de votre chef d'équipe. Si vous n’avez pas pris le temps de les récupérer, merci de le contacter.<BR><BR>" &
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
             vbCrLf & "If a new contract is registered within X days, your account will be automatically reactivated.<BR><BR>" &
             vbCrLf & vbCrLf & "Your mailbox will remain accessible for a period of " & nbrJourRestant & " days starting today.<BR>" &
             vbCrLf & "You can connect to it via the IGBMC webmail: <a href=""https://igbmcmail.igbmc.fr"">https://igbmcmail.igbmc.fr</a><BR><BR>" &
             vbCrLf & vbCrLf & "You no longer have access to the IT resources of the IGBMC (storage spaces, computing resources, workstations, etc.).<BR>" &
             vbCrLf & "The data in your Lab area has been made available to your team leader. If you have not taken the time to retrieve them, please contact her/him.<BR><BR>" &
             vbCrLf & vbCrLf & "<B>Warning :<BR>" &
             vbCrLf & "<FONT color=""red"">Your computer account will be permanently deleted on " & dateDeSuppressionPrevue & ".</FONT><BR>" &
             vbCrLf & "After this date, no more data recovery or account reactivation will be possible. Your IT account and your mailbox will be permanently destroyed.</B><BR><BR>" &
             vbCrLf & vbCrLf & "If the closure of your computer account appears abnormal," &
             " please contact your manager in the human resources department to request the update of your contract’s details.<BR><BR>" &
             vbCrLf & vbCrLf & "The IT department.<BR>" &
             vbCrLf & "</body>" &
             "</html>"


    End Function

    Shared Function UserMembreDeDestination(ByVal username As String) As String()
        Dim appartientA As String()
        Dim j As Integer = 0
        Dim Entry As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
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
    Public Shared Function chercheDepartementparRapporDestination(ByVal destination As String) As String
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
    Public Shared Sub CompararaisonAjoutRetraitDestinations(ByVal usrDir As DirectoryEntry)
        Dim login As String = usrDir.Properties("SAMAccountName").Value
        Dim EquipeUserAD As String() = UserMembreDeDestination(login)
        Dim EquipeUserFichier As String() = EquipeComptableFichierUser(login)
        Try

            'comparaison des 2 tableaux entre eux

            If EquipeUserAD Is Nothing Then
                EquipeUserAD = {} 'EquipeUserFichier
            End If

            If EquipeUserFichier Is Nothing Then
                EquipeUserFichier = {} 'EquipeUserAD
            End If

            Dim addToAD As String() = EquipeUserFichier.Except(EquipeUserAD).ToArray

            Dim remFromAD As String() = EquipeUserAD.Except(EquipeUserFichier).ToArray



            'Ajout de l'utilisateur aux groupes-destinations du fichier json auquels il n'appartient pas dans l'ad
            'L'ajout au departement se fait un peu plus bas
            For Each grp As String In addToAD
                If Commun.AppartientGroup(login, "G_Domain_DisableOpenSession") = False Then 'And Commun.AppartientGroup(login, grp) = False Then
                    Commun.AddRemoveADGroup(login, grp, "Add")
                    Commun.Journal("Succes : CompararaisonAjoutRetraitEquipesComptables : " & login & " ajouté dans " & grp)
                End If
            Next

            'Retrait de l'utilisateur des groupes-destinations et departement de l'AD auquels il n'appartient plus d'apres le fichier Json
            For Each grp As String In remFromAD
                Dim dptGrp As String = chercheDepartementparRapporDestination(grp)
                'If Commun.AppartientGroup(login, grp) = True Then
                Commun.AddRemoveADGroup(login, grp, "Remove")
                If dptGrp <> "" Then
                    Commun.AddRemoveADGroup(login, dptGrp, "Remove")
                End If
                Commun.Journal("Succes : CompararaisonAjoutRetraitEquipesComptables : " & login & " retiré de " & grp)
                'End If
            Next

            'Ajout de l'utilisateur au groupe du departement dont ses destinations dependent
            'Specifiquement dans le cas d'un nouvel utilisateur qui est ajouté a sa destination au moment de la création de son compte (autocompte)
            'Modifauto ne l'ajoutera pas a la destination donc pas au département
            If Not EquipeUserAD Is Nothing Then
                If Commun.AppartientGroup(login, "G_Domain_DisableOpenSession") = False Then
                    For Each dpt As String In EquipeUserAD
                        Dim dptGrp As String = chercheDepartementparRapporDestination(dpt)
                        If dptGrp <> "" Then
                            Commun.AddRemoveADGroup(login, dptGrp, "Add")
                        End If
                    Next
                End If
            End If

            EquipeUserAD = Nothing
            EquipeUserFichier = Nothing
        Catch ex As Exception
            Commun.Journal("ERREUR : CompararaisonAjoutRetraitDestinations : " & login & " : " & ex.Message, True)
        End Try
    End Sub

    Shared Sub modifAliasMail(ByVal objuser As DirectoryEntry, ByVal prenom As String, ByVal nom As String)

        Dim ctrlChangement As Boolean = False
        'On met le userID de la fonction Commun.DetermineAliasLibre afin d'optenir directement un alias dispo sans rechercher le userID dans l'historique alias 
        Dim aliasMail As String = Commun.DetermineAliasLibre(prenom, nom, "0")(0)


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
            'Dim nvlAdresse3 As String = "smtp:" & aliasMail & "@igbmc.u-strasbg.fr"
            'Dim nvlAdresse4 As String = "SMTP:" & aliasMail & "@igbmc.u-strasbg.fr"

            Dim iof1 As Integer = adress.IndexOf(nvlAdresse1)
            Dim iof2 As Integer = adress.IndexOf(nvlAdresse2)
            'Dim iof3 As Integer = adress.IndexOf(nvlAdresse3)
            'Dim iof4 As Integer = adress.IndexOf(nvlAdresse4)

            If iof1 = -1 And iof2 = -1 Then
                adress.Add(nvlAdresse1)
                objuser.Properties("proxyAddresses").Value = adress.ToArray()
                Commun.AppliquerChangement(objuser)
                ctrlChangement = True
            End If
            'If iof3 = -1 And iof4 = -1 Then
            '    adress.Add(nvlAdresse3)
            '    objuser.Properties("proxyAddresses").Value = adress.ToArray()
            '    Commun.AppliquerChangement(objuser)
            '    ctrlChangement = True
            'End If
            currentAddresses = Nothing
            adress.Clear()
            adress = Nothing

            'Si il y a eu des changements, enregistrement de l'alias mail dans l'historique
            If ctrlChangement = True Then
                Commun.ajoutAliasFichierHisto(aliasMail, objuser.Properties("EmployeeID").Value)
                objuser.Properties("msExchExtensionAttribute16").Value = aliasMail & "@igbmc.fr"
                Commun.AppliquerChangement(objuser)
            End If

        Catch ex As Exception
            Commun.Journal("ERREUR : Modification alias Mail sur AD : " & objuser.Properties("SAMAccountName").Value & " : " & ex.Message, True)
        End Try

    End Sub

    Shared Sub AttributionStrategieMDP()

        'Cas adminInfo
        Dim tabresults As String()
        Dim tabPoste As String()
        Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://OU=AdmInfo,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Dim results As SearchResultCollection = Commun.SearchFilterAll(objAD, "(&(objectClass=user)(!(distinguishedName=CN=userprog,OU=AdmInfo,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr)))", SearchScope.Subtree)
            If Not results Is Nothing Then
                For Each result As SearchResult In results
                    'Commun.AddRemoveADGroup(result.Properties("samAccountName")(0), "G_SMDPM_Admins", "Add")
                    tabresults.Add(Replace(result.Path, "LDAP://", ""))
                Next
            End If
        End Using

        updateGroupeWithArray("G_SMDPM_Admins", tabresults)

        'cas UsersAdm
        Dim tabresults1 As String()
        Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://OU=UsersAdm,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Dim results As SearchResultCollection = Commun.SearchFilterAll(objAD, "(&(objectClass=user)(!(distinguishedName=CN=userprog,OU=AdmInfo,OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr)))", SearchScope.Subtree)
            If Not results Is Nothing Then
                For Each result As SearchResult In results
                    'Commun.AddRemoveADGroup(result.Properties("samAccountName")(0), "G_SMDPM_Admins", "Add")
                    tabresults1.Add(Replace(result.Path, "LDAP://", ""))
                Next
            End If
        End Using


        updateGroupeWithArray("G_SMDPM_UsersAdm", tabresults1)

        'cas Users-Admins
        Dim tabresults2 As String()
        Using objGroupAdminPoste As DirectoryEntry = New DirectoryEntry("LDAP://OU=Admins,OU=Postes,OU=Micro,OU=Groupes,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Dim results As SearchResultCollection = Commun.SearchFilterAll(objGroupAdminPoste, "(&(objectCategory=group)(member=*))", SearchScope.Subtree)
            Dim dejaMembreStrategie As String() = Commun.MembresDuGroupe("G_SMDPM_Users-Admins")
            If Not results Is Nothing Then
                For Each result As SearchResult In results
                    Dim nomDuGroup As String = Commun.TransformeSAMACCOUNTenCN(Replace(result.Path, "LDAP://", ""))
                    Dim membresGroup As String() = Commun.MembresDuGroupe(nomDuGroup)
                    If membresGroup Is Nothing Then Continue For
                    Dim nomDuPoste As String = Replace((Replace(nomDuGroup, "DL_", "")), "_Admins", "")
                    'tabPoste.Add(Commun.TransformeSAMACCOUNTenCN(nomDuPoste & "$"))
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
        Using groupePoste As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.TransformeSAMACCOUNTenCN("PostesAvecAdmin"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            groupePoste.Properties("member").Value = tabPoste
            groupePoste.CommitChanges()
        End Using


        updateGroupeWithArray("G_SMDPM_Users-Admins", tabresults2)



    End Sub
    Shared Sub DisablePasswordNeverExpiresETPasswordLastSet(ByVal usrPath As String)
        Using user As DirectoryEntry = New DirectoryEntry("LDAP://" & usrPath, Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Const NON_EXPIRE_FLAG = &H10000
            Dim Val As Integer = user.Properties("userAccountControl").Value
            'Defini le flag ADS_UF_DONT_EXPIRE_PASSWD sur "non coché"
            user.Properties("userAccountControl").Value = Val And Not NON_EXPIRE_FLAG
            user.CommitChanges()
            'Defini le dernier changement de mot de passe a aujourd'hui
            user.Properties("pwdLastSet").Value = 0
            user.CommitChanges()
            user.Properties("pwdLastSet").Value = -1
            user.CommitChanges()
        End Using
    End Sub

    Shared Sub updateGroupeWithArray(ByVal nomGroupe As String, ByVal arr As String())
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
    Shared Sub ExpirationMDP()
        Dim samaccountname As String
        Try

            Using OUAdmins As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                Using OUAdminsearcher As DirectorySearcher = New DirectorySearcher(OUAdmins)
                    OUAdminsearcher.Filter = "(&(objectClass=user)(|((memberof=" & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_Admins") & ")(memberof=" & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_Users-Admins") & "))))"
                    OUAdminsearcher.Filter = "(&(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)(|((memberof=" & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_Admins") & ")(memberof=" & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_Users-Admins") & ")(memberof=" & Commun.TransformeSAMACCOUNTenCN("G_SMDPM_UsersAdm") & "))))"
                    OUAdminsearcher.PropertiesToLoad.Add("pwdLastSet")
                    OUAdminsearcher.PropertiesToLoad.Add("SamAccountName")
                    OUAdminsearcher.PropertiesToLoad.Add("Description")
                    OUAdminsearcher.PropertiesToLoad.Add("Mail")
                    Dim results As SearchResultCollection = OUAdminsearcher.FindAll()
                    For Each resultUser As SearchResult In results
                        samaccountname = resultUser.Properties("SamAccountName")(0)
                        'si le compte est userprog continuer sans traiter
                        If samaccountname = "userprog" Then Continue For

                        Dim rappelComplexite As String = ""
                        Dim type As String = ""
                        Dim sujet As String
                        Dim groupe As String = ""
                        Dim StrategieMDP As DirectoryEntry
                        If Commun.AppartientGroup(samaccountname, "G_SMDPM_Admins") = True Then
                            StrategieMDP = New DirectoryEntry("LDAP://CN=Strategie_MDPM_Admins,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                            type = "adminInfo"
                            sujet = "Expiration de Votre mot de passe d'administration"
                            groupe = "G_SMDPM_Admins"

                        ElseIf Commun.AppartientGroup(samaccountname, "G_SMDPM_Users-Admins") = True Then
                            StrategieMDP = New DirectoryEntry("LDAP://CN=Strategie_MDPM_Users-Admins,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                            type = "users-admin"
                            sujet = "Expiration de Votre mot de passe"
                            groupe = "G_SMDPM_Users-Admins"
                        ElseIf Commun.AppartientGroup(samaccountname, "G_SMDPM_UsersAdm") = True Then
                            StrategieMDP = New DirectoryEntry("LDAP://CN=Strategie_MDPM_Users-Admins,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                            type = "usersAdm"
                            sujet = "Expiration de Votre mot de passe d'administration"
                            groupe = "G_SMDPM_UsersAdm"
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
                        Dim expirationPWD As DateTime = lastSetPWD.AddDays(MaxPWDageJourPolicy - 1)

                        Using userADM As DirectoryEntry = resultUser.GetDirectoryEntry
                            Dim lastSetDateTxt As String = expirationPWD.ToString("dd/MM/yyy")
                            If userADM.Properties("physicalDeliveryOfficeName").Value <> "Expire le : " & lastSetDateTxt And resultUser.Properties("pwdLastSet")(0) <> 0 Then
                                userADM.Properties("physicalDeliveryOfficeName").Value = "Expire le : " & lastSetDateTxt
                                Commun.AppliquerChangement(userADM)
                            End If


                            'If Commun.AccountIsDisabled(userADM) = True And samaccoutname <> "Modele_Admin_Service" Then
                            '    Commun.ReactiveDesactiveCompte(userADM, "active")
                            'End If
                        End Using

                        Dim aujourdhui As DateTime = Format(Date.Now.AddDays(0), "dd/MM/yyyy")
                        Dim demain As DateTime = Format(Date.Now.AddDays(1), "dd/MM/yyyy")
                        Dim expiration30 As DateTime = Format(Date.Now.AddDays(30), "dd/MM/yyyy")
                        Dim expiration7 As DateTime = Format(Date.Now.AddDays(7), "dd/MM/yyyy")
                        Dim expiration3 As DateTime = Format(Date.Now.AddDays(3), "dd/MM/yyyy")
                        Dim expiration2 As DateTime = Format(Date.Now.AddDays(2), "dd/MM/yyyy")

                        Dim Email As String
                        If resultUser.Properties.Contains("Mail") Then
                            Email = resultUser.Properties("Mail")(0)
                        Else
                            Email = Strings.Left(samaccountname, Len(samaccountname) - 3) & "@igbmc.fr"
                        End If



                        Dim ctrl As Boolean = False
                        If type = "adminInfo" Then
                            ctrl = (aujourdhui = expirationPWD Or demain = expirationPWD Or expiration7 = expirationPWD Or expiration3 = expirationPWD Or expiration2 = expirationPWD)
                        Else
                            ctrl = (aujourdhui = expirationPWD Or demain = expirationPWD Or expiration30 = expirationPWD Or expiration7 = expirationPWD Or expiration3 = expirationPWD Or expiration2 = expirationPWD)
                        End If

                        If ctrl = True Then
                            Dim corpMail As String = MailTemplatePasswdExpire(type, samaccountname, expirationPWD, MaxPWDageJourPolicy, historyLenght, nbrChar)
                            Commun.SendEmail("serviceinfo@igbmc.fr", Email, sujet, corpMail)
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
    Shared Function MailTemplatePasswdExpire(ByVal type As String, ByVal samaccountname As String, ByVal expirationPWD As String, ByVal MaxPWDageJourPolicy As String, ByVal historyLenght As String, ByVal nbrChar As String) As String
        Dim corpMail As String
        If type = "adminInfo" Or type = "usersAdm" Then
            Dim rappelComplexite As String = "Votre mot de passe doit etre changé tous les " & MaxPWDageJourPolicy & " jours." & vbCrLf & "Il ne peut pas etre le meme que les " & historyLenght & " précédents." & vbCrLf & "Votre mot de passe doit contenir au moins :" & vbCrLf & vbTab & "- " & nbrChar & " caractères"
            rappelComplexite = rappelComplexite & vbCrLf & vbTab & "- 1 Majuscule" & vbCrLf & vbTab & "- 1 Minuscule" & vbCrLf & vbTab & "- 1 Chiffre" & vbCrLf & vbTab & "- 1 Caractère spécial (non-alphabétique)"

            corpMail = "Le mot de passe de votre compte administrateur (" & samaccountname & ") va expirer le " & expirationPWD & "." & vbCrLf _
                                                        & "Pensez à le changer en ouvrant une session sur un ordinateur du domaine ou AVANT qu'il ne soit expiré en vous connectant ici : https://password.igbmc.fr" & vbCrLf & vbCrLf _
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
    '    Shared Sub ExpirationMDP()

    '        Commun.Journal("Envoie des mails d'expiration de mot de passe", False)

    '        Dim samaccoutname As String
    '        Try

    '            Using OUAdmins As DirectoryEntry = New DirectoryEntry("LDAP://OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
    '                Using OUAdminsearcher As DirectorySearcher = New DirectorySearcher(OUAdmins)
    '                    OUAdminsearcher.Filter = "(&(objectClass=user))"
    '                    OUAdminsearcher.PropertiesToLoad.Add("pwdLastSet")
    '                    OUAdminsearcher.PropertiesToLoad.Add("SamAccountName")
    '                    OUAdminsearcher.PropertiesToLoad.Add("Description")
    '                    For Each resultUser As SearchResult In OUAdminsearcher.FindAll()
    '                        samaccoutname = resultUser.Properties("SamAccountName")(0)
    '                        'si le compte est userprog continuer sans traiter
    '                        If samaccoutname = "userprog" Then Continue For

    '                        Dim rappelComplexite As String = ""
    '                        Dim StrategieMDP As DirectoryEntry
    '                        If Commun.AppartientGroup(samaccoutname, "G_SMDPM_Admins") = True Then
    '                            StrategieMDP = New DirectoryEntry("LDAP://CN=Strategie_MDPM_Admins,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
    '                        ElseIf Commun.AppartientGroup(samaccoutname, "G_SMDPM_UsersAdm") = True Then
    '                            StrategieMDP = New DirectoryEntry("LDAP://CN=Strategie_MDPM_Users-Admins,CN=Password Settings Container,CN=System,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
    '                        Else
    '                            Commun.Journal("ERREUR : Determiner strategie de mot de passe : " & samaccoutname, True)
    '                            GoTo sortie
    '                        End If

    '                        Dim nbrChar As String = StrategieMDP.Properties("msDS-MinimumPasswordLength").Value.ToString
    '                        Dim historyLenght As String = StrategieMDP.Properties("msDS-PasswordHistoryLength").Value.ToString
    '                        Dim MaxPWDageJourPolicy As Integer = ConvertAttribute(StrategieMDP.Properties("msDS-MaximumPasswordAge").Value)

    '                        StrategieMDP.Close()
    '                        StrategieMDP.Dispose()
    '                        StrategieMDP = Nothing

    '                        rappelComplexite = "Votre mot de passe doit etre changé tous les " & MaxPWDageJourPolicy & " jours." & vbCrLf & "Il ne peut pas etre le meme que les " & historyLenght & " précédents." & vbCrLf & "Votre mot de passe doit contenir au moins :" & vbCrLf & vbTab & "- " & nbrChar & " caractères"
    '                        rappelComplexite = rappelComplexite & vbCrLf & vbTab & "- 1 Majuscule" & vbCrLf & vbTab & "- 1 Minuscule" & vbCrLf & vbTab & "- 1 Chiffre" & vbCrLf & vbTab & "- 1 Caractère spécial (non-alphabétique)"

    '                        Dim lastSetPWD As DateTime = Format(New DateTime(1601, 1, 2).AddTicks(resultUser.Properties("pwdLastSet")(0)), "dd/MM/yyyy")
    '                        Dim expirationPWD As DateTime = lastSetPWD.AddDays(MaxPWDageJourPolicy - 1)

    '                        Using userADM As DirectoryEntry = resultUser.GetDirectoryEntry
    '                            Dim lastSetDateTxt As String = expirationPWD.ToString("dd/MM/yyy")
    '                            If userADM.Properties("physicalDeliveryOfficeName").Value <> "Expire le : " & lastSetDateTxt And resultUser.Properties("pwdLastSet")(0) <> 0 Then
    '                                userADM.Properties("physicalDeliveryOfficeName").Value = "Expire le : " & lastSetDateTxt
    '                                Commun.AppliquerChangement(userADM)
    '                            End If


    '                            'If Commun.AccountIsDisabled(userADM) = True And samaccoutname <> "Modele_Admin_Service" Then
    '                            '    Commun.ReactiveDesactiveCompte(userADM, "active")
    '                            'End If
    '                        End Using

    '                        Dim demain As DateTime = Format(Date.Now.AddDays(1), "dd/MM/yyyy")
    '                        Dim SeptjoursAvant As DateTime = Format(Date.Now.AddDays(7), "dd/MM/yyyy")
    '                        If demain = expirationPWD Or SeptjoursAvant = expirationPWD Then

    '                            Dim Email As String = Strings.Left(samaccoutname, Len(samaccoutname) - 3) & "@igbmc.fr"
    '                            Dim corpMail As String = "Le mot de passe de votre compte d'administration (" & samaccoutname & ") va expirer le " & expirationPWD & "." & vbCrLf & "Pensez à le changer en ouvrant une session sur un ordinateur du domaine ou avant qu'il ne soit expiré en vous connectant ici : https://password.igbmc.fr" & vbCrLf & vbCrLf & rappelComplexite & vbCrLf & vbCrLf & vbCrLf & "Le service Informatique" & vbCrLf & "(Email généré automatiquement)"
    '                            Commun.SendEmail("serviceinfo@igbmc.fr", Email, "Expiration de Votre mot de passe de compte administrateur", corpMail)
    '                            Commun.Journal("Mot de passe ADM : mail envoyé a : " & Email, False)
    '                        End If
    '                    Next resultUser

    'sortie:
    '                End Using
    '            End Using
    '            Commun.Journal("Gestion des mots de passe des comptes ADM terminée avec succes", False)
    '        Catch ex As Exception
    '            Commun.Journal("ERREUR : Gestion de password des comptes ADM : " & samaccoutname & " : " & ex.Message, True)
    '        End Try

    '    End Sub
    Shared Function ConvertAttribute(ByVal li As Object) As Integer
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

        'Dim int64Val As ActiveDs.LargeInteger = CType(li, ActiveDs.LargeInteger)

        'If Not int64Val Is Nothing Then
        '    Dim largeInt As System.Int64 = int64Val.HighPart
        '    largeInt = largeInt << 32
        '    largeInt += int64Val.LowPart
        '    Return (Now.AddTicks(-largeInt) - Now).Days
        'End If


    End Function

    ''' <summary>
    ''' Controle qu'un utilisateur est en exception valide
    ''' </summary>
    ''' <param name="login">Login de l'utilisateur a controler.</param>
    ''' <remarks>Retourne "False" s'il n'est pas en exception, et la date (au format texte) s'il est en exception</remarks>

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


    Shared Function ShellSort(ByVal tab1 As String(),
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
    'Shared Function ExceptionCreationForcee(ByVal idUser As String) As Boolean
    '    Dim result = False
    '    Dim tab As String()
    '    Try
    '        Using monStreamReader As StreamReader = New StreamReader("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\ForceCreationDeCompte.txt")

    '            Dim ligne As String
    '            Dim h As Integer = -1

    '            Do
    '                ligne = monStreamReader.ReadLine()
    '                If ligne Is Nothing Then Exit Do
    '                h += 1
    '                ReDim Preserve tab(h)
    '                tab(h) = ligne
    '            Loop

    '        End Using

    '        If Not tab Is Nothing Then
    '            If Array.IndexOf(tab, idUser) > -1 Then
    '                result = True
    '            End If
    '        End If

    '        Return result

    '    Catch ex As Exception
    '        Commun.Journal("ERREUR : Fonction transforme fichier en tableau : " & ex.Message, True)
    '    End Try
    'End Function
    Shared Sub creationFichierParJson()
        FileOpen(2, "c:\temp\eq.txt", OpenMode.Output)
        'Dim data As String = Json.SendJson("", "destinations?is_team=true", "AD", "GET")
        Dim tabResultJson As String()
        Dim lineJson As String = ""


        'FileOpen(2, "c:\temp\listep1.txt", OpenMode.Output) 'fichier du personnel sans accent multi equipe


        'Dim destinations = Json.DeserializeJson(data, "destinations")
        For d = 0 To UBound(destinationsJsonIsTrue)
            Dim IDdest As String = destinationsJsonIsTrue(d).ID

            'If IDdest <> "189" Then Continue For

            Dim Dest_short_name As String = destinationsJsonIsTrue(d).short_name
            Dim dest_name As String = destinationsJsonIsTrue(d).name
            Dim department_id As String = destinationsJsonIsTrue(d).department_id
            Dim nbrUsersDest As Integer = 0
            Dim team_id As String = destinationsJsonIsTrue(d).team_group
            Dim group_id As String = destinationsJsonIsTrue(d).group_id
            Dim entity_id As String = destinationsJsonIsTrue(d).entity_id
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
                Dim dateEntree As Date = persons(p).entrance_date
                Dim dataContract As String = Json.SendJson("", "persons/" & IDuser & "/contracts?current_contract=true", "AD", "GET")

                Dim contracts = Json.DeserializeJson(dataContract, "contracts")
                Dim contractRunning As Boolean = ContratEnCours(contracts)
                'Dim finContrat As String '= DateDeFinDeContract(dataContract)
                'If contractRunning = True Then
                Dim finContrat As String = DateDeFinDeContract(contracts, IDuser)
                'Else
                '    finContrat = DateDeFinDeContract(Json.SendJson("", "persons/" & IDuser & "/contracts", "AD", "GET"))
                'End If

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
                If Commun.TransformeSAMACCOUNTenCN(login) = "" Then
                    compteAD = False
                End If

                'Si l'utilisateur n'a pas de compte dans l'AD et qu'il est entré apres le 1er janvier 2015
                If compteAD = False And (dateMini < dateEntree And contractRunning = True) Then 'Or testException = True)
                    Dim crea As Creation = New Creation

                    'si le login n'est pas defini
                    If login = "" Then
                        login = crea.DetermineLogin(firstname, lastname, IDuser)
                    End If

                    crea.createCompte(lastname, firstname, Dest_short_name, IDuser, login, genre, ld)

                    'quand le compte AD est créé, on defini la variable "compteAD" sur True
                    compteAD = True

                    'ecriture du fichier de creation de compte
                    Dim newUser As String = lastname & "," & firstname & "," & login & ",," & Dest_short_name & ",,,," & ld & "," & IDuser & "," & aliasMail & "," & genre
                    Dim sw As New StreamWriter(RecupDataini.RecupVar("[CheminPartage]") & "\todo\c" & Now.ToString("dd") & "-" & Now.ToString("MM") & "-" & Now.ToString("yy") & "-" & Now.ToString("HH") & "h.txt", True)
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
                        'If loc_building_roomT(l) = "" Then loc_building_roomT(l) = "----"
                        'If loc_phoneT(l) = "" Then loc_phoneT(l) = "----"
                    Next l
                    Dim BatimentUser As String = Join(loc_building_roomT, ";")

                    Dim TelUser As String = Join(loc_phoneT, ";")
                    Erase loc_phoneT, loc_building_roomT

                    Dim organism As String = Organisme(contracts)
                    Dim Nplus1ID As String = ChercherNPlus1deDestination(IDuser, IDdest)



                    '             0                 1               2                   3                   4               5               6                   7               8                  9              10              11              12          13                  14              15                    16
                    lineJson = lastname & "," & firstname & "," & login & "," & Dest_short_name & "," & dest_name & "," & "100" & "," & BatimentUser & "," & TelUser & "," & equipeUser & "," & Nplus1ID & "," & IDuser & "," & aliasMail & "," & ld & "," & organism & "," & finContrat & "," & genre & "," & diffusionPhotoInterne.ToString
                    If tabResultJson Is Nothing Then
                        ReDim Preserve tabResultJson(0)
                    Else
                        ReDim Preserve tabResultJson(UBound(tabResultJson) + 1)
                    End If

                    tabResultJson(UBound(tabResultJson)) = lineJson
                End If
            Next p



        Next d
        FileClose(2)

        'inserer les externes dans le tableau
        'tabResultJson = RecupJsonExterne(tabResultJson)

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
    Shared Sub exceptionCreationCompte()
        Dim lines() As String = IO.File.ReadAllLines("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\ForceCreationDeCompte.txt")
        For Each line As String In lines

            Dim dest_short_name As String = Split(line, ",")(1)

            Dim dataException As String = Json.SendJson("", "persons/" & Split(line, ",")(0), "AD", "GET")
            'Dim persons = Json.DeserializeJson(data, "persons")
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

            If login = "" Then
                login = crea.DetermineLogin(firstname, lastname, IDuser)
            End If
            crea.createCompte(lastname, firstname, dest_short_name, IDuser, login, genre)

        Next
        'nettoyage du fichier ForceCreationDeCompte.txt
        Try
            System.IO.File.WriteAllText("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\ForceCreationDeCompte.txt", "")
        Catch
            Commun.Journal("ERREUR : nettoyage du fichier ForceCreationDeCompte.txt", True)
        End Try
    End Sub
    Shared Function ConvertOrganism(ByVal orgID As String) As String
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
            Dim dataContract As String = Json.SendJson("", "persons/" & IDuser & "/contracts?current_contract=true", "AD", "GET")

            Dim contracts = Json.DeserializeJson(data, "contracts")
            Dim contractRunning As Boolean = ContratEnCours(contracts)

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

    Shared Sub ChangePasswordAccountPrestaImagerie()
        Try
            Dim passwordPrestaImagerieAdm = RandomPassword.Generate(8)
            Dim passwordPrestaImagerieUsr = RandomPassword.Generate(8)
            Using userEntry = New DirectoryEntry("LDAP://" & Commun.TransformeSAMACCOUNTenCN("cs-prestaadm"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                userEntry.Invoke("SetPassword", New Object() {passwordPrestaImagerieAdm})
                Commun.AppliquerChangement(userEntry)
            End Using
            Using userEntry1 = New DirectoryEntry("LDAP://" & Commun.TransformeSAMACCOUNTenCN("cs-prestausr"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
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
    Shared Function TrouverMailPrincipal(ByVal login As String, ByVal courtOuLong As String)
        Try
            Dim mail As String = ""
            Dim objusr As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.TransformeSAMACCOUNTenCN(login), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
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
    Shared Sub EnvoiMailCompteExpireXjour(ByVal j As Integer)

        j += 1
        Dim adresseMail As String = ""
        Dim ctrlMailEnvoye As Boolean = False
        Dim dateNowU As String = Now.Date.ToString("yyyyMMddHHmmss.sZ")

        Try
            Dim dateDeSuppressionPrevueUniversal As String = Now.Date.AddDays(j).ToString("yyyyMMddHHmmss.sZ")
            Dim dateDeSuppressionPrevueTxt As String = Now.Date.AddDays(j).ToString("dd/MM/yyyy")
            'Dim dateDeSuppressionPrevue As String = Strings.Left(CDate(Now.Day & "/" & Now.Month & "/" & Now.Year).AddDays(J).ToString, 10)
            Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursDesactives]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
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
                            Dim mail = MailCloture1(prenom, dateDeSuppressionPrevueTxt, dateDefinDeContrat)

                            Commun.SendEmail("noreply@igbmc.fr", adresseMail & ";Cc:serviceinfo@igbmc.fr", "ARRET DU COMPTE", mail)
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
    Shared Function ChercherNPlus1deDestination(ByVal idUser As String, ByVal idDest As String) As String
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
    Shared Function ContratEnCours(contracts) As Boolean
        Dim dateactu As Date = Convert.ToDateTime(Now.ToString("yyyy-MM-dd"))

        'organismContractName = ""
        ContratEnCours = False

        'Dim contracts = Json.DeserializeJson(data, "contracts")

        If contracts.length > 0 Then ContratEnCours = True

        'For c = 0 To UBound(contracts)
        '    Dim dateStartTxt As String = Strings.Left(contracts(c).start_date, 10)
        '    Dim dateEndTxt As String = Strings.Left(contracts(c).end_date, 10)


        '    Dim dateStart As Date = #1/1/3000#
        '    If dateStartTxt <> Nothing Then
        '        dateStart = Convert.ToDateTime(dateStartTxt)
        '    End If

        '    Dim dateEnd As Date = #1/1/3000#
        '    If dateEndTxt <> Nothing Then
        '        dateEnd = (Convert.ToDateTime(dateEndTxt)).AddDays(1)
        '    End If

        '    If dateStart <= dateactu And dateactu < dateEnd Then
        '        'If organismContract = "2" Or organism = "3" Or organism = "4" Or organism = "5" Or organism = "12" Or organism = "13" Then
        '        ContratEnCours = True
        '        'cas de plusieurs contrats avec des organismes differents


        Return ContratEnCours

    End Function
    Shared Function Organisme(contracts) As String
        Dim result As String = ""


        'If contracts.length > 0 Then ContratEnCours = True

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

    Shared Function DateDeFinDeContract(contracts, id) As String

        Dim result As String = ""
        'Dim dateactu As Date = Convert.ToDateTime(Now.ToString("yyyy-MM-dd"))


        Try
            'Dim ResponseData = New JavaScriptSerializer().Deserialize(Of Object)(Data)
            'Dim contracts = ResponseData("contracts")
            If contracts.length = 1 Then

                If contracts(0).end_date <> "" Then
                    Dim dateEndTxt As String = Strings.Left(contracts(0).end_date, 10)
                    Dim dateTab As String() = Split(dateEndTxt, "-")
                    result = dateTab(2) & "/" & dateTab(1) & "/" & dateTab(0)
                    'result = dateEndTxt.ToString("dd/MM/yyyy")
                End If
            Else
                Dim Data As String = Json.SendJson("", "persons/" & id & "/contracts", "AD", "GET")
                Dim ResponseData = New JavaScriptSerializer().Deserialize(Of Object)(Data)
                contracts = ResponseData("contracts")

                Dim dateEnd As Date = #1/1/1900#
                Dim dateStart As Date = #1/1/1900#
                Dim dateSelected As Date = #1/1/1900#

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

                    If dateEnd > dateSelected And dateStart <= Now Then
                        dateSelected = dateEnd
                    End If
                Next c

                result = dateSelected.ToString("dd/MM/yyyy")
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : DateDeFinDeContract : EmployeeID : " & id & " : " & ex.Message, True)
        End Try


        Return result

    End Function
    Shared Sub controlDoublonUIDNumber()
        Using Ldap As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Using searcher As DirectorySearcher = New DirectorySearcher(Ldap)
                searcher.Filter = "(&(objectClass=posixAccount)(uidNumber=*))"
                searcher.PropertiesToLoad.Add("uidNumber")
                searcher.PropertiesToLoad.Add("SAMAccountName")
                Dim resultSearcher As SearchResultCollection = searcher.FindAll()
                Dim tmp As String()
                Dim i = -1
                For Each result As SearchResult In resultSearcher
                    i += 1
                    Dim uid As String = result.Properties("uidNumber")(0).ToString
                    If i <> 0 Then
                        If Array.IndexOf(tmp, uid) = -1 Then

                            tmp.Add(uid)
                        Else
                            If i <> Array.IndexOf(tmp, uid) Then
                                Using results2 As SearchResultCollection = Commun.SearchFilterAll(Ldap, "(&(objectClass=posixAccount)(uidNumber=" & uid & "))", SearchScope.Subtree, "cn")
                                    Dim tmp2 As String()
                                    For Each result2 In results2
                                        tmp2.Add(result2.properties("cn")(0))


                                    Next
                                    MsgBox(Join(tmp2, " , ") & " : " & uid & vbCrLf & Commun.UIDNumberMini)
                                    Erase tmp2
                                End Using
                            End If
                        End If
                    Else
                        tmp.Add(uid)
                    End If
                Next

            End Using
        End Using
    End Sub
    Shared Sub UpdateFichierHistoAlias()

        Commun.Journal("Mise a jour du fichier d'historique des alias", False)

        'PARTIE ALIAS
        Dim tabAlias(,) As String = Commun.CreateTabHistoAliasLogin("alias")

        Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Using searcher As DirectorySearcher = New DirectorySearcher(objAD)
                searcher.Filter = "(&(proxyAddresses=*)(EmployeeID=*))"
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

                            If objADwithSMTPAddresses.Properties.Contains("EmployeeID") Then
                                userID = objADwithSMTPAddresses.Properties("EmployeeID").Value
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

                                    'Dim ctrlAlias As Boolean = Commun.ctrlAliasDispo(aliasMail, userID)
                                    Dim adresseDansLeFichier As Boolean = False
                                    'Dim text As String = File.ReadAllText("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\HistoAlias.txt")
                                    'Dim index As Integer = text.IndexOf(aliasMail & "," & userID)
                                    'If index >= 0 Then
                                    '    adresseDansLeFichier = True
                                    'End If

                                    'If adresseDansLeFichier = False Then
                                    '    Dim index1 As Integer = text.IndexOf(aliasMail & ",0")
                                    '    If index1 >= 0 Then
                                    '        adresseDansLeFichier = True
                                    '    End If
                                    'End If

                                    'si l'alias ne fait pas deja partie de la liste d'adresse de l'utilisateur et qu'il n'est pas encore dans le fichier
                                    'If InStr(aliasPrecedent, "§" & aliasMail & "§") = 0 And adresseDansLeFichier = False Then
                                    'Using sw As New StreamWriter("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\HistoAlias.txt", True)
                                    '    'on ecrit l'alias dans le fichier
                                    '    sw.WriteLine(aliasMail & "," & userID)
                                    'End Using
                                    'End If
                                    aliasPrecedent += "'" & aliasMail & "'"
                                End If
                            Next


                            If aliasPrecedent = "" Then Continue For

                            Dim listAlias = Replace(aliasPrecedent, "''", "','")
                            'listAlias = Replace(listAlias, "§", "'")

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
        Using objAD As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr")
            Using searcher As DirectorySearcher = New DirectorySearcher(objAD)
                searcher.Filter = "(&(SAMAccountName=*)(!(SAMAccountName=*$*)))"
                searcher.PageSize = 2000
                Dim results As SearchResultCollection = searcher.FindAll
                For Each result As SearchResult In results
                    'si l'objet est dans l'OU NAP on continue avec l'objet suivant
                    If InStr(result.Path, "OU=Nap,DC=igbmc,DC=u-strasbg,DC=fr") <> 0 Then Continue For
                    Using objADSAMAccount As DirectoryEntry = result.GetDirectoryEntry
                        Dim login As String = objADSAMAccount.Properties("SAMAccountName").Value
                        Dim userID As String = ""
                        If objADSAMAccount.Properties.Contains("EmployeeID") Then
                            userID = objADSAMAccount.Properties("EmployeeID").Value
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
    Shared Function EstPair(n As Long) As Boolean
        EstPair = (n Mod 2) = 0
    End Function


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
    Shared Function ListeDiffusion(ByVal IDuser As String) As String
        Dim dataLD As String = Json.SendJson("", "persons/" & IDuser & "/mailing_lists", "AD", "GET")
        Dim ld As String = "other"

        Dim aaa = InStr(dataLD, """name"":""postdoc""")
        If InStr(dataLD, """name"": ""postdoc""") > 0 Then
            ld = "postdoc"
        End If

        If InStr(dataLD, """name"": ""phd""") > 0 Then
            ld = "phd"
        End If
        Return ld
    End Function
    'Shared Sub GestionOuOut()

    '    Using users As SearchResultCollection = Commun.SearchFilterAll(RecupDataini.RecupVar("[OUUtilisateursSortis]"), "(&(objectClass=user))", SearchScope.OneLevel)

    '        For Each user As SearchResult In users
    '            Dim differenceModification As Integer = DateDiff("H", user.Properties("whenChanged")(0), Now)
    '            Dim cn As String = Replace(user.Path, "LDAP://", "")
    '            Dim login As String = Commun.FindAttribut(cn, "SAMAccountName")
    '            Dim employeeID As String = Commun.FindAttribut(cn, "EmployeeID")
    '            Dim nomFichierPST As String = login & "-" & employeeID & "-IGBMC.pst"
    '            Dim testFichierPST As Boolean = File.Exists(dossierArchivePST & nomFichierPST)
    '            If testFichierPST = True And differenceModification >= 24 Then
    '                PWSDisableMailbox(login)
    '            End If
    '        Next
    '    End Using

    'End Sub
    Shared Function PWSDisableMailbox(ByVal aliasMail As String) As Boolean
        Dim result As Boolean = False
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
            pCredential = DirectCast(Nothing, PSCredential)
            'pCredential = New PSCredential("igbmc\steph", CreateSecurePasswordString("aaaaaa"))

            '-- set connection info
            pConnectionInfo = New WSManConnectionInfo(New Uri("http://serv-mbx12.igbmc.u-strasbg.fr/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", pCredential)
            pConnectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos
            '-- create remote runspace
            pRunspace = RunspaceFactory.CreateRunspace(pConnectionInfo)
            pRunspace.Open()

            '-- create powershell
            pShell = PowerShell.Create
            pShell.Runspace = pRunspace

            '-- create command
            pCommand = New PSCommand
            With pCommand
                .AddCommand("Disable-Mailbox")
                .AddParameter("identity", aliasMail)
                .AddParameter("confirm", False)
            End With

            '-- add command to powershell
            pShell.Commands = pCommand

            '-- invoke the powershell
            pResult = pShell.Invoke
            If pResult.Count = 1 Then result = True
            pRunspace.Close()
            pRunspace.Dispose()

        Catch e As Exception

        End Try
        Return result
    End Function

End Class