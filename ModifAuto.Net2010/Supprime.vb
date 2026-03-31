Imports Microsoft.VisualBasic
Imports System.DirectoryServices
Imports System.IO


Public Class Supprime

    Shared Sub SupprimCompte(ByVal userAD As DirectoryEntry)


        Dim login As String = userAD.Properties("SAMAccountName").Value
        'Dim path As String = result.Path
        Try

            'Si l'utilisateur a un EmployeeID, on le recupere, s'il n'en a pas (utilisateurs Externe), userID reste defini sur nothing
            Dim userID As String = Nothing
            If userAD.Properties.Contains("EmployeeID") Then
                userID = userAD.Properties("EmployeeID").Value
            Else

            End If
            Dim archiveEnabled As Boolean = False
            If userAD.Properties.Contains("msExchArchiveGUID") Then archiveEnabled = True

            'Dim dateSuppressionTxt As String = result.Properties("accountDeletionDate")(0)
            Dim dateSuppression As Date = userAD.Properties("accountDeletionDT").Value
            'Dim dateSuppression As Date = Date.ParseExact(dateSuppressionTxt, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)

            If Now >= dateSuppression Then

                If userID <> "" Then
                    Pws.CommandePWSCreatePSTMailbox(login, userID)
                    If archiveEnabled = True Then
                        Pws.CommandePWSCreatePSTMailbox(login, userID, True)
                    End If
                End If
                removeAllGroup(login)

                CompteSorti(userAD)
                'result.GetDirectoryEntry.DeleteTree()

                Commun.Journal("SupprimCompte : Suppression Compte AD : " & login)

                
                'si userID a une valeur (utilisateur enregistré dans la base du perso) on supprime son mail dans la BDP
                If Not userID Is Nothing Then
                    Try

                        Json.SendJson("", "persons/" & userID & "/email", "AD", "DELETE")
                        Commun.Journal("SupprimCompte : Suppression IGBMCServices Réussi : " & login)
                    Catch ex As Exception
                        Commun.Journal("ERREUR : SupprimCompte : Suppression IGBMCServices : " & login & " : " & ex.Message, True)
                    End Try
                End If


                'Suppression du compte ADM
                Dim cheminLdapCompteAdm As String = Commun.TransformeSAMACCOUNTenCN(login & "adm")
                If cheminLdapCompteAdm <> "" Then
                    Using userAdm As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(cheminLdapCompteAdm), Commun.admin, Commun.passwd, auth)
                        Try
                            userAdm.DeleteTree()
                            Commun.Journal("SupprimCompte : Suppression Compte AD Adm : " & login)
                        Catch ex As Exception
                            Commun.Journal("ERREUR : SupprimCompte : Suppression Compte AD Adm : " & login & " : " & ex.Message, True)
                        End Try
                    End Using
                End If
            End If

        Catch ex As Exception
            Commun.Journal("ERREUR : SupprimCompte : Suppression Compte AD : " & login & " : " & ex.Message, True)
        End Try

    End Sub
    Shared Sub SupprimCompteADM()
        ' OU contenant les comptes admin
        Dim adminOU As New DirectoryEntry(
        "LDAP://" & Commun.LdapPath("OU=Admins,DC=igbmc,DC=u-strasbg,DC=fr"),
        Commun.admin, Commun.passwd, auth
    )

        ' OU contenant les comptes utilisateurs
        Dim usersOU As New DirectoryEntry(
        "LDAP://" & Commun.LdapPath("OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr"),
        Commun.admin, Commun.passwd, auth
    )

        ' Recherche uniquement les comptes utilisateurs (pas les groupes)
        Dim adminSearcher As New DirectorySearcher(adminOU)
        adminSearcher.Filter = "(objectClass=user)"
        adminSearcher.PropertiesToLoad.Add("sAMAccountName")
        adminSearcher.PropertiesToLoad.Add("cn")

        Dim adminResults As SearchResultCollection = adminSearcher.FindAll()

        For Each admin As SearchResult In adminResults

            If Not admin.Properties.Contains("sAMAccountName") Then Continue For

            Dim adminLogin As String = admin.Properties("sAMAccountName")(0).ToString()
            If adminLogin.Equals("userprog", StringComparison.OrdinalIgnoreCase) Then Continue For
            ' Exclure les comptes modèles
            Dim cn As String = admin.Properties("cn")(0).ToString()
            If cn.StartsWith("Modele", StringComparison.OrdinalIgnoreCase) Then Continue For

            ' Déduction du compte utilisateur (retire les 3 derniers caractères "adm")
            If adminLogin.Length <= 3 Then Continue For
            Dim userLogin As String = adminLogin.Substring(0, adminLogin.Length - 3)

            ' Recherche du compte utilisateur correspondant
            Dim userSearcher As New DirectorySearcher(usersOU)
            userSearcher.Filter = "(sAMAccountName=" & userLogin & ")"

            Dim userFound As SearchResult = userSearcher.FindOne()

            ' Si le compte utilisateur n'existe pas, supprimer le compte admin
            If userFound Is Nothing Then
                Dim adminEntry As DirectoryEntry = admin.GetDirectoryEntry()
                Console.WriteLine("Suppression du compte admin orphelin : " & adminLogin)
                adminEntry.DeleteTree()
            End If

        Next
    End Sub
    Shared Sub SupprimeMailbox()

        Commun.Journal("Creation des archives PST", False)

        'Dim dateSuppressionMB As String = Now.AddDays(-2).ToString("dd/MM/yyyy")
        Using OUDisable As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateursSortis), Commun.admin, Commun.passwd, auth)
            Using dirSearcher As DirectorySearcher = New DirectorySearcher(OUDisable)
                dirSearcher.Filter = "(&(objectClass=user)(homeMDB=*))"
                dirSearcher.SearchScope = SearchScope.OneLevel
                Dim results As SearchResultCollection = dirSearcher.FindAll()

                For Each result As SearchResult In results
                    Using userAD As DirectoryEntry = result.GetDirectoryEntry
                        Dim login As String = userAD.Properties("SAMAccountName").Value
                        Try
                            'Si l'utilisateur a un EmployeeID, on le recupere, s'il n'en a pas (utilisateurs Externe), userID reste defini sur nothing
                            Dim userID As String = Nothing
                            If userAD.Properties.Contains("EmployeeID") Then
                                userID = userAD.Properties("EmployeeID").Value
                            End If

                            Dim archiveEnabled As Boolean = False
                            If userAD.Properties.Contains("msExchArchiveGUID") Then archiveEnabled = True

                            Dim statusPSTMailbox As String = Pws.ControleArchiveCreee(login, userID)
                            Dim statusPSTArchive As String
                            If archiveEnabled = True Then statusPSTArchive = Pws.ControleArchiveCreee(login, userID, True)

                            'Controle si le PST a été créé
                            If statusPSTMailbox = "Completed" Then
                                If archiveEnabled = True Then
                                    If statusPSTArchive = "Completed" Then
                                        Pws.DeleteExportRequest(login, userID)
                                        Pws.DeleteExportRequest(login, userID, True)
                                        Pws.commandePWSDisableMailbox(login)
                                    End If
                                Else
                                    Pws.DeleteExportRequest(login, userID)
                                    Pws.commandePWSDisableMailbox(login)
                                End If
                            End If

                            If statusPSTMailbox = "Failed" Then
                                DeleteIncompletePSTFile(login, userID)
                                Pws.CommandePWSCreatePSTMailbox(login, userID)
                            End If

                            If archiveEnabled = True Then
                                If statusPSTArchive = "Failed" Then
                                    DeleteIncompletePSTFile(login, userID, True)
                                    Pws.CommandePWSCreatePSTMailbox(login, userID, True)
                                End If
                                If statusPSTArchive Is Nothing Then
                                    Pws.CommandePWSCreatePSTMailbox(login, userID)
                                End If
                            End If

                            If statusPSTMailbox Is Nothing Then
                                Pws.CommandePWSCreatePSTMailbox(login, userID)
                            End If


                        Catch ex As Exception
                            Commun.Journal("ERREUR : SupprimeMailbox : Suppression de la boite mail : " & login & " : " & ex.Message, True)
                        End Try
                    End Using
                Next
            End Using
        End Using


    End Sub
    Shared Sub DeleteOldPST()
        Commun.Journal("Suppression des anciennes archives PST", False)

        ' Spécifiez l'extension des fichiers PST
        Dim extensionPST As String = ".pst"

        ' Obtenez tous les fichiers PST dans le dossier réseau
        Dim fichiersPST() As String = Directory.GetFiles(dossierArchivePST, "*" & extensionPST)

        ' Définissez la période à partir de laquelle les fichiers doivent être supprimés (3 semaines)
        Dim jourSuppressionPST As Integer = ini.ReadValue("MODIFAUTO", "jourSuppressionPST", 21)
        Dim periodeSuppression As TimeSpan = TimeSpan.FromDays(jourSuppressionPST)

        ' Obtenez la date actuelle
        Dim dateActuelle As DateTime = DateTime.Now

        ' Parcourez tous les fichiers PST
        For Each fichier As String In fichiersPST
            ' Obtenez la date de création du fichier
            Dim dateCreation As DateTime = File.GetCreationTime(fichier)

            ' Calculez la différence de temps entre la date actuelle et la date de création
            Dim differenceTemps As TimeSpan = dateActuelle - dateCreation

            ' Vérifiez si la différence de temps est supérieure à la période de suppression
            If differenceTemps > periodeSuppression Then
                ' Supprimez le fichier PST
                Try
                    File.Delete(fichier)
                    Commun.Journal("DeleteOldPST : Suppression du fichier PST : " & fichier)
                Catch ex As Exception
                    Commun.Journal("ERREUR : DeleteOldPST : Suppression du fichier PST : " & fichier & " : " & ex.Message, True)
                End Try
            End If
        Next


    End Sub

    Shared Sub DeleteIncompletePSTFile(ByVal login As String, ByVal userID As String, Optional ByVal archive As Boolean = False)
        Dim PSTFileName As String
        If archive = False Then
            PSTFileName = login & "-" & userID & "-IGBMC.pst"
        Else
            PSTFileName = login & "(Archive)-" & userID & "-IGBMC.pst"
        End If

        Try
            If File.Exists(dossierArchivePST & PSTFileName) Then
                My.Computer.FileSystem.DeleteFile(dossierArchivePST & PSTFileName)
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : DeleteIncompletePSTFile : Suppression du fichier PST incomplet : " & PSTFileName & " : " & ex.Message, True)
        End Try
    End Sub


    Shared Sub CompteSorti(ByVal DirEntry As DirectoryEntry)
        Try
            Dim login As String = DirEntry.Properties("SAMAccountName").Value

            Commun.SetADLDAPProperty(DirEntry, "Description", "Supprimé le: " & Strings.Left(CStr(Now), 10), False)

            Commun.SetADLDAPProperty(DirEntry, "Comment", " : Supprimé le: " & Strings.Left(CStr(Now), 10) & " (Autocompte)", True)
            Commun.SetADLDAPProperty(DirEntry, "Manager", "")
            Commun.SetADLDAPProperty(DirEntry, "ipPhone", "")
            Commun.SetADLDAPProperty(DirEntry, "physicalDeliveryOfficeName", "")
            Commun.SetADLDAPProperty(DirEntry, "telephoneNumber", "")
            Commun.SetADLDAPProperty(DirEntry, "departmentNumber", "")
            Commun.SetADLDAPProperty(DirEntry, "department", "")
            Commun.SetADLDAPProperty(DirEntry, "company", "")
            DirEntry.Properties("accountDeletionDate").Value = Nothing
            'DirEntry.Properties("accountDeletionDT").Value = Nothing

            Commun.AppliquerChangement(DirEntry)

            Commun.ReactiveDesactiveCompte(login, "desactive")

            Using ouOut As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateursSortis), Commun.admin, Commun.passwd, auth)
                DirEntry.MoveTo(ouOut)
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Shared Sub ClearAttributes(ByVal chemin As String)
        If Directory.Exists(chemin) Then
            Dim dirFolder As DirectoryInfo = New DirectoryInfo(chemin)
            dirFolder.Attributes = FileAttributes.Normal
            Dim files As String() = Directory.GetFiles(chemin)
            For Each File As String In files
                System.IO.File.SetAttributes(File, FileAttributes.Normal)
            Next
            Dim subDirs As String() = Directory.GetDirectories(chemin)
            For Each dir As String In subDirs
                ClearAttributes(dir)
            Next
        End If
    End Sub
    Shared Sub removeAllGroup(ByVal SamAccount As String)
        Dim i As Integer = -1
        Try
            'Recupération de l'attribut Member pour le mettre dans le tableau des resultats
            Using AD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
                Using searcherGroup As DirectorySearcher = New DirectorySearcher(AD)
                    searcherGroup.Filter = "(&(objectClass=user) (SAMAccountName=" & SamAccount & "))"
                    searcherGroup.PropertiesToLoad.Add("Member")
                    searcherGroup.PropertiesToLoad.Add("objectSid")
                    Dim resultGroup As SearchResult = searcherGroup.FindOne()
                    If Not resultGroup Is Nothing Then
                        Dim Group As DirectoryEntry = resultGroup.GetDirectoryEntry
                        For Each unMembre In Group.Properties("memberof")
                            Commun.AddRemoveADGroup(SamAccount, unMembre.ToString, "Remove")
                        Next unMembre
                    End If
                End Using
            End Using

            gestion.removeAdopte(SamAccount)
        Catch ex As Exception

        End Try
    End Sub
End Class
