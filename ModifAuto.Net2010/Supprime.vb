Imports Microsoft.VisualBasic
Imports System.DirectoryServices
Imports System.IO


Public Class Supprime
    Private Shared ReadOnly jourSuppressionPST As Integer = ini.ReadValue("MODIFAUTO", "jourSuppressionPST", 21)

    Shared Sub TraiterSortieCompte(ByVal userAD As DirectoryEntry)
        Dim login As String = userAD.Properties("sAMAccountName").Value

        Try
            Dim userID As String = Nothing
            If userAD.Properties.Contains("employeeID") Then
                userID = userAD.Properties("employeeID").Value
            End If

            Dim dateSuppression As Date = userAD.Properties("accountDeletionDT").Value

            If Now >= dateSuppression Then
                RetirerGroupesCompte(login)
                FinaliserSortieCompteDansAD(userAD)

                Commun.Journal("TraiterSortieCompte : Suppression Compte AD : " & login)

                If Not userID Is Nothing Then
                    Try
                        Json.SendJson("", "persons/" & userID & "/email", "AD", "DELETE")
                        Commun.Journal("TraiterSortieCompte : Suppression IGBMCServices Réussi : " & login)
                    Catch ex As Exception
                        Commun.Journal("ERREUR : TraiterSortieCompte : Suppression IGBMCServices : " & login & " : " & ex.Message, True)
                    End Try
                End If
            End If

        Catch ex As Exception
            Commun.Journal("ERREUR : TraiterSortieCompte : Suppression Compte AD : " & login & " : " & ex.Message, True)
        End Try
    End Sub

    Shared Sub GererArchivesPSTEtMailboxesSorties()
        Commun.Journal("Gestion des archives PST et des mailboxes des comptes sortis", False)

        Using ouSortis As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateursSortis), Nothing, Nothing, auth)
            Using searcher As DirectorySearcher = New DirectorySearcher(ouSortis)
                searcher.Filter = "(&(objectClass=user)(homeMDB=*))"
                searcher.SearchScope = SearchScope.OneLevel

                For Each result As SearchResult In searcher.FindAll()
                    Using userAD As DirectoryEntry = result.GetDirectoryEntry()
                        Dim login As String = userAD.Properties("sAMAccountName").Value

                        Try
                            Dim userID As String = Nothing
                            If userAD.Properties.Contains("employeeID") Then
                                userID = userAD.Properties("employeeID").Value
                            End If

                            Dim archiveEnabled As Boolean = userAD.Properties.Contains("msExchArchiveGUID")
                            Dim statusPSTMailbox As String = Pws.ControlerEtNettoyerExportPST(login, userID)
                            Dim statusPSTArchive As String = Nothing
                            If archiveEnabled Then
                                statusPSTArchive = Pws.ControlerEtNettoyerExportPST(login, userID, True)
                            End If

                            Dim exportsTermines As Boolean =
                                statusPSTMailbox = "Completed" AndAlso
                                (Not archiveEnabled OrElse statusPSTArchive = "Completed")

                            If exportsTermines Then
                                Pws.DeleteExportRequest(login, userID)
                                If archiveEnabled Then
                                    Pws.DeleteExportRequest(login, userID, True)
                                End If

                                Commun.Journal("PST termine : suppression de la boite mail : " & login, False)
                                Pws.DesactiverBoiteMailExchange(login)
                                Continue For
                            End If

                            If statusPSTMailbox = "Failed" Then
                                DeleteIncompletePSTFile(login, userID)
                                Commun.Journal("Relance creation PST principal : " & login & " : demande precedente en echec", False)
                                Pws.CreerDemandeExportPST(login, userID)
                            ElseIf statusPSTMailbox Is Nothing Then
                                Commun.Journal("Creation PST principal : " & login & " : demande absente", False)
                                Pws.CreerDemandeExportPST(login, userID)
                            End If

                            If archiveEnabled Then
                                If statusPSTArchive = "Failed" Then
                                    DeleteIncompletePSTFile(login, userID, True)
                                    Commun.Journal("Relance creation PST archive : " & login & " : demande precedente en echec", False)
                                    Pws.CreerDemandeExportPST(login, userID, True)
                                ElseIf statusPSTArchive Is Nothing Then
                                    Commun.Journal("Creation PST archive : " & login & " : demande absente", False)
                                    Pws.CreerDemandeExportPST(login, userID, True)
                                End If
                            End If

                        Catch ex As Exception
                            Commun.Journal("ERREUR : GererArchivesPSTEtMailboxesSorties : Gestion de la boite mail : " & login & " : " & ex.Message, True)
                        End Try
                    End Using
                Next
            End Using
        End Using
    End Sub
    Shared Sub DeleteOldPST()
        Commun.Journal("Suppression des anciennes archives PST", False)

        IdentitePartagePST.Executer(
            Sub()
                Dim fichiersPST() As String = Directory.GetFiles(dossierArchivePST, "*.pst")

                Dim periodeSuppression As TimeSpan = TimeSpan.FromDays(jourSuppressionPST)
                Dim dateActuelle As DateTime = DateTime.Now

                For Each fichier As String In fichiersPST
                    If dateActuelle - File.GetCreationTime(fichier) > periodeSuppression Then
                        Try
                            File.Delete(fichier)
                            Commun.Journal("DeleteOldPST : Suppression du fichier PST : " & fichier)
                        Catch ex As Exception
                            Commun.Journal("ERREUR : DeleteOldPST : Suppression du fichier PST : " & fichier & " : " & ex.Message, True)
                        End Try
                    End If
                Next
            End Sub)
    End Sub

    Shared Sub DeleteIncompletePSTFile(ByVal login As String, ByVal userID As String, Optional ByVal archive As Boolean = False)
        Dim PSTFileName As String
        If archive = False Then
            PSTFileName = login & "-" & userID & "-IGBMC.pst"
        Else
            PSTFileName = login & "(Archive)-" & userID & "-IGBMC.pst"
        End If

        Try
            IdentitePartagePST.Executer(
                Sub()
                    Dim cheminPST As String = dossierArchivePST & PSTFileName

                    If File.Exists(cheminPST) Then
                        File.Delete(cheminPST)
                    End If
                End Sub)
        Catch ex As Exception
            Commun.Journal("ERREUR : DeleteIncompletePSTFile : Suppression du fichier PST incomplet : " & PSTFileName & " : " & ex.Message, True)
        End Try
    End Sub


    Shared Sub FinaliserSortieCompteDansAD(ByVal DirEntry As DirectoryEntry)
        Dim login As String = ""
        Try
            login = DirEntry.Properties("sAMAccountName").Value.ToString()

            Commun.SetADLDAPProperty(DirEntry, "Description", "Supprimé le: " & Strings.Left(CStr(Now), 10), False)

            Commun.SetADLDAPProperty(DirEntry, "Comment", " : Supprimé le: " & Strings.Left(CStr(Now), 10) & " (Autocompte)", True)
            Commun.SetADLDAPProperty(DirEntry, "manager", "")
            Commun.SetADLDAPProperty(DirEntry, "ipPhone", "")
            Commun.SetADLDAPProperty(DirEntry, "physicalDeliveryOfficeName", "")
            Commun.SetADLDAPProperty(DirEntry, "telephoneNumber", "")
            Commun.SetADLDAPProperty(DirEntry, "departmentNumber", "")
            Commun.SetADLDAPProperty(DirEntry, "department", "")
            Commun.SetADLDAPProperty(DirEntry, "company", "")
            'DirEntry.Properties("accountDeletionDate").Value = Nothing
            'DirEntry.Properties("accountDeletionDT").Value = Nothing

            Commun.AppliquerChangement(DirEntry)

            Commun.ReactiveDesactiveCompte(login, "desactive")

            Using ouOut As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateursSortis), Nothing, Nothing, auth)
                DirEntry.MoveTo(ouOut)
            End Using
        Catch ex As Exception
            Commun.Journal("ERREUR : FinaliserSortieCompteDansAD : " & login & " : " & ex.Message, True)
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
    Shared Sub RetirerGroupesCompte(ByVal SamAccount As String)
        Dim i As Integer = -1
        Try
            'Recupération de l'attribut Member pour le mettre dans le tableau des resultats
            Using AD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("DC=igbmc,DC=u-strasbg,DC=fr"), Nothing, Nothing, auth)
                Using searcherGroup As DirectorySearcher = New DirectorySearcher(AD)
                    searcherGroup.Filter = "(&(objectClass=user) (sAMAccountName=" & SamAccount & "))"
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
            Commun.Journal("ERREUR : RetirerGroupesCompte : " & SamAccount & " : " & ex.Message, True)
        End Try
    End Sub
End Class
