Imports Microsoft.VisualBasic

Public Class Supprime
    Sub SupprimCompte()

        Using OUDisable As DirectoryEntry = New DirectoryEntry("LDAP://DC=igbmc,DC=u-strasbg,DC=fr", admin, passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Using dirSearcher As DirectorySearcher = New DirectorySearcher(OUDisable)
                dirSearcher.PropertiesToLoad.Add("EmployeeID")
                dirSearcher.PropertiesToLoad.Add("accountDeletionDate")
                dirSearcher.PropertiesToLoad.Add("SAMAccountName")
                dirSearcher.PropertiesToLoad.Add("DisplayName")
                dirSearcher.PropertiesToLoad.Add("homeMDB")
                dirSearcher.Filter = "(&(objectClass=user)(accountDeletionDate=*))"
                dirSearcher.SearchScope = SearchScope.Subtree
                Dim results As SearchResultCollection = dirSearcher.FindAll()

                For Each result As SearchResult In results
                    Dim login As String = result.Properties("SAMAccountName")(0)
                    Dim path As String = result.Path
                    Try

                        'Si l'utilisateur a un EmployeeID, on le recupere, s'il n'en a pas (utilisateurs Externe), userID reste defini sur nothing
                        Dim userID As String = Nothing
                        If result.Properties.Contains("EmployeeID") Then
                            userID = result.Properties("EmployeeID")(0)
                        Else

                        End If

                        Dim dateSuppressionTxt As String = result.Properties("accountDeletionDate")(0)
                        Dim dateSuppression As Date = Date.ParseExact(dateSuppressionTxt, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)

                        If Format(Now, "dd/MM/yyyy") >= dateSuppression Then

                            If userID <> "" Then
                                CommandePWSCreatePSTMailbox(login, userID)
                            End If
                            removeAllGroup(login)

                            CompteSorti(result.GetDirectoryEntry)
                            'result.GetDirectoryEntry.DeleteTree()

                            Commun.Journal("SupprimCompte : Suppression Compte AD : " & login)



                            Try
                                NettoyageDossierRedirectionEtProfil(login)
                                Commun.Journal("SupprimCompte : Suppression des dossiers profils/redirections : " & login)
                            Catch ex As Exception
                                Commun.Journal("ERREUR : SupprimCompte : Suppression des dossiers profils/redirections : " & login & " : " & ex.Message, True)
                            End Try

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
                                Using userAdm As DirectoryEntry = New DirectoryEntry("LDAP://" & cheminLdapCompteAdm, admin, passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                                    Try
                                        userAdm.DeleteTree()
                                        Commun.Journal("SupprimCompte : Suppression Compte AD Adm : " & login)
                                    Catch ex As Exception
                                        Commun.Journal("ERREUR : SupprimCompte : Suppression Compte AD Adm : " & login & " : " & ex.Message, True)
                                    End Try
                                    NettoyageDossierRedirectionEtProfil(login & "adm")
                                End Using
                            End If
                        End If

                    Catch ex As Exception
                        Commun.Journal("ERREUR : SupprimCompte : Suppression Compte AD : " & login & " : " & ex.Message, True)
                    End Try


                Next

            End Using
        End Using
    End Sub

    Sub SupprimeMailbox()
        'Dim dateSuppressionMB As String = Now.AddDays(-2).ToString("dd/MM/yyyy")
        Using OUDisable As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursSortis]"), admin, passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Using dirSearcher As DirectorySearcher = New DirectorySearcher(OUDisable)
                dirSearcher.PropertiesToLoad.Add("accountDeletionDate")
                dirSearcher.PropertiesToLoad.Add("SAMAccountName")
                dirSearcher.PropertiesToLoad.Add("homeMDB")
                dirSearcher.PropertiesToLoad.Add("EmployeeID")
                dirSearcher.Filter = "(&(objectClass=user))"
                dirSearcher.SearchScope = SearchScope.OneLevel
                Dim results As SearchResultCollection = dirSearcher.FindAll()

                For Each result As SearchResult In results
                    Dim login As String = result.Properties("SAMAccountName")(0)
                    Try

                        'Si l'utilisateur a un EmployeeID, on le recupere, s'il n'en a pas (utilisateurs Externe), userID reste defini sur nothing
                        Dim userID As String = Nothing
                        If result.Properties.Contains("EmployeeID") Then
                            userID = result.Properties("EmployeeID")(0)
                        Else

                        End If

                        Dim dateSuppressionTxt As String = result.Properties("accountDeletionDate")(0)
                        Dim dateSuppression As Date = Date.ParseExact(dateSuppressionTxt, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)

                        Dim aaa As Integer = DateDiff("d", dateSuppressionTxt, Now())
                        If aaa >= 1 And result.Properties.Contains("homeMDB") And File.Exists("\\Serv-mbx4\pst\" & login & "-" & userID & "-IGBMC.pst") Then
                            commandePWSDisableMailbox(login)
                        End If
                    Catch ex As Exception

                    End Try

                Next
            End Using
        End Using


    End Sub

    Sub CompteSorti(ByVal DirEntry As DirectoryEntry)
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
            Commun.SetADLDAPProperty(DirEntry, "accountDeletionDate", "")

            DirEntry.CommitChanges()

            Commun.ReactiveDesactiveCompte(login, "desactive")

            Using ouOut As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursSortis]"), admin, passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                DirEntry.MoveTo(ouOut)
            End Using
        Catch ex As Exception

        End Try

    End Sub
End Class
