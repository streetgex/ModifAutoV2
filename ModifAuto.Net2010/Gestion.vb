Imports System.DirectoryServices

Imports System.IO


Public Class Gestion
    Shared Sub CompleterDatesContratManquantesComptesDesactivesEtSortis()
        Commun.Journal("Recherche des dates de fin de contrat manquantes dans les comptes desactives et sortis", False)

        CompleterDatesContratManquantesDansOU(OUUtilisateursDesactives, "OU comptes desactives")
        CompleterDatesContratManquantesDansOU(OUUtilisateursSortis, "OU comptes sortis")
    End Sub

    Private Shared Sub CompleterDatesContratManquantesDansOU(ByVal cheminOu As String, ByVal libelleOu As String)
        Try
            Using ouEntry As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(cheminOu), Commun.admin, Commun.passwd, auth)
                Using searcher As DirectorySearcher = New DirectorySearcher(ouEntry)
                    searcher.Filter = "(&(objectClass=user)(employeeID=*)(!(extensionAttribute1=*)))"
                    searcher.SearchScope = SearchScope.OneLevel
                    searcher.PageSize = 5000
                    searcher.PropertiesToLoad.Add("employeeID")
                    searcher.PropertiesToLoad.Add("sAMAccountName")
                    searcher.PropertiesToLoad.Add("extensionAttribute1")

                    Using results As SearchResultCollection = searcher.FindAll()
                        For Each result As SearchResult In results
                            Using userAD As DirectoryEntry = result.GetDirectoryEntry()
                                Try
                                    If userAD.Properties.Contains("extensionAttribute1") AndAlso
                                    userAD.Properties("extensionAttribute1").Value IsNot Nothing AndAlso
                                    userAD.Properties("extensionAttribute1").Value.ToString().Trim() <> "" Then
                                        Continue For
                                    End If

                                    Dim employeeID As String = ""
                                    If userAD.Properties.Contains("employeeID") AndAlso userAD.Properties("employeeID").Value IsNot Nothing Then
                                        employeeID = userAD.Properties("employeeID").Value.ToString().Trim()
                                    End If

                                    If employeeID = "" Then Continue For

                                    Dim login As String = ""
                                    If userAD.Properties.Contains("sAMAccountName") AndAlso userAD.Properties("sAMAccountName").Value IsNot Nothing Then
                                        login = userAD.Properties("sAMAccountName").Value.ToString()
                                    End If

                                    Dim datacontract As String = Json.SendJson("", "persons/" & employeeID & "/contracts?current_contract=true", "AD", "GET", False)
                                    If String.IsNullOrWhiteSpace(datacontract) Then
                                        Commun.Journal(vbTab & "ATTENTION : " & libelleOu & " : recuperation contrats impossible, utilisateur ignore : " & login & " : employeeID : " & employeeID, True)
                                        Continue For
                                    End If

                                    Dim contracts = Json.DeserializeJson(datacontract, "contracts")
                                    Dim dateFinContratTxt As String = DateDeFinDeContract(contracts, employeeID, True)
                                    Dim dateFinContrat As Date = Now.Date
                                    Dim renseignerDateOfficielle As Boolean = True

                                    If String.IsNullOrWhiteSpace(dateFinContratTxt) Then
                                        Commun.Journal(vbTab & "ATTENTION : " & libelleOu & " : date de fin de contrat introuvable, utilisation de la date du jour pour les attributs techniques : " & login & " : employeeID : " & employeeID, False)
                                        renseignerDateOfficielle = False
                                    ElseIf Not DateTime.TryParseExact(
                                    dateFinContratTxt,
                                    "dd/MM/yyyy",
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    System.Globalization.DateTimeStyles.None,
                                    dateFinContrat
                                ) Then
                                        Commun.Journal(vbTab & "ATTENTION : " & libelleOu & " : date de fin de contrat invalide, utilisation de la date du jour pour les attributs techniques : " & login & " : employeeID : " & employeeID & " : " & dateFinContratTxt, False)
                                        dateFinContrat = Now.Date
                                        renseignerDateOfficielle = False
                                    ElseIf dateFinContrat <= #1/1/1900# Then
                                        Commun.Journal(vbTab & "ATTENTION : " & libelleOu & " : aucune date de fin de contrat exploitable, utilisation de la date du jour pour les attributs techniques : " & login & " : employeeID : " & employeeID, False)
                                        dateFinContrat = Now.Date
                                        renseignerDateOfficielle = False
                                    End If

                                    AppliquerDateFinContratRetrouvee(userAD, dateFinContratTxt, dateFinContrat, renseignerDateOfficielle)
                                    If renseignerDateOfficielle Then
                                        Commun.Journal(vbTab & libelleOu & " : date de fin de contrat completee : " & login & " : " & dateFinContratTxt, False)
                                    Else
                                        Commun.Journal(vbTab & libelleOu & " : attributs techniques completes avec la date du jour : " & login & " : employeeID : " & employeeID, False)
                                    End If
                                Catch exUser As Exception
                                    Commun.Journal(vbTab & "ERREUR : " & libelleOu & " : completion date de fin de contrat : " & userAD.Name & " : " & exUser.Message, True)
                                End Try
                            End Using
                        Next
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Commun.Journal("ERREUR : CompleterDatesContratManquantesDansOU : " & libelleOu & " : " & ex.Message, True)
        End Try
    End Sub

    Private Shared Sub AppliquerDateFinContratRetrouvee(ByVal userAD As DirectoryEntry, ByVal dateFinContratTxt As String, ByVal dateFinContrat As Date, Optional ByVal renseignerDateOfficielle As Boolean = True)
        Dim dateSuppression As Date = dateFinContrat.Date.AddMonths(3)
        Dim dateFinContratUtc As Date = dateFinContrat.Date.ToUniversalTime()
        Dim dateSuppressionUtc As Date = dateSuppression.ToUniversalTime()

        If renseignerDateOfficielle Then
            Commun.SetADLDAPProperty(userAD, "extensionAttribute1", dateFinContratTxt)
        End If
        Commun.SetADLDAPProperty(userAD, "accountDeletionDate", dateSuppression.ToString("dd/MM/yyyy"))
        userAD.Properties("accountDeactivationDT").Value = dateFinContratUtc.ToString("yyyyMMddHHmmss") & ".0Z"
        userAD.Properties("accountDeletionDT").Value = dateSuppressionUtc.ToString("yyyyMMddHHmmss") & ".0Z"
        Commun.AppliquerChangement(userAD)
    End Sub

    Shared Sub GestionAttributsDT()

        Commun.Journal("Activation/desactivation des comptes par accountDeletionDT et accountDeactivationDT", False)

        Const ADS_UF_ACCOUNT_DISABLE = 2
        Dim dateNowU As String = Now.ToUniversalTime.Date.ToString("yyyyMMddHHmmss.sZ")


        Using OuUsers As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateurs), Commun.admin, Commun.passwd, auth)

            'accountDeletionDT
            Using searcher As DirectorySearcher = New DirectorySearcher(OuUsers)
                'Exclusion de l'OU "Out" et des comptes où la date de suppression n'est pas echue
                searcher.Filter = "(&(objectClass=user)(accountDeletionDT<=" & dateNowU & ")(!(msDS-parentdistname=" & OUUtilisateursSortis & ")))"
                searcher.PageSize = 5000
                searcher.SearchScope = SearchScope.Subtree
                Dim results As SearchResultCollection = searcher.FindAll()

                For Each user As SearchResult In results

                    Using userAD As DirectoryEntry = user.GetDirectoryEntry
                        Try
                            'Cas des utilisateurs "Interne"
                            If userAD.Parent.Path = "LDAP://" & Commun.LdapPath(OUUtilisateursDesactives) Then
                                Supprime.SupprimCompte(userAD)
                                Commun.Journal("Suppression du compte interne : " & userAD.Properties("sAMAccountName").Value)

                                'Cas des utilisateurs "Externes", "Provisoires" et "Invité"
                            ElseIf userAD.Parent.Path = "LDAP://" & Commun.LdapPath(OUUtilisateursExternes) Or userAD.Parent.Path = "LDAP://" & Commun.LdapPath(ini.ReadValue("MODIFAUTO", "OUUtilisateursProvisoires")) Or userAD.Parent.Path = "LDAP://" & Commun.LdapPath(ini.ReadValue("MODIFAUTO", "OUUtilisateursInvites")) Then

                                'suppression des comptes en fonction de "accountDeletionDT"
                                If userAD.Properties.Contains("accountDeletionDT") Then
                                    Dim interval As Integer = DateDiff("d", userAD.Properties("accountDeletionDT").Value, Now.ToUniversalTime)
                                    If interval >= 1 Or userAD.Parent.Path = "LDAP://" & Commun.LdapPath(ini.ReadValue("MODIFAUTO", "OUUtilisateursInvites")) Then
                                        userAD.DeleteTree()
                                        Commun.Journal("Suppression du compte : " & userAD.Properties("sAMAccountName").Value)
                                        'GoTo fermerUsing
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            Commun.Journal("ERREUR : GestionAttributsDT : Suppression du compte : " & userAD.Properties("sAMAccountName").Value)
                        End Try
                    End Using
                Next user
            End Using



            'accountDeactivationDT
            Using searcher As DirectorySearcher = New DirectorySearcher(OuUsers)
                'Exclusion des OU "Utilisateurs", "Comptes Désactivés", "Out", des comptes désactivés et où la date de desactivation n'est pas echue
                searcher.Filter = "(&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(accountDeactivationDT<=" & dateNowU & ")(!(msDS-parentdistname=" & OUUtilisateursSortis & "))(!(msDS-parentdistname=" & OUUtilisateursDesactives & "))(!(msDS-parentdistname=" & OUUtilisateursActifs & ")))"
                searcher.PageSize = 5000
                Dim results1 As SearchResultCollection = searcher.FindAll()

                For Each user As SearchResult In results1
                    Using userAD As DirectoryEntry = user.GetDirectoryEntry
                        Try
                            'désactivation des utilisateurs en fonction de "accountDeactivationDT"
                            If Commun.AccountIsDisabled(userAD) = False Then
                                If userAD.Properties.Contains("accountDeactivationDT") Then
                                    Dim dateDesactivation As Date = userAD.Properties("accountDeactivationDT").Value
                                    Dim interval As Integer = DateDiff("d", dateDesactivation, Now.Date)
                                    If interval >= 1 Then
                                        Dim dateDesactivationtxt As String = dateDesactivation.ToString("dd/MM/yyyy") 'Date.ParseExact(dateDesactivation, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                        userAD.Properties("description").Value = "Compte éxpiré le: " & dateDesactivationtxt
                                        Commun.SetADLDAPProperty(userAD, "info", "Compte éxpiré le: " & dateDesactivationtxt, True)
                                        Commun.AppliquerChangement(userAD)
                                        Commun.ReactiveDesactiveCompte(userAD, "desactive")
                                        Commun.Journal("Désactivation du compte : " & userAD.Properties("sAMAccountName").Value)
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            Commun.Journal("ERREUR : GestionAttributsDT : Désactivation du compte : " & userAD.Properties("sAMAccountName").Value)
                        End Try
                    End Using

                Next
            End Using

            'accountActivationDT
            Using searcher As DirectorySearcher = New DirectorySearcher(OuUsers)
                'Exclusion des OU "Utilisateurs", "Comptes Désactivés", "Out", des comptes actifs et où la date de desactivation n'est pas echue et la date d'activation dépassée
                searcher.Filter = "(&(objectClass=user)((userAccountControl:1.2.840.113556.1.4.803:=2))(accountActivationDT>=" & dateNowU & ")(accountDeactivationDT<=" & dateNowU & ")(!(msDS-parentdistname=" & OUUtilisateursSortis & "))(!(msDS-parentdistname=" & OUUtilisateursDesactives & "))(!(msDS-parentdistname=" & OUUtilisateursActifs & ")))"
                searcher.PageSize = 5000
                Dim results1 As SearchResultCollection = searcher.FindAll()

                For Each user As SearchResult In results1
                    Using userAD As DirectoryEntry = user.GetDirectoryEntry
                        Try
                            'désactivation des utilisateurs en fonction de "accountDeactivationDT"
                            If Commun.AccountIsDisabled(userAD) = False Then
                                If userAD.Properties.Contains("accountActivationDT") Then
                                    Dim dateActivation As Date = userAD.Properties("accountActivationDT").Value
                                    Dim dateDesactivation As Date = userAD.Properties("accountDeactivationDT").Value
                                    Dim interval As Integer = DateDiff("d", dateDesactivation, Now.Date)
                                    If dateActivation >= dateNowU And dateDesactivation <= dateNowU Then

                                        Commun.ReactiveDesactiveCompte(userAD, "active")
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            Commun.Journal("ERREUR : GestionAttributsDT : Activation du compte : " & userAD.Properties("sAMAccountName").Value)
                        End Try
                    End Using

                Next
            End Using


            ''Reactivation des comptes qui ont été désactivés par une date "accountDeactivationDT" dépassée, mais qui a été mise a jour
            'Using searcher As DirectorySearcher = New DirectorySearcher(OuUsers)
            '    searcher.Filter = "(&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2)(accountDeactivationDT>=" & dateNowU & ")(description=Compte éxpiré le*))"
            '    searcher.PageSize = 5000
            '    Dim results1 As SearchResultCollection = searcher.FindAll()

            '    For Each user As SearchResult In results1
            '        Using userAD As DirectoryEntry = user.GetDirectoryEntry
            '            'désactivation des utilisateurs en fonction de "accountDeactivationDT"
            '            If Commun.AccountIsDisabled(userAD) = True Then
            '                If userAD.Properties.Contains("accountDeactivationDT") Then
            '                    Dim dateDesactivation As Date = userAD.Properties("accountDeactivationDT").Value
            '                    userAD.Properties("description").Clear()
            '                    Commun.SetADLDAPProperty(userAD, "info", "Compte réactivé le: " & Now.Date.ToString("dd/MM/yyyy"), True)
            '                    Commun.AppliquerChangement(userAD)
            '                    Commun.ReactiveDesactiveCompte(userAD, "active")

            '                End If
            '            End If

            '        End Using
            '    Next
            'End Using
        End Using


    End Sub
    Shared Sub GestionSuppressionProfilsItinerantsEtDossiersRedirigés()

        Commun.Journal("Debut de la gestion des profils itinerants", False)

        Dim chemin As String = "\\Clust-roamingP\Profils\BRDScientifique\"
        Dim serveur As String = Replace(chemin, "\\", "")
        serveur = Strings.Left(serveur, InStr(serveur, "\") - 1)


        Dim uVHDFiles As String() = Directory.GetFiles(chemin)
        For Each uVHD As String In uVHDFiles
            Dim fileName As String = Replace(uVHD, chemin, "")
            Dim sid As String = Replace(Replace(fileName, "UVHD-", ""), ".vhdx", "")
            Dim checkUserSID As Boolean = VerifyUserProfil(sid, "objectSid")
            If checkUserSID = False And sid <> "Sidder.exe" And sid <> "template" Then
                Try
                    File.Delete(uVHD)
                    Commun.Journal("GestionPIetDR : Suppression de l'UVHD : " & uVHD, False)
                Catch ex As Exception
                    Commun.Journal("ERREUR : GestionPIetDR : Suppression de l'UVHD : " & uVHD & " : " & ex.Message, True)
                End Try
            End If
        Next uVHD


        'For b = 0 To 1
        '    chemin = "\\Labo4\home1\"

        '    If b = 1 Then
        '        chemin = "\\Clust-roamingP\Profils\BRDAdministratif\"
        '    End If

        '    serveur = Replace(chemin, "\\", "")
        '    serveur = Strings.Left(serveur, InStr(serveur, "\") - 1)
        '    Dim subDirs As String() = Directory.GetDirectories(chemin)
        '    Array.Sort(subDirs)
        '    For Each dir As String In subDirs
        '        If Strings.Right(dir, 6) = " - old" Then Continue For
        '        Dim login As String = Replace(dir, chemin, "")
        '        Dim checkUserLogin As Boolean = VerifyUserProfil(login, "sAMAccountName")
        '        If checkUserLogin = False And login <> ".etc" And login <> "All Users" Then
        '            'MsgBox(login)
        '            If Directory.Exists(dir) Then

        '                'fermer les fichiers ouverts par l'utilisateur
        '                Dim ProcessProperties As New ProcessStartInfo
        '                ProcessProperties.FileName = "openfiles"
        '                ProcessProperties.Arguments = "/disconnect /s " & serveur & " /a " & login
        '                ProcessProperties.WindowStyle = ProcessWindowStyle.Hidden
        '                Dim myProcess As Process = Process.Start(ProcessProperties)
        '                Try
        '                    ClearAttributes(chemin & login)
        '                    Directory.Delete(dir, True)
        '                    Commun.Journal("GestionPIetDR : Suppression du profil : " & dir, False)
        '                Catch ex As Exception
        '                    Commun.Journal("ERREUR : GestionPIetDR : Suppression du profil : " & dir & " : " & ex.Message, True)
        '                End Try
        '            End If
        '        End If
        '    Next dir

        'Next b


        'Dim checkUser As Boolean = VerifyUser(login, "objectSid")


        Commun.Journal("Gestion des profils itinerants terminée", False)
    End Sub
    Shared Function VerifyUserProfil(ByVal value As String, ByVal attribut As String) As Boolean
        Dim result As Boolean = True
        Dim searcher As SearchResult = Commun.SearchFilterOne("DC=igbmc,DC=u-strasbg,DC=fr", "(&(objectClass=user)(" & attribut & "=" & value & ")(!(msDS-parentdistname=" & "OUUtilisateursSortis" & ")))", SearchScope.Subtree)
        'Dim searcher As SearchResult = Commun.SearchFilterOne("DC=igbmc,DC=u-strasbg,DC=fr", "(&(objectClass=user)(objectSid=S-1-5-21-839522115-1935655697-1343024091-10671)(!(msDS-parentdistname=" & OUUtilisateursSortis & ")))", SearchScope.Subtree)
        'Exclusion de l'OU "Out" et des comptes où la date de suppression n'est pas echue
        If searcher Is Nothing Then result = False

        Return result

    End Function
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
    Shared Sub CtrlGroupAdmins()

        Commun.Journal("Controle des groupes Admins de l'AD", False)

        Dim samaccountGroup As String = ""
        Dim listMembresAutorises As String = ""
        For compteur = 0 To 4
            Select Case compteur
                Case 0
                    samaccountGroup = "Admins du domaine"
                    listMembresAutorises = LCase(ini.ReadValue("MODIFAUTO", "Group_Admins_du_domaine"))
                    'listMembresAutorises = "Administrateur,steph,stephadm,tina,guiseithadm,SERV-TMG$,SERV-CLUSTER2$,SERV-EXCHANGE$,userprog,krbtgt"
                Case 1
                    samaccountGroup = "Administrateurs de l'entreprise"
                    listMembresAutorises = LCase(ini.ReadValue("MODIFAUTO", "Group_Administrateurs_de_l_entreprise"))
                    'listMembresAutorises = "Administrateur,steph,stephadm,guiseithadm,userprog"
                Case 2
                    samaccountGroup = "Administrateurs"
                    listMembresAutorises = LCase(ini.ReadValue("MODIFAUTO", "Group_Administrateurs"))
                    'listMembresAutorises = "Administrateur,steph,stephadm,Administrateurs,Admins du domaine,scripsteph"
                Case 3
                    samaccountGroup = "Administrateurs du schéma"
                    listMembresAutorises = LCase(ini.ReadValue("MODIFAUTO", "Group_Administrateurs_du_schema"))
                    'listMembresAutorises = "Administrateur,steph,stephadm"
                Case 4
                    samaccountGroup = "Administrateurs DHCP"
                    listMembresAutorises = LCase(ini.ReadValue("MODIFAUTO", "Group_Administrateurs_DHCP"))
                    'listMembresAutorises = "Administrateur,steph,stephadm,guiseithadm,userprog"
            End Select

            If InStr(listMembresAutorises, "administrateur,stephadm") <> 0 Then

                Dim tabMembresAutorises = Split(listMembresAutorises, ",")
                Dim monGroupe As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(Commun.TransformeSAMACCOUNTenCN(samaccountGroup)), Commun.admin, Commun.passwd, auth)
                Try
                    ' Groupe dont les membres sont à lister
                    Dim tabMembre As String() = Commun.MembresDuGroupe(samaccountGroup, True)
                    'tabMembre = Split(LCase(Join(tabMembre, ";")), ";")
                    For Each unMembre In tabMembre

                        Dim userLogin As String = Commun.TransformeSAMACCOUNTenCN(unMembre.ToString) 'user.Properties("sAMAccountName").Value
                        userLogin = LCase(userLogin)

                        If Array.IndexOf(tabMembresAutorises, userLogin) = -1 Then
                            monGroupe.Invoke("Remove", New Object() {"LDAP://" & Commun.LdapPath(unMembre.ToString)})
                        End If
                    Next unMembre

                    For i = 0 To UBound(tabMembresAutorises)
                        Dim DNUser As String = Commun.TransformeSAMACCOUNTenCN(tabMembresAutorises(i))
                        'DNUser = LCase(DNUser)
                        Dim found As Integer = Array.IndexOf(tabMembre, DNUser)
                        If found = -1 Then
                            monGroupe.Invoke("Add", New Object() {"LDAP://" & Commun.LdapPath(DNUser)})
                        End If
                    Next i

                Catch ex As Exception
                    Commun.Journal("ERREUR : Controle des groupes Admins : case " & compteur & " : " & ex.Message, True)
                Finally
                    monGroupe.Close()
                    monGroupe.Dispose()
                    monGroupe = Nothing
                End Try
            Else
                Commun.Journal("ERREUR : Controle des groupes Admins : Liste Vide : case " & compteur, True)
            End If
        Next
        Commun.Journal("Fin du Controle des groupes Admins", False)
    End Sub

    Shared Sub UpdateComptesProvisoires()

        Commun.Journal("Mise a jour des utilisateurs provisoires", False)

        'Mise a jour des utilisateurs provisoires
        Using OUProvisoire As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(ini.ReadValue("MODIFAUTO", "OUUtilisateursProvisoires")), Commun.admin, Commun.passwd, auth)
            Dim searcher As DirectorySearcher = New DirectorySearcher(OUProvisoire)
            searcher.Filter = "(&(objectClass=user))"

            Dim result As SearchResultCollection = searcher.FindAll()
            'Dim dateExpTexteDescription As String

            For Each utilisateurProv As SearchResult In result
                Using userUpdate As DirectoryEntry = utilisateurProv.GetDirectoryEntry

                    Dim dateExpDate As Date = userUpdate.Properties("accountDeactivationDT").Value
                    Dim dateExpDateTxt As String = dateExpDate.ToString("dd/MM/yyyy")
                    Try

                        Try
                            'Nettoyage des destinations sur les utilisateurs provisoires
                            For i As Integer = 0 To userUpdate.Properties("MemberOf").Count - 1
                                If InStr(userUpdate.Properties("MemberOf")(i), " grp,") > 0 Then
                                    Commun.AddRemoveADGroup(userUpdate.Properties("sAMAccountName").Value, userUpdate.Properties("MemberOf")(i), "Remove")
                                End If
                            Next
                        Catch ex As Exception
                            Commun.Journal("ERREUR : UpdateComptesProvisoires : Nettoyage destination : " & userUpdate.Properties("sAMAccountName").Value & " : " & ex.Message, True)
                        End Try

                        Try
                            If dateExpDateTxt <> Strings.Right(userUpdate.Properties("Description").Value, 10) Then
                                userUpdate.Properties("description").Value = "COMPTE PROVISOIRE expire le: " & dateExpDateTxt
                                Commun.AppliquerChangement(userUpdate)
                            End If
                        Catch ex As Exception
                            Commun.Journal("ERREUR : UpdateComptesProvisoires : Mise a jour Description :" & userUpdate.Properties("sAMAccountName").Value & " : " & ex.Message, True)
                        End Try

                    Catch ex As Exception
                        Commun.Journal("ERREUR : UpdateComptesProvisoires : " & userUpdate.Properties("sAMAccountName").Value & " : " & ex.Message, True)
                    End Try
                End Using
            Next

        End Using
        Commun.Journal("Fin de la mise a jour des utilisateurs provisoires", False)
    End Sub





    Shared Sub ControleOUUtilisateurs()

        Commun.Journal("Controle de l'OU Utilisateurs", False)

        'Desactivation des comptes dans l'OU Exception qui n'ont rien a y faire
        Dim OUException As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateursExceptions), Commun.admin, Commun.passwd, auth)
        Dim searchOUException As DirectorySearcher = New DirectorySearcher(OUException)
        searchOUException.Filter = "(&(objectClass=user))"
        searchOUException.PropertiesToLoad.Add("samAccountName")
        Dim searchOUExceptionResult As SearchResultCollection = searchOUException.FindAll()
        For Each userOUException As SearchResult In searchOUExceptionResult
            If Commun.IsException(userOUException.Properties("samAccountName")(0), "") = False Then
                Dim userToDisable As DirectoryEntry = userOUException.GetDirectoryEntry
                userToDisable.NativeObject.accountdisabled = True
                Commun.AppliquerChangement(userToDisable)
                userToDisable.Close()
                userToDisable.Dispose()
                userToDisable = Nothing
                'userOUException.GetDirectoryEntry.DeleteTree()
            End If
        Next
        searchOUExceptionResult.Dispose()
        searchOUExceptionResult = Nothing
        searchOUException.Dispose()
        searchOUException = Nothing
        OUException.Close()
        OUException.Dispose()
        OUException = Nothing

        'Ajout au groupe G_Domain_DisableOpenSession des comptes de l'OU "Comptes Désactivés"
        Dim DNDisableOpenSession As String = Commun.TransformeSAMACCOUNTenCN("G_Domain_DisableOpenSession")
        Dim IDDisableOpenSession As Integer = Commun.PrimaryGroupId(Commun.FindAttribut("G_Domain_DisableOpenSession", "objectSid"))
        Using OUComptesDesactives As New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateursDesactives), Commun.admin, Commun.passwd, auth)
            Using searchDesactiv As DirectorySearcher = New DirectorySearcher(OUComptesDesactives)
                'primaryGroupID=513 correspond au groupe Utilisa. du domaine
                searchDesactiv.PageSize = 5000
                searchDesactiv.Filter = "(&(objectCategory=person)(objectClass=user)" &
                                            "(|(primaryGroupID=513)))"
                Using searchDesactivResult As SearchResultCollection = searchDesactiv.FindAll()
                    For Each objUserDesactiv As SearchResult In searchDesactivResult

                        Using DirEntry As DirectoryEntry = objUserDesactiv.GetDirectoryEntry

                            Try
                                GestionGroupeUserDesactive(DirEntry)
                            Catch
                            End Try
                        End Using
                    Next
                End Using
            End Using
        End Using

        Commun.Journal("Fin du controle de l'OU Utilisateurs", False)
    End Sub


    Shared Sub GestionDestinationsDepartements()
        Commun.Journal("Gestion des destinations et departements")

        If DicoDestinationsRH Is Nothing OrElse DicoDestinationsRH.Count = 0 Then Throw New Exception("DicoDestinationsRH est vide")



        'MODIFICATION ET AJOUT DES NOUVELLES DESTINATIONS
        Using OUDomDest As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("OU=Equipes,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
            Using dirDestSearcher As DirectorySearcher = New DirectorySearcher(OUDomDest)
                Dim dicoDestAd As New Dictionary(Of String, SearchResult)(StringComparer.OrdinalIgnoreCase)
                dirDestSearcher.Filter = "(&(objectClass=group)(languageCode=*))"
                dirDestSearcher.SearchScope = SearchScope.OneLevel
                dirDestSearcher.PropertiesToLoad.Clear()
                dirDestSearcher.PropertiesToLoad.Add("languageCode")

                Dim resultDestExistants As SearchResultCollection = dirDestSearcher.FindAll()
                For Each sr As SearchResult In resultDestExistants
                    If sr.Properties.Contains("languageCode") Then
                        Dim idDestExistant As String = sr.Properties("languageCode")(0).ToString().Trim()
                        If idDestExistant <> "" AndAlso Not dicoDestAd.ContainsKey(idDestExistant) Then
                            dicoDestAd.Add(idDestExistant, sr)
                        End If
                    End If
                Next

                Dim maxTadDest As Integer = DicoDestinationsRH.Count - 1
                Dim l As Integer = -1
                For Each dest As DestinationInfo In DicoDestinationsRH.Values
                    l += 1
                    Commun.AfficherBarre("gestion des destinations", l, maxTadDest, False)

                    Dim dept As DepartementInfo = Nothing
                    If Not String.IsNullOrWhiteSpace(dest.id_dept) AndAlso DicoDepartementsRH.ContainsKey(dest.id_dept) Then
                        dept = DicoDepartementsRH(dest.id_dept)
                    End If

                    Dim SAMAccountGroup As String = dest.nom_court_dest & " grp"
                    Dim description As String = dest.nom_long_dest
                    Dim IdDest As String = dest.id_dest
                    Dim samManager As String = dest.login_responsable_dest
                    Dim department_nom As String = If(dept IsNot Nothing, dept.nom_long_dept, "")
                    Dim department_ID As String = If(dept IsNot Nothing, dept.id_dept, "")


                    Try
                        'L'attribut "languageCode" contient le groupe ID de GDPI
                        'L'attribut "languageCode" contient le groupe ID de GDPI
                        Dim result As SearchResult = Nothing
                        If dicoDestAd.ContainsKey(IdDest) Then
                            result = dicoDestAd(IdDest)
                        End If
                        Dim dirDest As New DirectoryEntry
                        If result Is Nothing Then
                            'Creation d'une nouvelle destination
                            Commun.NouveauGroupe(OUDomDest, SAMAccountGroup, description)
                            dirDest = New DirectoryEntry("LDAP://" & Commun.LdapPath(Commun.TransformeSAMACCOUNTenCN(SAMAccountGroup)), Commun.admin, Commun.passwd, auth)
                            'L'attribut "languageCode" contient le groupe ID de GDPI
                            dirDest.Properties("languageCode").Value = IdDest
                            Commun.AppliquerChangement(dirDest)
                            Commun.Journal(vbTab & "Creation d'une nouvelle destination: " & SAMAccountGroup)
                            dirDestSearcher.Filter = "(&(objectClass=group)(languageCode=" & IdDest & "))"
                            dirDestSearcher.SearchScope = SearchScope.OneLevel
                            Dim createdResult As SearchResult = dirDestSearcher.FindOne()
                            If createdResult IsNot Nothing Then
                                dicoDestAd(IdDest) = createdResult
                            End If
                        Else
                            dirDest = result.GetDirectoryEntry
                        End If



                        'Modification des destinations 


                        If dirDest.Properties("SAMAccountName").Value <> SAMAccountGroup Then
                            dirDest.Properties("SAMAccountName").Value = SAMAccountGroup
                            Commun.AppliquerChangement(dirDest)
                            dirDest.Rename("CN=" & SAMAccountGroup)
                            Commun.AppliquerChangement(dirDest)
                        End If

                        If samManager <> "?" Then
                            Dim CNchefEquipe As String = Commun.TransformeSAMACCOUNTenCN(samManager)
                            If dirDest.Properties("managedBy").Value <> CNchefEquipe Then
                                dirDest.Properties("managedBy").Value = CNchefEquipe
                                Commun.AppliquerChangement(dirDest)
                            End If
                        End If

                        If dirDest.Properties("description").Value <> description Then
                            dirDest.Properties("description").Value = description
                            Commun.AppliquerChangement(dirDest)
                        End If
                        If dirDest.Properties("department").Value <> department_nom Then
                            dirDest.Properties("department").Value = department_nom
                            Commun.AppliquerChangement(dirDest)
                        End If



                        dirDest.Close()
                        dirDest.Dispose()
                        dirDest = Nothing
                        result = Nothing
                    Catch ex As Exception
                        Commun.Journal(vbTab & "ERREUR : Gestion des Destinations : Ajout/Modification " & SAMAccountGroup & " : " & ex.Message, True)
                    End Try

                Next


                'gestion des departements
                Using OUDomDepart As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
                    Using dirDepartSearcher As DirectorySearcher = New DirectorySearcher(OUDomDepart)
                        Dim dicoDepartAd As New Dictionary(Of String, SearchResult)(StringComparer.OrdinalIgnoreCase)
                        dirDepartSearcher.Filter = "(&(objectClass=group)(languageCode=*))"
                        dirDepartSearcher.SearchScope = SearchScope.OneLevel
                        dirDepartSearcher.PropertiesToLoad.Clear()
                        dirDepartSearcher.PropertiesToLoad.Add("languageCode")

                        Dim resultDepartExistants As SearchResultCollection = dirDepartSearcher.FindAll()
                        For Each sr As SearchResult In resultDepartExistants
                            If sr.Properties.Contains("languageCode") Then
                                Dim idDepartExistant As String = sr.Properties("languageCode")(0).ToString().Trim()
                                If idDepartExistant <> "" AndAlso Not dicoDepartAd.ContainsKey(idDepartExistant) Then
                                    dicoDepartAd.Add(idDepartExistant, sr)
                                End If
                            End If
                        Next

                        Dim maxTabDept As Integer = DicoDepartementsRH.Count - 1
                        Dim d As Integer = -1
                        For Each dept As DepartementInfo In DicoDepartementsRH.Values
                            d += 1
                            Commun.AfficherBarre("gestion des departements", d, maxTabDept, False)

                            Dim SAMAccountDepart As String = "Dpt_" & dept.nom_court_dept
                            Dim department_nom As String = dept.nom_long_dept
                            Dim IdDepart As String = dept.id_dept
                            Try
                                'L'attribut "languageCode" contient le departement ID de GDPI
                                'L'attribut "languageCode" contient le departement ID de GDPI
                                Dim result As SearchResult = Nothing
                                If dicoDepartAd.ContainsKey(IdDepart) Then
                                    result = dicoDepartAd(IdDepart)
                                End If
                                Dim dirDepart As New DirectoryEntry
                                If result Is Nothing Then
                                    'Creation d'un nouveau departement
                                    Try
                                        Commun.NouveauGroupe(OUDomDepart, SAMAccountDepart, "Departement " & department_nom)
                                        Commun.Journal(vbTab & "Creation d'un nouveau departement: " & SAMAccountDepart)
                                        dirDepartSearcher.Filter = "(&(objectClass=group)(languageCode=" & IdDepart & "))"
                                        dirDepartSearcher.SearchScope = SearchScope.OneLevel
                                        Dim createdResult As SearchResult = dirDepartSearcher.FindOne()
                                        If createdResult IsNot Nothing Then
                                            dicoDepartAd(IdDepart) = createdResult
                                        End If
                                    Catch

                                    End Try
                                    dirDepart = New DirectoryEntry("LDAP://" & Commun.LdapPath(Commun.TransformeSAMACCOUNTenCN(SAMAccountDepart)), Commun.admin, Commun.passwd, auth)
                                    'L'attribut "languageCode" contient le groupe ID de GDPI
                                    dirDepart.Properties("languageCode").Value = IdDepart
                                    Commun.AppliquerChangement(dirDepart)
                                Else
                                    dirDepart = result.GetDirectoryEntry
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

                                'If dirDepart.Properties("displayName").Value <> "Departement " & department_nom Then
                                dirDepart.Properties("displayName").Value = dirDepart.Properties("description").Value
                                Commun.AppliquerChangement(dirDepart)
                                'End If

                                dirDepart.Close()
                                dirDepart.Dispose()
                                dirDepart = Nothing
                                result = Nothing
                            Catch ex As Exception
                                Commun.Journal(vbTab & "ERREUR : Gestion des Departements: Ajout/Modification " & SAMAccountDepart & " : " & ex.Message, True)
                            End Try


                        Next
                    End Using

                    'gestion des PI departements
                    PIDepartementGroup()
                    '----------------------------------------------------------------------------------------------------------------------------------
                    'CONTROLE DES DESTINATIONS
                    dirDestSearcher.Filter = "(&(objectClass=group))"
                    dirDestSearcher.SearchScope = SearchScope.OneLevel
                    dirDestSearcher.PropertiesToLoad.Add("SAMAccountName")
                    dirDestSearcher.PropertiesToLoad.Add("description")
                    Dim result1 As SearchResultCollection = dirDestSearcher.FindAll

                    Dim dicoSamDestRh As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    For Each dest As DestinationInfo In DicoDestinationsRH.Values
                        dicoSamDestRh.Add(dest.nom_court_dest & " grp")
                    Next
                    dicoSamDestRh.Add("EXTERNE grp")

                    For Each grpResult As SearchResult In result1
                        Dim SAMAccountGroup As String = grpResult.Properties("SAMAccountName")(0)
                        Try
                            Dim descriptionGroup As String = ""
                            If grpResult.Properties.Contains("description") Then
                                descriptionGroup = grpResult.Properties("description")(0)
                            End If
                            'Controle si l'equipe Existe dans la BDP
                            Dim ctrlGrpExistBDP As Boolean = dicoSamDestRh.Contains(SAMAccountGroup)

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
                                        Commun.Journal(vbTab & "Equipe supprimée : " & descriptionGroup & " (" & SAMAccountGroup & ")")
                                        'Pour ne pas passer dans la fonction de controle d'association destination<->equipeinfo
                                        Continue For

                                    Else
                                        'controle si l'utilisateur est en exception valide avant de dire que le repertoire n'est pas vide
                                        Dim dernierMembre As String = Commun.TransformeSAMACCOUNTenCN(grpResult.GetDirectoryEntry.Properties("Member")(0))
                                        If exceptionUser(dernierMembre) = "False" Then
                                            Throw New Exception("Nettoyage : Le groupe n'est pas vide")
                                        End If
                                    End If
                                End If
                            End If

                            'Controle si la destination est associé a une Equipeinfo 
                            If grpResult.GetDirectoryEntry.Properties("Member").Count <> 0 Then
                                If Commun.RecupEquipeinfo(Replace(SAMAccountGroup, " grp", "")) = "" Then
                                    If Hour(Now) = 5 Or Hour(Now) = 6 Then
                                        Commun.SendEmail(vbTab & "administrateur@igbmc.fr", ini.ReadValue("GLOBAL", "MailEquipeReso"), "Destination non associée a une Equipe Info", "La destination " & descriptionGroup & " (" & SAMAccountGroup & "), dans l'active Directory, n'est actuellement pas associé à une Equipe info." & vbCrLf & "Cette destination contient un ou des utilisateurs." & vbCrLf & "Tant que cette association ne sera pas faite, les utilisateurs n'auront pas d'acces a leur Zone Labo")
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            Commun.Journal(vbTab & "ERREUR : Gestion des Destinations : Controle : " & SAMAccountGroup & " : " & ex.Message, True)
                        End Try
                    Next grpResult
                    result1.Dispose()
                    result1 = Nothing
                End Using
            End Using
        End Using
        Commun.Journal("Gestion des destinations et départements réussie", False)

    End Sub




    Shared Sub PIDepartementGroup()
        Using dptAD As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath("OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
            Dim resultsDpt As SearchResultCollection = Commun.SearchFilterAll("OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr", "objectClass=group", SearchScope.OneLevel, "description,cn")

            Dim maxTabPI As Integer = resultsDpt.Count
            Dim p As Integer = 0
            For Each resultDpt As SearchResult In resultsDpt
                p += 1
                Commun.AfficherBarre("gestion des PI", p, maxTabPI, False)
                Dim description As String = Replace(resultDpt.Properties("description")(0), "Departement", "")
                Dim nom As String = resultDpt.Properties("cn")(0)
                Dim resultsDest As SearchResultCollection = Commun.SearchFilterAll("OU=Equipes,DC=igbmc,DC=u-strasbg,DC=fr", "(&(objectClass=group)(department=" & description & "))", SearchScope.OneLevel, "department")
                Dim nomGrpPI As String = Replace(nom, "Dpt_", "Dpt_PI-")
                For Each resultDest As SearchResult In resultsDest
                    Using destAD As DirectoryEntry = New DirectoryEntry(resultDest.Path)
                        'MsgBox(destAD.Properties("sAMAccountName").Value)
                        Dim chefCN As String = destAD.Properties("managedBy").Value
                        If Not DirectoryEntry.Exists("LDAP://" & Commun.LdapPath("CN=" & nomGrpPI & ",OU=Departements PI,OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr")) Then
                            Commun.NouveauGroupe(New DirectoryEntry("LDAP://" & Commun.LdapPath("OU=Departements PI,OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth), nomGrpPI, "")
                        End If
                        Commun.AddRemoveADGroup(chefCN, nomGrpPI, "Add")
                    End Using
                Next
            Next

        End Using
    End Sub


    Shared Sub GestionReactiveDesactiveComptesInterne(adUsersByEmployeeId As Dictionary(Of String, UtilisateurADIndex))
        Dim i As Integer = -1
        Commun.Journal("Gestion des activation/desactivation des comptes interne")

        For Each kvp As KeyValuePair(Of String, UtilisateurADIndex) In adUsersByEmployeeId
            i += 1
            Commun.AfficherBarre("Gestion de l'activation/desactivation des comptes", i, adUsersByEmployeeId.Count, False)

            Dim userAD As UtilisateurADIndex = kvp.Value
            If userAD Is Nothing Then Continue For
            If String.IsNullOrWhiteSpace(userAD.samAccountName) Then Continue For

            Dim login As String = userAD.samAccountName
            Dim employeeID As String = kvp.Key
            If String.IsNullOrWhiteSpace(employeeID) Then employeeID = userAD.employeeID

            Dim dateTxtFinException As String = exceptionUser(login)
            Dim estEnException As Boolean = LCase(Trim(dateTxtFinException)) <> "false"
            Dim estPresentRH As Boolean = PresenceBDP(employeeID)
            Dim dateFinException As Date = Date.MinValue
            Dim dateExceptionValide As Boolean = False

            If estEnException Then
                dateExceptionValide = TryLireDateException(dateTxtFinException, dateFinException)
            End If

            'L'utilisateur est present RH : il doit etre dans l'etat actif.
            'S'il n'est pas deja dans l'OU actifs, on devra le reactiver et le deplacer vers OUUtilisateursActifs.
            Dim doitPasserActif As Boolean =
            estPresentRH AndAlso Not userAD.isInOUActifs

            Dim descriptionExceptionAttendue As String = ""
            If estEnException Then
                descriptionExceptionAttendue = "EXCEPTION jusqu'au: " & dateTxtFinException
            End If

            'L'utilisateur n'est plus present RH, a une exception valide,
            'et n'est pas deja dans l'OU exceptions : il devra etre deplace vers OUUtilisateursExceptions.
            Dim doitPasserException As Boolean =
            Not estPresentRH AndAlso
            estEnException AndAlso
            Not userAD.isInOUExceptions

            'L'utilisateur n'est plus present RH, a une exception valide,
            'et sa description AD ne correspond pas a la date d'exception actuelle.
            Dim doitMettreAJourException As Boolean =
            Not estPresentRH AndAlso
            estEnException AndAlso
            userAD.description <> descriptionExceptionAttendue

            'L'utilisateur n'est plus present RH, a une exception valide,
            'et ses dates techniques doivent correspondre a la fin de l'exception.
            Dim doitMettreAJourDatesException As Boolean =
            Not estPresentRH AndAlso
            estEnException AndAlso
            dateExceptionValide AndAlso
            DatesTechniquesExceptionDifferentes(userAD, dateFinException)

            'L'utilisateur n'est plus present RH, n'a pas d'exception valide,
            'et n'est ni deja desactive, ni deja sorti : il devra passer dans OUUtilisateursDesactives.
            Dim doitPasserDesactive As Boolean =
            Not estPresentRH AndAlso
            Not estEnException AndAlso
            Not userAD.isInOUDesactives AndAlso
            Not userAD.isInOUSortis

            'Aucune ouverture LDAP n'est faite si le compte est deja dans le bon etat.
            Dim actionNecessaire As Boolean =
            doitPasserActif OrElse
            doitPasserException OrElse
            doitMettreAJourException OrElse
            doitMettreAJourDatesException OrElse
            doitPasserDesactive

            If Not actionNecessaire Then
                Continue For
            End If

            Using dirEntry As New DirectoryEntry("LDAP://" & Commun.LdapPath(userAD.distinguishedName), Commun.admin, Commun.passwd, auth)

                'Si l'utilisateur est present RH mais n'est pas dans l'OU actifs,
                'on nettoie les attributs de desactivation, on reactive les comptes,
                'puis on le deplace vers OUUtilisateursActifs.
                If doitPasserActif Then
                    Try
                        GestionGroupeUserActive(dirEntry)

                        Commun.SetADLDAPProperty(dirEntry, "description", "")
                        Commun.SetADLDAPProperty(dirEntry, "Comment", "Réactivé le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf, True)
                        ClearAttributDeactivationDeletionDate(dirEntry, userAD)
                        Commun.AppliquerChangement(dirEntry)

                        userAD.description = ""
                        userAD.comment = "Réactivé le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf

                        Commun.ReactiveDesactiveCompte(login, "Active")
                    Catch ex As Exception
                        Commun.Journal(vbTab & "ERREUR : réactivation du compte(propriétés) : " & dirEntry.Name & " : " & ex.Message, True)
                    End Try

                    'Active le compte loginadm s'il existe
                    Commun.ReactiveDesactiveCompte(login & "adm", "Active")

                    Try
                        CompararaisonAjoutRetraitDestinationsDepartement(dirEntry)
                        MoveToOu(dirEntry, userAD, OUUtilisateursActifs)
                    Catch ex As Exception
                        Commun.Journal(vbTab & "ERREUR : réactivation du compte (deplacement Vers OU Utilisateurs) : " & dirEntry.Name, True)
                    End Try
                End If

                'Si l'utilisateur n'est plus present RH mais a une exception valide,
                'on corrige uniquement ce qui est necessaire : description d'exception,
                'dates de fin d'exception, ou deplacement vers OUUtilisateursExceptions.
                If doitPasserException OrElse doitMettreAJourException OrElse doitMettreAJourDatesException Then
                    Try
                        Dim commentaireException As String = "Exception le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf

                        If doitMettreAJourException Then
                            Commun.SetADLDAPProperty(dirEntry, "description", descriptionExceptionAttendue)
                            Commun.SetADLDAPProperty(dirEntry, "Comment", commentaireException, True)
                            userAD.description = descriptionExceptionAttendue
                            userAD.comment = commentaireException
                        End If

                        If doitMettreAJourDatesException Then
                            AppliquerDatesException(dirEntry, userAD, dateFinException)
                        End If

                        If doitMettreAJourException OrElse doitMettreAJourDatesException Then
                            Commun.AppliquerChangement(dirEntry)
                        End If
                    Catch ex As Exception
                        Commun.Journal(vbTab & "ERREUR : mise à jour exception du compte : " & dirEntry.Name, True)
                    End Try

                    'Si l'utilisateur est en exception valide mais n'est pas deja dans l'OU exceptions,
                    'on reactive le compte adm si besoin, on remet les groupes actifs,
                    'puis on le deplace vers OUUtilisateursExceptions.
                    If doitPasserException Then
                        Commun.ReactiveDesactiveCompte(login & "adm", "Active")

                        Try
                            GestionGroupeUserActive(dirEntry)
                            MoveToOu(dirEntry, userAD, OUUtilisateursExceptions)
                        Catch ex As Exception
                            Commun.Journal(vbTab & "ERREUR : réactivation du compte (deplacement vers OU Exceptions) : " & dirEntry.Name, True)
                        End Try
                    End If
                End If

                'Si l'utilisateur n'est plus present RH, n'a pas d'exception valide,
                'et n'est pas deja dans l'etat desactive ou sorti,
                'on le fait passer dans l'etat desactive.
                'Le compte login reste dans la periode desactive, avec une date de sortie prevue a +3 mois.
                'Le compte loginadm est desactive.
                'La sortie definitive du compte login est geree plus tard dans OUUtilisateursSortis.
                If doitPasserDesactive Then
                    Dim dateDefinDeContrat As String = ""

                    Try
                        Dim commentaireDesactivation As String = ""
                        Dim descriptionDesactivation As String = ""

                        Dim id As String = userAD.employeeID

                        'Si l'employeeID est disponible, on recupere la date de fin de contrat reelle.
                        'Cette date sert de reference pour accountDeactivationDT et pour calculer la date de sortie a +3 mois.
                        If id <> "" Then
                            Dim datacontract As String = Json.SendJson("", "persons/" & id & "/contracts?current_contract=true", "AD", "GET")
                            Dim contracts = Json.DeserializeJson(datacontract, "contracts")
                            dateDefinDeContrat = DateDeFinDeContract(contracts, id)

                        End If

                        If String.IsNullOrWhiteSpace(dateDefinDeContrat) AndAlso Not String.IsNullOrWhiteSpace(userAD.extensionAttribute1) Then
                            dateDefinDeContrat = userAD.extensionAttribute1.Trim()
                        End If

                        Dim dateReferenceDesactivation As Date = Now.Date
                        Dim dateFinContratValide As Boolean = False

                        If dateDefinDeContrat <> "" Then
                            Dim dateFinContrat As Date
                            If DateTime.TryParseExact(
                                 dateDefinDeContrat,
                                 "dd/MM/yyyy",
                                 System.Globalization.CultureInfo.InvariantCulture,
                                 System.Globalization.DateTimeStyles.None,
                                 dateFinContrat
                             ) Then
                                dateReferenceDesactivation = dateFinContrat.Date
                                dateFinContratValide = True
                            Else
                                Commun.Journal(vbTab & "ERREUR : GestionReactiveDesactiveCompte : conversion de la date de fin de contrat : " & dirEntry.Name & " : " & dateDefinDeContrat & " : utilisation de la date du jour pour les attributs techniques", True)
                                dateDefinDeContrat = ""
                            End If
                        Else
                            Commun.Journal(vbTab & "ATTENTION : GestionReactiveDesactiveCompte : date de fin de contrat absente, utilisation de la date du jour pour les attributs techniques : " & dirEntry.Name & " : employeeID : " & id, True)
                        End If

                        Dim dateSortiePrevue As Date = dateReferenceDesactivation.AddMonths(3)
                        Dim dateReferenceDesactivationUtc As Date = dateReferenceDesactivation.ToUniversalTime()
                        Dim dateSortiePrevueUtc As Date = dateSortiePrevue.ToUniversalTime()

                        'Si le compte vient de l'etat exception, on garde une trace explicite dans la description.
                        If userAD.isInOUExceptions Then
                            descriptionDesactivation = "EXCEPTION Désactivé le: " & Strings.Left(CStr(Now), 10)
                            commentaireDesactivation = "EXCEPTION Désactivé le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf
                        Else
                            descriptionDesactivation = "Désactivé le: " & Strings.Left(CStr(Now), 10)
                            commentaireDesactivation = "Désactivé le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf
                        End If

                        Commun.SetADLDAPProperty(dirEntry, "description", descriptionDesactivation)
                        Commun.SetADLDAPProperty(dirEntry, "Comment", commentaireDesactivation, True)
                        Commun.SetADLDAPProperty(dirEntry, "accountDeletionDate", dateSortiePrevue.ToString("dd/MM/yyyy"))
                        dirEntry.Properties("accountDeactivationDT").Value = dateReferenceDesactivationUtc.ToString("yyyyMMddHHmmss") & ".0Z"
                        dirEntry.Properties("accountDeletionDT").Value = dateSortiePrevueUtc.ToString("yyyyMMddHHmmss") & ".0Z"

                        userAD.description = descriptionDesactivation
                        userAD.comment = commentaireDesactivation
                        userAD.accountDeactivationDT = dateReferenceDesactivationUtc
                        userAD.accountDeletionDT = dateSortiePrevueUtc
                        userAD.accountDeletionDate = dateSortiePrevue.ToString("dd/MM/yyyy")

                        If dateFinContratValide Then
                            Commun.SetADLDAPProperty(dirEntry, "extensionAttribute1", dateDefinDeContrat)
                            userAD.extensionAttribute1 = dateDefinDeContrat
                        End If

                        Commun.SetADLDAPProperty(dirEntry, "serialNumber", Nothing)
                        Commun.SetADLDAPProperty(dirEntry, "employeeNumber", Nothing)
                        Commun.SetADLDAPProperty(dirEntry, "title", "")
                        Commun.AppliquerChangement(dirEntry)

                        userAD.title = ""
                        userAD.serialNumber = Nothing
                        userAD.employeeNumber = Nothing
                    Catch ex As Exception
                        Commun.Journal(vbTab & "ERREUR : désactivation du compte(propriétés) : " & dirEntry.Name, True)
                    End Try

                    'Lorsque le compte principal passe dans l'etat desactive, le compte admin associe est desactive.
                    Commun.ReactiveDesactiveCompte(login & "adm", "Desactive")

                    'Avant le deplacement vers OUUtilisateursDesactives,
                    'on applique les groupes/restrictions correspondant a l'etat desactive.
                    GestionGroupeUserDesactive(dirEntry)
                    CompararaisonAjoutRetraitDestinationsDepartement(dirEntry)

                    'Preparation du mail d'information envoye lors du passage en etat desactive.
                    Dim adresseMail As String = userAD.mail
                    Dim prenom As String = userAD.givenName

                    Dim dateDeSuppressionPrevue As String = userAD.accountDeletionDate

                    Try
                        Dim mail As String = MailCloture(prenom, dateDeSuppressionPrevue, dateDefinDeContrat)
                        'Commun.SendEmail("noreply@igbmc.fr", adresseMail & ";Cc:serviceinfo@igbmc.fr", "ARRET DU COMPTE", mail)
                        Commun.Journal(vbTab & "Mail de fermeture de compte envoyé : " & adresseMail)
                    Catch ex As Exception
                        Commun.Journal(vbTab & "ERREUR : Mail de fermeture de compte envoyé : " & adresseMail & " : " & ex.Message, True)
                    End Try

                    'Deplacement final du compte vers l'OU des comptes desactives.
                    Try
                        MoveToOu(dirEntry, userAD, OUUtilisateursDesactives)
                    Catch ex As Exception
                        Commun.Journal(vbTab & "ERREUR : deplacement du compte vers l'OU Utilisateurs Desactives : " & dirEntry.Name & " : " & ex.Message, True)
                    End Try
                End If
            End Using
        Next

        Commun.Journal("Gestion de l'activation/desactivation des comptes réussie", False)
    End Sub
    Private Shared Function TryLireDateException(ByVal dateTxtFinException As String, ByRef dateFinException As Date) As Boolean
        If String.IsNullOrWhiteSpace(dateTxtFinException) Then Return False

        Return Date.TryParseExact(
        dateTxtFinException.Trim(),
        "dd/MM/yyyy",
        System.Globalization.CultureInfo.InvariantCulture,
        System.Globalization.DateTimeStyles.None,
        dateFinException
    )
    End Function
    Private Shared Function DatesTechniquesExceptionDifferentes(ByVal userAD As UtilisateurADIndex, ByVal dateFinException As Date) As Boolean
        If userAD Is Nothing Then Return True

        Dim dateSuppressionException As Date = dateFinException.Date.AddMonths(3)
        Dim dateFinExceptionUtc As Date = dateFinException.Date.ToUniversalTime()
        Dim dateSuppressionExceptionUtc As Date = dateSuppressionException.ToUniversalTime()
        Dim accountDeletionDateAttendue As String = dateSuppressionException.ToString("dd/MM/yyyy")

        Return Not userAD.accountDeactivationDT.HasValue OrElse
        userAD.accountDeactivationDT.Value.Date <> dateFinExceptionUtc.Date OrElse
        Not userAD.accountDeletionDT.HasValue OrElse
        userAD.accountDeletionDT.Value.Date <> dateSuppressionExceptionUtc.Date OrElse
        userAD.accountDeletionDate <> accountDeletionDateAttendue
    End Function
    Private Shared Sub AppliquerDatesException(ByVal dirEntry As DirectoryEntry, ByVal userAD As UtilisateurADIndex, ByVal dateFinException As Date)
        Dim dateSuppressionException As Date = dateFinException.Date.AddMonths(3)
        Dim dateFinExceptionUtc As Date = dateFinException.Date.ToUniversalTime()
        Dim dateSuppressionExceptionUtc As Date = dateSuppressionException.ToUniversalTime()
        Dim accountDeletionDate As String = dateSuppressionException.ToString("dd/MM/yyyy")

        Commun.SetADLDAPProperty(dirEntry, "accountDeletionDate", accountDeletionDate)
        dirEntry.Properties("accountDeactivationDT").Value = dateFinExceptionUtc.ToString("yyyyMMddHHmmss") & ".0Z"
        dirEntry.Properties("accountDeletionDT").Value = dateSuppressionExceptionUtc.ToString("yyyyMMddHHmmss") & ".0Z"

        If userAD IsNot Nothing Then
            userAD.accountDeactivationDT = dateFinExceptionUtc
            userAD.accountDeletionDT = dateSuppressionExceptionUtc
            userAD.accountDeletionDate = accountDeletionDate
        End If
    End Sub
    Shared Sub ClearAttributDeactivationDeletionDate(ByVal dirEntry As DirectoryEntry, ByVal userAD As UtilisateurADIndex)
        Try
            Dim extensionAttribute1 As String = ""

            If dirEntry.Properties.Contains("extensionAttribute1") AndAlso
            dirEntry.Properties("extensionAttribute1").Value IsNot Nothing Then
                extensionAttribute1 = dirEntry.Properties("extensionAttribute1").Value.ToString().Trim()
            End If

            If extensionAttribute1 <> "" Then
                Commun.Journal("Dates de desactivation/suppression conservees car extensionAttribute1 est renseigne : " & dirEntry.Properties("sAMAccountName").Value, False)
                Exit Sub
            End If

            dirEntry.Properties("accountDeletionDate").Clear()
            dirEntry.Properties("accountDeletionDT").Clear()
            dirEntry.Properties("accountDeactivationDT").Clear()

            Commun.AppliquerChangement(dirEntry)

            If userAD IsNot Nothing Then
                userAD.accountDeactivationDT = Nothing
                userAD.accountDeletionDT = Nothing
                userAD.accountDeletionDate = ""
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : ClearAttributDeactivationDeletionDate : " & dirEntry.Properties("sAMAccountName").Value, True)
        End Try
    End Sub

    Private Shared Sub MoveToOu(ByVal dirEntry As DirectoryEntry,
                            ByVal userAD As UtilisateurADIndex,
                            ByVal destinationOu As String)

        Dim ancienDn As String = userAD.distinguishedName
        Dim login As String = userAD.samAccountName

        Using ouDest As New DirectoryEntry("LDAP://" & Commun.LdapPath(destinationOu), Commun.admin, Commun.passwd, auth)
            dirEntry.MoveTo(ouDest)
            dirEntry.RefreshCache(New String() {"distinguishedName"})

            If dirEntry.Properties.Contains("distinguishedName") Then
                userAD.distinguishedName = CStr(dirEntry.Properties("distinguishedName").Value)
            End If
        End Using

        Commun.Journal(vbTab & "Déplacement OU réussi : " & login &
                    " : " & ancienDn &
                    " -> " & userAD.distinguishedName, False)
    End Sub

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
    Shared Function PresenceBDP(ByVal employeeID As String) As Boolean
        If String.IsNullOrWhiteSpace(employeeID) Then Return False
        If usersRH Is Nothing Then Return False

        employeeID = employeeID.Trim()

        For Each userRH As UtilisateurRH In usersRH
            If userRH IsNot Nothing AndAlso
            Not String.IsNullOrWhiteSpace(userRH.employeeID_id) AndAlso
            String.Equals(userRH.employeeID_id.Trim(), employeeID, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Return False
    End Function

    Shared Sub GestionGroupeUserActive(ByVal dirEntry1 As DirectoryEntry)

        Dim CNUser As String = Replace(dirEntry1.Path, "LDAP://" & Commun.LdapServerPrefix(), "")
        Try
            'ajout au groupe utilisa. du domaine 
            If Commun.AppartientGroup(CNUser, "Utilisa. du domaine") = False Then
                Commun.AddRemoveADGroup(CNUser, "Utilisa. du domaine", "Add")
            End If

            'Definition du groupe utilisa. du domaine comme groupe principal
            Dim guid = Commun.PrimaryGroupId(Commun.FindAttribut("Utilisa. du domaine", "objectSid"))
            If dirEntry1.Properties("primaryGroupID").Value <> guid Then
                dirEntry1.Properties("primaryGroupID").Value = guid
                Commun.AppliquerChangement(dirEntry1)
            End If


            'Retrait du groupe G_Domain_DisableOpenSession
            If Commun.AppartientGroup(CNUser, "G_Domain_DisableOpenSession") = True Then
                Commun.AddRemoveADGroup(CNUser, "G_Domain_DisableOpenSession", "Remove")
            End If


            Commun.AddRemoveADGroup(CNUser, "G_Utilisateurs sous contrat", "Add")
            Commun.AddRemoveADGroup(CNUser, "G_Acces VPN-SSTP", "Add")
            'Commun.AddRemoveADGroup(CNUser, "LicenceO365ProPlus", "Add")
            Commun.AddRemoveADGroup(CNUser, "Internes", "Add")
        Catch ex As Exception
            Commun.Journal("ERREUR : réactivation du compte(modification des groupes) : " & CNUser, True)
        End Try
    End Sub
    Shared Sub GestionGroupeUserDesactive(ByVal dirEntry1 As DirectoryEntry)

        Dim CNUser As String = Replace(dirEntry1.Path, "LDAP://" & Commun.LdapServerPrefix(), "")

        Try

            'Ajout au groupe G_Domaine_DisableOpenSession
            If Commun.AppartientGroup(CNUser, "G_Domain_DisableOpenSession") = False Then
                Commun.AddRemoveADGroup(CNUser, "G_Domain_DisableOpenSession", "Add")
                'definition du groupe G_Domain_DisableOpenSession comme groupe principal
                dirEntry1.Properties("primaryGroupID").Value = Commun.PrimaryGroupId(Commun.FindAttribut("G_Domain_DisableOpenSession", "objectSid"))
                Commun.AppliquerChangement(dirEntry1)
            End If

            'Retrait du groupe Utilisa. du domaine
            If Commun.AppartientGroup(CNUser, "Utilisa. du domaine") = True Then
                Commun.AddRemoveADGroup(CNUser, "Utilisa. du domaine", "Remove")
            End If

            Commun.AddRemoveADGroup(CNUser, "G_Utilisateurs sous contrat", "Remove")
            Commun.AddRemoveADGroup(CNUser, "G_Acces VPN-SSTP", "Remove")
            'Commun.AddRemoveADGroup(CNUser, "LicenceO365ProPlus", "Remove")
            Commun.AddRemoveADGroup(CNUser, "Internes", "Remove")
            Commun.AddRemoveADGroup(CNUser, "phd", "Remove")
            Commun.AddRemoveADGroup(CNUser, "postdoc", "Remove")
            Commun.AddRemoveADGroup(CNUser, "Seafile-IGBMC", "Remove")

            dirEntry1.Properties("employeeNumber").Clear()
            dirEntry1.Properties("serialNumber").Clear()

            removeAdopte(Commun.TransformeSAMACCOUNTenCN(CNUser))

        Catch ex As Exception
            Commun.Journal("ERREUR : désactivation du compte(modification des groupes) : " & Replace(CNUser, "LDAP://" & Commun.LdapServerPrefix(), ""), True)
        End Try
    End Sub
    Shared Sub removeAdopte(ByVal login As String)
        Dim results As SearchResultCollection = Commun.SearchFilterAll("OU=Equipes,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr", "accountNameHistory=" & login, SearchScope.Subtree, "accountNameHistory")
        For Each result As SearchResult In results
            Using group As DirectoryEntry = New DirectoryEntry(result.Path)

                group.Properties("accountNameHistory").Remove(login)

                Commun.AppliquerChangement(group)
            End Using
        Next

    End Sub
End Class
