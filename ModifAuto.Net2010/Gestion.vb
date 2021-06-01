Imports System.DirectoryServices

Imports System.IO



Public Class gestion
    Shared Sub GestionAttributsDT()

        Commun.Journal("Ecriture des Attribut accountDeletionDT et accountDeactivationDT", False)

        Const ADS_UF_ACCOUNT_DISABLE = 2
        Dim dateNowU As String = Now.Date.ToString("yyyyMMddHHmmss.sZ")


        Using OuUsers As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateurs]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)

            'accountDeletionDT
            Using searcher As DirectorySearcher = New DirectorySearcher(OuUsers)
                'Exclusion de l'OU "Out" et des comptes où la date de suppression n'est pas echue
                searcher.Filter = "(&(objectClass=user)(accountDeletionDT<=" & dateNowU & ")(!(msDS-parentdistname=" & RecupDataini.RecupVar("[OUUtilisateursSortis]") & ")))"
                searcher.PageSize = 5000
                searcher.SearchScope = SearchScope.Subtree
                Dim results As SearchResultCollection = searcher.FindAll()

                For Each user As SearchResult In results

                    Using userAD As DirectoryEntry = user.GetDirectoryEntry
                        Try
                            'Cas des utilisateurs "Interne"
                            If userAD.Parent.Path = "LDAP://" & RecupDataini.RecupVar("[OUUtilisateursDesactives]") Then
                                Supprime.SupprimCompte(userAD)
                                Commun.Journal("Suppression du compte interne : " & userAD.Properties("SAMAccountName").Value)

                                'Cas des utilisateurs "Externes" et "Provisoires"
                            ElseIf userAD.Parent.Path = "LDAP://" & RecupDataini.RecupVar("[OUUtilisateursExternes]") Or userAD.Parent.Path = "LDAP://" & RecupDataini.RecupVar("[OUUtilisateursProvisoires]") Then

                                'suppression des comptes en fonction de "accountDeletionDT"
                                If userAD.Properties.Contains("accountDeletionDT") Then
                                    Dim interval As Integer = DateDiff("d", userAD.Properties("accountDeletionDT").Value, Now.Date)
                                    If interval >= 1 Then
                                        userAD.DeleteTree()
                                        Commun.Journal("Suppression du compte : " & userAD.Properties("SAMAccountName").Value)
                                        'GoTo fermerUsing
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            Commun.Journal("ERREUR : GestionAttributsDT : Suppression du compte : " & userAD.Properties("SAMAccountName").Value)
                        End Try
                    End Using
                Next user
            End Using



            'accountDeactivationDT
            Using searcher As DirectorySearcher = New DirectorySearcher(OuUsers)
                'Exclusion des OU "Utilisateurs", "Comptes Désactivés", "Out", des comptes désactivés et où la date de desactivation n'est pas echue
                searcher.Filter = "(&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(accountDeactivationDT<=" & dateNowU & ")(!(msDS-parentdistname=" & RecupDataini.RecupVar("[OUUtilisateursSortis]") & "))(!(msDS-parentdistname=" & RecupDataini.RecupVar("[OUUtilisateursDesactives]") & "))(!(msDS-parentdistname=" & RecupDataini.RecupVar("[OUUtilisateursActifs]") & ")))"
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
                                        Commun.Journal("Désactivation du compte : " & userAD.Properties("SAMAccountName").Value)
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            Commun.Journal("ERREUR : GestionAttributsDT : Désactivation du compte : " & userAD.Properties("SAMAccountName").Value)
                        End Try
                    End Using

                Next
            End Using

            'accountActivationDT
            Using searcher As DirectorySearcher = New DirectorySearcher(OuUsers)
                'Exclusion des OU "Utilisateurs", "Comptes Désactivés", "Out", des comptes actifs et où la date de desactivation n'est pas echue et la date d'activation dépassée
                searcher.Filter = "(&(objectClass=user)((userAccountControl:1.2.840.113556.1.4.803:=2))(accountActivationDT>=" & dateNowU & ")(accountDeactivationDT<=" & dateNowU & ")(!(msDS-parentdistname=" & RecupDataini.RecupVar("[OUUtilisateursSortis]") & "))(!(msDS-parentdistname=" & RecupDataini.RecupVar("[OUUtilisateursDesactives]") & "))(!(msDS-parentdistname=" & RecupDataini.RecupVar("[OUUtilisateursActifs]") & ")))"
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
                            Commun.Journal("ERREUR : GestionAttributsDT : Activation du compte : " & userAD.Properties("SAMAccountName").Value)
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
    Shared Sub GestionSuppressionPIetDR()

        Commun.Journal("Debut de la gestion de la suppression des profils", False)

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


        For b = 0 To 1
            chemin = "\\Labo4\home1\"

            If b = 1 Then
                chemin = "\\Clust-roamingP\Profils\BRDAdministratif\"
            End If

            serveur = Replace(chemin, "\\", "")
            serveur = Strings.Left(serveur, InStr(serveur, "\") - 1)
            Dim subDirs As String() = Directory.GetDirectories(chemin)
            Array.Sort(subDirs)
            For Each dir As String In subDirs
                If Strings.Right(dir, 6) = " - old" Then Continue For
                Dim login As String = Replace(dir, chemin, "")
                Dim checkUserLogin As Boolean = VerifyUserProfil(login, "SAMAccountName")
                If checkUserLogin = False And login <> ".etc" And login <> "All Users" Then
                    'MsgBox(login)
                    If Directory.Exists(dir) Then

                        'fermer les fichiers ouverts par l'utilisateur
                        Dim ProcessProperties As New ProcessStartInfo
                        ProcessProperties.FileName = "openfiles"
                        ProcessProperties.Arguments = "/disconnect /s " & serveur & " /a " & login
                        ProcessProperties.WindowStyle = ProcessWindowStyle.Hidden
                        Dim myProcess As Process = Process.Start(ProcessProperties)
                        Try
                            ClearAttributes(chemin & login)
                            Directory.Delete(dir, True)
                            Commun.Journal("GestionPIetDR : Suppression du profil : " & dir, False)
                        Catch ex As Exception
                            Commun.Journal("ERREUR : GestionPIetDR : Suppression du profil : " & dir & " : " & ex.Message, True)
                        End Try
                    End If
                End If
            Next dir

        Next b


        'Dim checkUser As Boolean = VerifyUser(login, "objectSid")



    End Sub
    Shared Function VerifyUserProfil(ByVal value As String, ByVal attribut As String) As Boolean
        Dim result As Boolean = True
        Dim searcher As SearchResult = Commun.SearchFilterOne("DC=igbmc,DC=u-strasbg,DC=fr", "(&(objectClass=user)(" & attribut & "=" & value & ")(!(msDS-parentdistname=" & RecupDataini.RecupVar("[OUUtilisateursSortis]") & ")))", SearchScope.Subtree)
        'Dim searcher As SearchResult = Commun.SearchFilterOne("DC=igbmc,DC=u-strasbg,DC=fr", "(&(objectClass=user)(objectSid=S-1-5-21-839522115-1935655697-1343024091-10671)(!(msDS-parentdistname=" & RecupDataini.RecupVar("[OUUtilisateursSortis]") & ")))", SearchScope.Subtree)
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
                    listMembresAutorises = LCase(RecupDataini.RecupVar("[Group_Admins_du_domaine]"))
                    'listMembresAutorises = "Administrateur,steph,stephadm,tina,guiseithadm,SERV-TMG$,SERV-CLUSTER2$,SERV-EXCHANGE$,userprog,krbtgt"
                Case 1
                    samaccountGroup = "Administrateurs de l'entreprise"
                    listMembresAutorises = LCase(RecupDataini.RecupVar("[Group_Administrateurs_de_l_entreprise]"))
                    'listMembresAutorises = "Administrateur,steph,stephadm,guiseithadm,userprog"
                Case 2
                    samaccountGroup = "Administrateurs"
                    listMembresAutorises = LCase(RecupDataini.RecupVar("[Group_Administrateurs]"))
                    'listMembresAutorises = "Administrateur,steph,stephadm,Administrateurs,Admins du domaine,scripsteph"
                Case 3
                    samaccountGroup = "Administrateurs du schéma"
                    listMembresAutorises = LCase(RecupDataini.RecupVar("[Group_Administrateurs_du_schema]"))
                    'listMembresAutorises = "Administrateur,steph,stephadm"
                Case 4
                    samaccountGroup = "Administrateurs DHCP"
                    listMembresAutorises = LCase(RecupDataini.RecupVar("[Group_Administrateurs_DHCP]"))
                    'listMembresAutorises = "Administrateur,steph,stephadm,guiseithadm,userprog"
            End Select

            If InStr(listMembresAutorises, "administrateur,stephadm") <> 0 Then

                Dim tabMembresAutorises = Split(listMembresAutorises, ",")
                Dim monGroupe As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.TransformeSAMACCOUNTenCN(samaccountGroup), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                Try
                    ' Groupe dont les membres sont à lister
                    Dim tabMembre As String() = Commun.MembresDuGroupe(samaccountGroup, True)
                    'tabMembre = Split(LCase(Join(tabMembre, ";")), ";")
                    For Each unMembre In tabMembre

                        Dim userLogin As String = Commun.TransformeSAMACCOUNTenCN(unMembre.ToString) 'user.Properties("SAMAccountName").Value
                        userLogin = LCase(userLogin)

                        If Array.IndexOf(tabMembresAutorises, userLogin) = -1 Then
                            monGroupe.Invoke("Remove", New Object() {"LDAP://" & unMembre.ToString})
                        End If
                    Next unMembre

                    For i = 0 To UBound(tabMembresAutorises)
                        Dim DNUser As String = Commun.TransformeSAMACCOUNTenCN(tabMembresAutorises(i))
                        'DNUser = LCase(DNUser)
                        Dim found As Integer = Array.IndexOf(tabMembre, DNUser)
                        If found = -1 Then
                            monGroupe.Invoke("Add", New Object() {"LDAP://" & DNUser})
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
    End Sub

    Shared Sub UpdateComptesProvisoires()

        Commun.Journal("Mise a jour des utilisateurs provisoires", False)

        'Mise a jour des utilisateurs provisoires
        Using OUProvisoire As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursProvisoires]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Dim searcher As DirectorySearcher = New DirectorySearcher(OUProvisoire)
            searcher.Filter = "(&(objectClass=user))"

            Dim result As SearchResultCollection = searcher.FindAll()
            'Dim dateExpTexteDescription As String

            For Each utilisateurProv As SearchResult In result
                Using userUpdate As DirectoryEntry = utilisateurProv.GetDirectoryEntry

                    'Dim dateExpDateTxt As String = Date.ParseExact(userUpdate.Properties("accountDeactivationDT").Value, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo).Date
                    Dim dateExpDate As Date = userUpdate.Properties("accountDeactivationDT").Value
                    Dim dateExpDateTxt As String = dateExpDate.ToString("dd/MM/yyyy")
                    Try

                        Try
                            'Nettoyage des destinations sur les utilisateurs provisoires
                            For i As Integer = 0 To userUpdate.Properties("MemberOf").Count - 1
                                If InStr(userUpdate.Properties("MemberOf")(i), " grp,") > 0 Then
                                    Commun.AddRemoveADGroup(userUpdate.Properties("SAMAccountName").Value, userUpdate.Properties("MemberOf")(i), "Remove")
                                End If
                            Next
                        Catch ex As Exception
                            Commun.Journal("ERREUR : UpdateComptesProvisoires : Nettoyage destination : " & userUpdate.Properties("SAMAccountName").Value & " : " & ex.Message, True)
                        End Try

                        Try
                            'dateExpDate = Strings.Left(DateAdd(DateInterval.Day, 0, Commun.ConvertADValueToDateTime(userUpdate.Properties("accountDeactivationDT").Value)), 10)
                            If dateExpDateTxt <> Strings.Right(userUpdate.Properties("Description").Value, 10) Then
                                userUpdate.Properties("description").Value = "COMPTE PROVISOIRE expire le: " & dateExpDateTxt
                                Commun.AppliquerChangement(userUpdate)
                            End If
                        Catch ex As Exception
                            Commun.Journal("ERREUR : UpdateComptesProvisoires : Mise a jour Description :" & userUpdate.Properties("SAMAccountName").Value & " : " & ex.Message, True)
                        End Try

                    Catch ex As Exception
                        Commun.Journal("ERREUR : UpdateComptesProvisoires : " & userUpdate.Properties("SAMAccountName").Value & " : " & ex.Message, True)
                    End Try
                End Using
            Next

        End Using
    End Sub

    Shared Sub ControleOUUtilisateurs()

        Commun.Journal("Controle de l'OU Utilisateurs", False)

        'Desactivation des comptes dans l'OU Exception qui n'ont rien a y faire
        Dim OUException As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursExceptions]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
        Dim searchOUException As DirectorySearcher = New DirectorySearcher(OUException)
        searchOUException.Filter = "(&(objectClass=user))"
        searchOUException.PropertiesToLoad.Add("SamAccountName")
        Dim searchOUExceptionResult As SearchResultCollection = searchOUException.FindAll()
        For Each userOUException As SearchResult In searchOUExceptionResult
            If Commun.IsException(userOUException.Properties("SamAccountName")(0), "") = False Then
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

        Dim OUComptesDesactives As New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursDesactives]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
        Dim searchDesactiv As DirectorySearcher = New DirectorySearcher(OUComptesDesactives)
        searchDesactiv.Filter = "(&(objectClass=user))"
        Dim searchDesactivResult As SearchResultCollection = searchDesactiv.FindAll()
        For Each objUserDesactiv As SearchResult In searchDesactivResult

            Dim DirEntry As DirectoryEntry = objUserDesactiv.GetDirectoryEntry

            Try
                GestionGroupeUserDesactive(DirEntry)
            Catch
            End Try
            DirEntry.Close()
            DirEntry.Dispose()
            DirEntry = Nothing
        Next
        searchDesactiv.Dispose()
        searchDesactiv = Nothing
        searchDesactivResult.Dispose()
        searchDesactivResult = Nothing
        OUComptesDesactives.Close()
        OUComptesDesactives = Nothing
    End Sub
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
        Using OUDomDest As DirectoryEntry = New DirectoryEntry("LDAP://OU=Equipes,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Using dirDestSearcher As DirectorySearcher = New DirectorySearcher(OUDomDest)

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
                                dirDest = New DirectoryEntry("LDAP://" & Commun.TransformeSAMACCOUNTenCN(SAMAccountGroup), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                                'L'attribut "languageCode" contient le groupe ID de GDPI
                                dirDest.Properties("languageCode").Value = IdDest
                                Commun.AppliquerChangement(dirDest)
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

                        End If
                    Catch ex As Exception
                        Commun.Journal("ERREUR : Gestion des Destinations : Ajout/Modification " & SAMAccountGroup & " : " & ex.Message, True)
                    End Try

                Next l


                'gestion des departements
                Using OUDomDepart As DirectoryEntry = New DirectoryEntry("LDAP://OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                    Using dirDepartSearcher As DirectorySearcher = New DirectorySearcher(OUDomDepart)
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
                                            Commun.NouveauGroupe(OUDomDepart, SAMAccountDepart, "Departement " & department_nom)
                                        Catch

                                        End Try
                                        dirDepart = New DirectoryEntry("LDAP://" & Commun.TransformeSAMACCOUNTenCN(SAMAccountDepart), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
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

                                    dirDepart.Close()
                                    dirDepart.Dispose()
                                    dirDepart = Nothing
                                    result = Nothing
                                End If

                            Catch ex As Exception
                                Commun.Journal("ERREUR : Gestion des Departements: Ajout/Modification " & SAMAccountDepart & " : " & ex.Message, True)
                            End Try


                        Next p
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
                                        Commun.SendEmail("administrateur@igbmc.fr", RecupDataini.RecupVar("[MailEquipeReso]"), "Destination non associée a une Equipe Info", "La destination " & descriptionGroup & " (" & SAMAccountGroup & "), dans l'active Directory, n'est actuellement pas associé à une Equipe info." & vbCrLf & "Cette destination contient un ou des utilisateurs." & vbCrLf & "Tant que cette association ne sera pas faite, les utilisateurs n'auront pas d'acces a leur Zone Labo")
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            Commun.Journal("ERREUR : Gestion des Destinations : Controle : " & SAMAccountGroup & " : " & ex.Message, True)
                        End Try
                    Next grpResult
                    result1.Dispose()
                    result1 = Nothing
                End Using
            End Using
            'dirSearcher = Nothing
            '----------------------------------------------------------------------------------------------------------------------------------

        End Using
        Erase tableauDest
        Commun.Journal("Gestion des destinations et départements réussie", False)

    End Sub

    Shared Sub PIDepartementGroup()
        Using dptAD As DirectoryEntry = New DirectoryEntry("LDAP://OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Dim resultsDpt As SearchResultCollection = Commun.SearchFilterAll("OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr", "objectClass=group", SearchScope.OneLevel, "description,cn")
            For Each resultDpt As SearchResult In resultsDpt
                Dim description As String = Replace(resultDpt.Properties("description")(0), "Departement", "")
                Dim nom As String = resultDpt.Properties("cn")(0)
                Dim resultsDest As SearchResultCollection = Commun.SearchFilterAll("OU=Equipes,DC=igbmc,DC=u-strasbg,DC=fr", "(&(objectClass=group)(department=" & description & "))", SearchScope.OneLevel, "department")
                For Each resultDest As SearchResult In resultsDest
                    Dim nomGrpPI As String = Replace(nom, "Dpt_", "Dpt_PI-")
                    Using destAD As DirectoryEntry = New DirectoryEntry(resultDest.Path)
                        'MsgBox(destAD.Properties("sAMAccountName").Value)
                        Dim chefCN As String = destAD.Properties("managedBy").Value
                        If Not DirectoryEntry.Exists("LDAP://CN=" & nomGrpPI & ",OU=Departements PI,OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr") Then
                            Commun.NouveauGroupe(New DirectoryEntry("LDAP://OU=Departements PI,OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr", Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure), nomGrpPI, "")
                        End If
                        Commun.AddRemoveADGroup(chefCN, nomGrpPI, "Add")
                    End Using
                Next
            Next

        End Using
    End Sub
    Shared Sub GestionReactiveDesactiveComptesInterne()

        Dim OUSource As DirectoryEntry
        Using ouDestDesactive As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursDesactives]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Using ouDestException As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursExceptions]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
                Using ouDestUser As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursActifs]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)


                    Dim typeSearchScope As SearchScope

                    For cas = 0 To 2
                        If cas = 0 Then
                            OUSource = ouDestDesactive
                            typeSearchScope = SearchScope.Subtree
                        ElseIf cas = 1 Then
                            OUSource = ouDestException
                            typeSearchScope = SearchScope.OneLevel
                        Else
                            OUSource = ouDestUser
                            typeSearchScope = SearchScope.OneLevel
                        End If

                        Dim users As SearchResultCollection
                        Using searcher As DirectorySearcher = New DirectorySearcher(OUSource)
                            searcher.PageSize = 2000
                            searcher.SearchScope = typeSearchScope

                            searcher.Filter = "(&(objectClass=user))"
                            ' pour la recherche non recursive(SearchScope.OneLevel)
                            users = searcher.FindAll()
                        End Using

                        For Each user As SearchResult In users
                            Dim DirEntry As DirectoryEntry = user.GetDirectoryEntry

                            'If DirEntry.Properties.Contains("accountDeletionDate") Then
                            '    Dim dateSuppressionTxt As String = DirEntry.Properties("accountDeletionDate").Value
                            '    Dim dateSuppression As Date = Date.ParseExact(dateSuppressionTxt, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            '    'Dim tz As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("UTC")
                            '    'Dim aaa As Date = TimeZoneInfo.ConvertTimeToUtc(dateSuppression, tz)
                            '    DirEntry.Properties("accountDeletionDT").Value = dateSuppression
                            '    DirEntry.CommitChanges()
                            'End If

                            Dim login As String = DirEntry.Properties("SAMAccountName").Value()
                            Dim except As String = exceptionUser(login)
                            Dim ctrlPresentBDP As Boolean = PresentBDP(login)

                            'Cas d'un compte Actif par BDP ou par Exception
                            If ctrlPresentBDP = True Or except <> "False" Then
                                GestionGroupeUserActive(DirEntry)
                                'Si l'utilistateur est present dans la base du perso mais n'est pas dans l'OU Utilisateur
                                If ctrlPresentBDP = True And cas <> 2 Then
                                    Try
                                        Commun.SetADLDAPProperty(DirEntry, "description", "")
                                        Commun.SetADLDAPProperty(DirEntry, "Comment", "Réactivé le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf, True)
                                        Commun.AppliquerChangement(DirEntry)
                                        Commun.ClearAttributDeactivationDeletionDate(DirEntry)
                                        Commun.ReactiveDesactiveCompte(login, "Active")
                                    Catch ex As Exception
                                        Commun.Journal("ERREUR : réactivation du compte(propriétés) : " & DirEntry.Name & " : " & ex.Message, True)
                                    End Try

                                    Try

                                    Catch ex As Exception
                                        Commun.Journal("ERREUR : date d'expiration du compte (deplacement Vers OU Utilisateurs) : " & DirEntry.Name, True)
                                    End Try

                                    'Réactivation du compte et deplacement dans l'OU Utilisateurs

                                    'Active l'utilisateur ADM s'il existe
                                    Commun.ReactiveDesactiveCompte(login & "adm", "Active")
                                    Try
                                        Form1.CompararaisonAjoutRetraitDestinations(DirEntry)
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
                                            Commun.AppliquerChangement(DirEntry)
                                            Commun.ClearAttributDeactivationDeletionDate(DirEntry)
                                        End If
                                    Catch ex As Exception
                                        Commun.Journal("ERREUR : réactivation du compte(propriétés) : " & DirEntry.Name, True)
                                    End Try

                                    'si il n'est pas deja dans l'OU Exception
                                    If cas <> 1 Then

                                        'Active l'utilisateur ADM s'il existe
                                        Commun.ReactiveDesactiveCompte(login & "adm", "Active")

                                        GestionGroupeUserActive(DirEntry)

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
                                Dim dateDefinDeContrat As String
                                Try
                                    'si il est dans l'OU Exception
                                    If cas = 1 Then
                                        Commun.SetADLDAPProperty(DirEntry, "description", "EXCEPTION Désactivé le: " & Strings.Left(CStr(Now), 10))
                                        Commun.SetADLDAPProperty(DirEntry, "Comment", "EXCEPTION Désactivé le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf, True)
                                        Commun.SetADLDAPProperty(DirEntry, "accountDeletionDate", Strings.Left(DateTime.UtcNow.AddMonths(3).ToString, 10))
                                        DirEntry.Properties("accountDeactivationDT").Value = Now.Date
                                        DirEntry.Properties("accountDeletionDT").Value = Now.Date.AddMonths(3)
                                        Commun.AppliquerChangement(DirEntry)
                                    End If

                                    'si il est dans l'OU Utilisateurs
                                    If cas = 2 Then
                                        Commun.SetADLDAPProperty(DirEntry, "description", "Désactivé le: " & Strings.Left(CStr(Now), 10))
                                        Commun.SetADLDAPProperty(DirEntry, "Comment", "Désactivé le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf, True)
                                        Commun.SetADLDAPProperty(DirEntry, "accountDeletionDate", Strings.Left(DateTime.UtcNow.AddMonths(3).ToString, 10))
                                        DirEntry.Properties("accountDeactivationDT").Value = Now.Date
                                        DirEntry.Properties("accountDeletionDT").Value = Now.Date.AddMonths(3)
                                        Commun.AppliquerChangement(DirEntry)
                                    End If

                                    If DirEntry.Properties.Contains("EmployeeID") Then
                                        Dim id As String = DirEntry.Properties("EmployeeID").Value
                                        Dim datacontract As String = Json.SendJson("", "persons/" & id & "/contracts?current_contract=true", "AD", "GET")
                                        Dim contracts = Json.DeserializeJson(datacontract, "contracts")
                                        dateDefinDeContrat = Form1.DateDeFinDeContract(contracts, id)
                                        If dateDefinDeContrat = "" Then
                                            Commun.Journal("ERREUR : GestionReactiveDesactiveCompte : Determination de la date de fin de contrat : " & DirEntry.Name & " : EmployeeID : " & DirEntry.Properties("EmployeeID").Value, True)
                                        End If
                                        Commun.SetADLDAPProperty(DirEntry, "extensionAttribute1", dateDefinDeContrat)
                                        Commun.AppliquerChangement(DirEntry)
                                    End If

                                    Commun.SetADLDAPProperty(DirEntry, "Comment", "Désactivé le: " & Strings.Left(CStr(Now), 10) & " (ModifAuto)" & vbCrLf, True)
                                    Commun.AppliquerChangement(DirEntry)

                                Catch ex As Exception
                                    Commun.Journal("ERREUR : désactivation du compte(propriétés) : " & DirEntry.Name, True)
                                End Try

                                'Desactive l'utilisateur ADM s'il existe
                                Commun.ReactiveDesactiveCompte(login & "adm", "Desactive")

                                'Le deplacer dans l'ou Comptes Désactivés
                                GestionGroupeUserDesactive(DirEntry)
                                Form1.CompararaisonAjoutRetraitDestinations(DirEntry)

                                Dim adresseMail As String = DirEntry.Properties("Mail").Value
                                Dim prenom As String = DirEntry.Properties("givenName").Value
                                Dim dateDeSuppressionPrevue As String = Format(DirEntry.Properties("accountDeletionDT").Value, "dd/MM/yyyy")



                                Try
                                    Dim mail As String = Form1.MailCloture1(prenom, dateDeSuppressionPrevue, dateDefinDeContrat)
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

                End Using
            End Using
        End Using

        Commun.Journal("Gestion de l'activation/desactivation des comptes réussie", False)
    End Sub

    Shared Function exceptionUser(ByVal login As String) As String
        'controle si l'utilisateur fait partie des exceptions (toujours valide) en retournant la date limite de l'exception
        'sinon retourne False
        Dim resultat As String = "False"
        Try
            If Form1.tabExcepUser Is Nothing Then
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
                            ReDim Preserve Form1.tabExcepUser(1, i)
                            Form1.tabExcepUser(0, i) = tabLigneExceptUser(0)
                            Form1.tabExcepUser(1, i) = tabLigneExceptUser(1)
                        End If

                        tabLigneExceptUser = Nothing
                    End If
                Loop Until ligne Is Nothing
                monStreamReader.Close()
                'monStreamReader.Dispose()
            End If

            If Not Form1.tabExcepUser Is Nothing Then
                Dim positionTab As Integer = Form1.IndexOfMulti(Form1.tabExcepUser, login, 0)
                If positionTab = -1 Then
                    resultat = "False"
                Else
                    resultat = Form1.tabExcepUser(1, positionTab)
                End If
            End If

        Catch ex As Exception
            Commun.Journal("ERREUR : exceptionUser : login : " & login & " : " & ex.Message, True)
        End Try

        Return resultat

    End Function
    Shared Function PresentBDP(ByVal login As String) As Boolean
        For i = 0 To UBound(Form1.tabPersoMonoEquipe, 2)
            If login = Form1.tabPersoMonoEquipe(2, i) Then
                Return True
                Exit Function
            End If
        Next i
        Return False
    End Function
    
    Shared Sub GestionGroupeUserActive(ByVal dirEntry1 As DirectoryEntry)

        Dim CNUser As String = Replace(dirEntry1.Path, "LDAP://", "")
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
            Commun.AddRemoveADGroup(CNUser, "LicenceO365ProPlus", "Add")
            Commun.AddRemoveADGroup(CNUser, "Internes", "Add")
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
                Commun.AppliquerChangement(dirEntry1)
            End If

            'Retrait du groupe Utilisa. du domaine
            If Commun.AppartientGroup(CNUser, "Utilisa. du domaine") = True Then
                Commun.AddRemoveADGroup(CNUser, "Utilisa. du domaine", "Remove")
            End If

            Commun.AddRemoveADGroup(CNUser, "G_Utilisateurs sous contrat", "Remove")
            Commun.AddRemoveADGroup(CNUser, "G_Acces VPN-SSTP", "Remove")
            Commun.AddRemoveADGroup(CNUser, "LicenceO365ProPlus", "Remove")
            Commun.AddRemoveADGroup(CNUser, "Internes", "Remove")

            removeAdopte(Commun.TransformeSAMACCOUNTenCN(CNUser))

        Catch ex As Exception
            Commun.Journal("ERREUR : désactivation du compte(modification des groupes) : " & Replace(CNUser, "LDAP://", ""), True)
        End Try
    End Sub
    Shared Sub removeAdopte(ByVal login As String)
        Dim results As SearchResultCollection = Commun.SearchFilterAll("OU=Equipes,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr", "accountNameHistory=" & login, SearchScope.Subtree, "accountNameHistory")
        For Each result As SearchResult In results
            Using group As DirectoryEntry = New DirectoryEntry(result.Path)

                group.Properties("accountNameHistory").Remove(login)
                group.CommitChanges()
            End Using
        Next

    End Sub
End Class
