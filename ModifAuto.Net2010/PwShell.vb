Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Collections.ObjectModel


Public Class Pws
    Shared Sub commandePWSMailbox(ByVal login As String, ByVal db As String)
        Dim ctrlDomain As String = Commun.DCName
        Dim exchangeServer As String = ini.ReadValue("MODIFAUTO", "CasExchangeServer")

        Try
            If String.IsNullOrWhiteSpace(login) Then
                Throw New ArgumentException("Le login est vide.", "login")
            End If

            If String.IsNullOrWhiteSpace(db) Then
                Throw New ArgumentException("La base mailbox est vide.", "db")
            End If

            If String.IsNullOrWhiteSpace(exchangeServer) Then
                Throw New Exception("Le serveur Exchange n'est pas renseigne dans MODIFAUTO/CasExchangeServer.")
            End If

            Dim pCredential As PSCredential = Nothing

            Dim connectionUri As New Uri("http://" & exchangeServer & "/powershell")
            Dim pConnectionInfo As New WSManConnectionInfo(
            connectionUri,
            "http://schemas.microsoft.com/powershell/Microsoft.Exchange",
            pCredential
        )

            Using pRunspace As Runspace = RunspaceFactory.CreateRunspace(pConnectionInfo)
                pRunspace.Open()

                Dim pCommand As New PSCommand()
                With pCommand
                    .AddCommand("Enable-Mailbox")
                    .AddParameter("Identity", login)
                    .AddParameter("Alias", login)
                    .AddParameter("Database", db)
                End With
                InvokeExchangeCommand(pRunspace, pCommand, "Enable-Mailbox", ctrlDomain)


                Dim pCommand1 As New PSCommand()
                With pCommand1
                    .AddCommand("Set-CASMailbox")
                    .AddParameter("Identity", login)
                    .AddParameter("ActiveSyncEnabled", True)
                    .AddParameter("ImapEnabled", False)
                    .AddParameter("PopEnabled", False)
                End With
                TryInvokeExchangeCommand(pRunspace, pCommand1, "Set-CASMailbox", ctrlDomain)


                Dim pCommand4 As New PSCommand()
                With pCommand4
                    .AddCommand("Set-Mailbox")
                    .AddParameter("Identity", login)
                    .AddParameter("ArchiveQuota", "20GB")
                    .AddParameter("ArchiveWarningQuota", "19GB")
                End With
                TryInvokeExchangeCommand(pRunspace, pCommand4, "Set-Mailbox", ctrlDomain)

                If TryWaitMailboxCalendarReady(pRunspace, login, ctrlDomain) Then
                    Dim pCommand2 As New PSCommand()
                    With pCommand2
                        .AddCommand("Set-MailboxCalendarConfiguration")
                        .AddParameter("Identity", login)
                        .AddParameter("FirstWeekOfYear", "FirstFourDayWeek")
                    End With
                    TryInvokeExchangeCommand(pRunspace, pCommand2, "Set-MailboxCalendarConfiguration", ctrlDomain)

                    Dim pCommand3 As New PSCommand()
                    With pCommand3
                        .AddCommand("Set-MailboxRegionalConfiguration")
                        .AddParameter("Identity", login)
                        .AddParameter("TimeZone", "Romance Standard Time")
                        .AddParameter("Language", "fr-FR")
                        .AddParameter("LocalizeDefaultFolderName", True)
                    End With
                    TryInvokeExchangeCommand(pRunspace, pCommand3, "Set-MailboxRegionalConfiguration", ctrlDomain)
                Else
                    Commun.Journal("ATTENTION : mailbox creee mais configuration calendrier/regionale differee : " & login & vbCrLf & CommandesConfigurationMailboxDifferee(login, ctrlDomain), True)
                End If
            End Using

            Commun.Journal("Creation de la boite mail reussie : " & login)

        Catch ex As Exception
            Commun.Journal("ERREUR : Creation de compte : Creation du compte mail : " & ex.Message & " : " & login, True)
        End Try
    End Sub
    Private Shared Function TryWaitMailboxCalendarReady(pRunspace As Runspace, login As String, ctrlDomain As String) As Boolean
        For tentative As Integer = 1 To 24
            Try
                Dim pTest As New PSCommand()
                With pTest
                    .AddCommand("Get-MailboxCalendarConfiguration")
                    .AddParameter("Identity", login)
                End With

                InvokeExchangeCommand(pRunspace, pTest, "Get-MailboxCalendarConfiguration", ctrlDomain)

                Return True

            Catch ex As Exception
                System.Threading.Thread.Sleep(5000)
            End Try
        Next

        Commun.Journal("ATTENTION : mailbox non prete pour configuration calendrier : " & login & vbCrLf & CommandesConfigurationMailboxDifferee(login, ctrlDomain), True)
        Return False
    End Function
    Private Shared Function TryInvokeExchangeCommand(
    ByVal runspace As Runspace,
    ByVal pCommand As PSCommand,
    ByVal commandName As String,
    Optional ByVal ctrlDomain As String = Nothing
) As Boolean

        Try
            InvokeExchangeCommand(runspace, pCommand, commandName, ctrlDomain)
            Return True
        Catch ex As Exception
            Commun.Journal("ATTENTION : " & commandName & " non applique : " & ex.Message, True)
            Return False
        End Try
    End Function
    Private Shared Function DescribeExchangeCommand(ByVal pCommand As PSCommand) As String
        Dim commandText As New System.Text.StringBuilder()

        For Each command As Command In pCommand.Commands
            If commandText.Length > 0 Then commandText.Append(" | ")

            commandText.Append(command.CommandText)

            For Each parameter As CommandParameter In command.Parameters
                commandText.Append(" -")
                commandText.Append(parameter.Name)

                If parameter.Value IsNot Nothing Then
                    commandText.Append(" ")
                    commandText.Append(FormatPowerShellValue(parameter.Value))
                End If
            Next
        Next

        Return commandText.ToString()
    End Function
    Private Shared Function FormatPowerShellValue(ByVal value As Object) As String
        If value Is Nothing Then Return "$null"
        If TypeOf value Is Boolean Then Return If(CBool(value), "$true", "$false")

        Dim textValue As String = value.ToString()
        Return "'" & textValue.Replace("'", "''") & "'"
    End Function
    Private Shared Function InvokeExchangeCommand(
    ByVal runspace As Runspace,
    ByVal pCommand As PSCommand,
    ByVal commandName As String,
    Optional ByVal ctrlDomain As String = Nothing
) As Collection(Of PSObject)

        Using pShell As PowerShell = PowerShell.Create()
            pShell.Runspace = runspace

            If Not String.IsNullOrWhiteSpace(ctrlDomain) Then
                pCommand.AddParameter("DomainController", ctrlDomain)
            End If

            Dim commandText As String = DescribeExchangeCommand(pCommand)
            pShell.Commands = pCommand

            Dim result As Collection(Of PSObject)
            Try
                result = pShell.Invoke()
            Catch ex As Exception
                Throw New Exception("Erreur PowerShell pendant " & commandName & " : " & commandText & " : " & ex.Message, ex)
            End Try

            If pShell.Streams.Error.Count > 0 Then
                Dim errors As New System.Text.StringBuilder()

                For Each err As ErrorRecord In pShell.Streams.Error
                    errors.AppendLine(err.ToString())
                Next

                Throw New Exception("Erreur PowerShell pendant " & commandName & " : " & commandText & " : " & errors.ToString())
            End If

            Return result
        End Using
    End Function

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


    Shared Function ControleArchiveCreee(ByVal login As String, ByVal userID As String, Optional ByVal archive As Boolean = False) As String
        Dim result As String = Nothing
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
            'pCredential = New PSCredential("IGBMC\userprog", CreateSecurePasswordString("FV,k,~?qa3 8ESYjYF9%"))

            '-- set connection info
            Dim server As String = ini.ReadValue("MODIFAUTO", "CasExchangeServer")
            pConnectionInfo = New WSManConnectionInfo(New Uri("http://" & server & "/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", pCredential)

            '-- create remote runspace
            pRunspace = RunspaceFactory.CreateRunspace(pConnectionInfo)
            'pRunspace.InitialSessionState.LanguageMode = PSLanguageMode.FullLanguage
            pRunspace.Open()

            '-- create powershell
            pShell = PowerShell.Create
            pShell.Runspace = pRunspace

            Dim jobName As String
            If archive = False Then
                jobName = login & "-" & userID & "-IGBMC"
            Else
                jobName = login & "(Archive)-" & userID & "-IGBMC"
            End If

            '-- create command
            pCommand = New PSCommand
            With pCommand
                .AddCommand("get-MailboxExportRequest")
                .AddParameter("identity", login)
            End With

            '-- add command to powershell
            pShell.Commands = pCommand

            '-- invoke the powershell
            pResult = pShell.Invoke

            For Each item As PSObject In pResult
                result = item.Members("Status").Value
            Next

            pRunspace.Close()
            pRunspace.Dispose()
            If result = "Failed" Then
                Commun.Journal("ERREUR : La création de l'archive PST a échouée, elle sera recréée a la prochaine execution du script: " & jobName, True)
                DeleteExportRequest(jobName)
            End If
            If result Is Nothing Then
                'Commun.Journal("ERREUR : La demande de creation de fichier PST n'existe pas. Verifier la création de l'archive: " & jobName, True)
            End If
            Return result

        Catch e As Exception
            Commun.Journal("ERREUR : Controle de creation de l'archive: " & e.Message & " : " & login, True)
        End Try
    End Function
    Shared Sub DeleteExportRequest(ByVal login As String, ByVal userID As String, Optional ByVal archive As Boolean = False)
        Dim jobName As String
        If archive = False Then
            jobName = login & "-" & userID & "-IGBMC"
        Else
            jobName = login & "(Archive)-" & userID & "-IGBMC"
        End If
        Dim PwsClass As New Pws
        PwsClass.DeleteExportRequestCommun(jobName)
    End Sub
    Shared Sub DeleteExportRequest(ByVal jobName As String)
        Dim PwsClass As New Pws
        PwsClass.DeleteExportRequestCommun(jobName)
    End Sub
    Sub DeleteExportRequestCommun(ByVal jobName As String)
        'Try
        'Dim ps As New PwShell

        'Dim results = ps.exec("Get-MailboxExportRequest -name " & jobName & " |remove-MailboxExportRequest -force -Confirm:$false")
        'ps.close()


        Dim result = False
        Dim ctrlDomain As String = "serv-ad1"
        Try
            Dim pCredential As PSCredential
            Dim pConnectionInfo As WSManConnectionInfo
            Dim pRunspace As Runspace
            Dim pShell As PowerShell
            Dim pCommand As PSCommand
            'Dim pCommand1 As PSCommand
            Dim pResult As Collection(Of PSObject)
            Dim pResult1 As Collection(Of PSObject)

            '-- set credentials      
            pCredential = DirectCast(Nothing, PSCredential)
            'pCredential = New PSCredential("IGBMC\userprog", CreateSecurePasswordString("FV,k,~?qa3 8ESYjYF9%"))

            '-- set connection info
            Dim server As String = ini.ReadValue("MODIFAUTO", "CasExchangeServer")
            pConnectionInfo = New WSManConnectionInfo(New Uri("http://" & server & "/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", pCredential)

            '-- create remote runspace
            pRunspace = RunspaceFactory.CreateRunspace(pConnectionInfo)
            pRunspace.Open()

            '-- create powershell
            pShell = PowerShell.Create
            pShell.Runspace = pRunspace

            '-- create command
            pCommand = New PSCommand
            With pCommand
                .AddCommand("get-MailboxExportRequest")
                .AddParameter("name", jobName)
                .AddCommand("remove-MailboxExportRequest")
                .AddParameter("force")
                .AddParameter("Confirm", False)

            End With

            '-- add command to powershell
            pShell.Commands = pCommand

            '-- invoke the powershell
            pResult = pShell.Invoke


            pRunspace.Close()
            pRunspace.Dispose()

            Commun.Journal("Suppression de la requete de création du PST: " & jobName, True)
        Catch e As Exception
            Commun.Journal("ERREUR : Suppression de la requete de création du PST: " & e.Message & " : " & jobName, True)
        End Try
    End Sub

    Public Shared Sub ForceSyncPeperCutAD(server As String)
        ' Connexion WSMan avec l’utilisateur courant
        Dim connectionInfo As New WSManConnectionInfo(
            New Uri("http://" & server & ":5985/wsman"),
            "http://schemas.microsoft.com/powershell/Microsoft.PowerShell",
            DirectCast(Nothing, PSCredential)
        )
        connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Default

        Using runspace = RunspaceFactory.CreateRunspace(connectionInfo)
            runspace.Open()
            Using ps As PowerShell = PowerShell.Create()
                ps.Runspace = runspace

                ' Commande PaperCut (appel correct avec &)
                Dim cmd As String = "& 'C:\Program Files\PaperCut MF\server\bin\win\server-command.exe' perform-user-and-group-sync"
                ps.AddScript(cmd)

                ' Exécution
                Dim results = ps.Invoke()

                ' Sortie standard
                For Each r In results
                    Commun.Journal(vbTab & "Resultat de la synchronisation de PaperCut avec l'AD : " & r.ToString(), False)
                Next

                ' Erreurs éventuelles
                If ps.Streams.Error.Count > 0 Then
                    For Each e In ps.Streams.Error
                        Commun.Journal(vbTab & "ERREUR : Resultat de la synchronisation de PaperCut avec l'AD : " & e.ToString(), True)
                    Next
                End If
            End Using
        End Using
    End Sub
    Shared Function CreateSecurePasswordString(ByVal strPassword As String) As Security.SecureString
        Dim secureStr = New Security.SecureString()

        If strPassword.Length > 0 Then

            For Each c In strPassword.ToCharArray()
                secureStr.AppendChar(c)
            Next
        End If

        Return secureStr
    End Function

    Shared Sub CommandePWSCreatePSTMailbox(ByVal login As String, ByVal userID As String, Optional ByVal archive As Boolean = False)
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
            'pCredential = New PSCredential("IGBMC\userprog", CreateSecurePasswordString("FV,k,~?qa3 8ESYjYF9%"))

            '-- set connection info
            Dim server As String = ini.ReadValue("MODIFAUTO", "CasExchangeServer")
            pConnectionInfo = New WSManConnectionInfo(New Uri("http://" & server & "/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", pCredential)

            '-- create remote runspace
            pRunspace = RunspaceFactory.CreateRunspace(pConnectionInfo)
            'pRunspace.InitialSessionState.LanguageMode = PSLanguageMode.FullLanguage
            pRunspace.Open()

            '-- create powershell
            pShell = PowerShell.Create
            pShell.Runspace = pRunspace

            Dim pstName As String = login & "-" & userID & "-IGBMC"
            If archive = True Then
                pstName = login & "(Archive)-" & userID & "-IGBMC"
            End If

            '-- create command
            pCommand = New PSCommand
            With pCommand
                .AddCommand("New-MailboxExportRequest")
                .AddParameter("Mailbox", login)
                .AddParameter("FilePath", dossierArchivePST & pstName & ".pst")
                .AddParameter("name", pstName)
                If archive = True Then
                    .AddParameter("isArchive")
                End If
            End With

            '-- add command to powershell
            pShell.Commands = pCommand

            '-- invoke the powershell
            pResult = pShell.Invoke

            pRunspace.Close()
            pRunspace.Dispose()
            Commun.Journal("Creation du fichier PST avant suppression du compte: " & login)
        Catch e As Exception
            Commun.Journal("ERREUR : Creation du fichier PST avant suppression du compte: " & e.Message & " : " & login, True)
        End Try
    End Sub

    Shared Sub commandePWSDisableMailbox(ByVal login As String)
        Dim ctrlDomain As String = "serv-ad1"
        Try
            Dim pCredential As PSCredential
            Dim pConnectionInfo As WSManConnectionInfo
            Dim pRunspace As Runspace
            Dim pShell As PowerShell
            Dim pCommand As PSCommand
            Dim pResult As Collection(Of PSObject)

            '-- set credentials      
            pCredential = DirectCast(Nothing, PSCredential) 'New PSCredential("igbmc\steph", CreateSecurePasswordString("aaaaaa"))

            '-- set connection info
            pConnectionInfo = New WSManConnectionInfo(New Uri("http://" & ini.ReadValue("MODIFAUTO", "CasExchangeServer") & "/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", pCredential)

            '-- create remote runspace
            pRunspace = RunspaceFactory.CreateRunspace(pConnectionInfo)
            'pRunspace.InitialSessionState.LanguageMode = PSLanguageMode.FullLanguage
            pRunspace.Open()

            '-- create powershell
            pShell = PowerShell.Create
            pShell.Runspace = pRunspace

            '-- create command
            pCommand = New PSCommand
            With pCommand
                .AddCommand("disable-mailbox")
                .AddParameter("identity", login)
                .AddParameter("DomainController", ctrlDomain)
                .AddParameter("Confirm", False)
            End With

            '-- add command to powershell
            pShell.Commands = pCommand

            '-- invoke the powershell
            pResult = pShell.Invoke



            pRunspace.Close()
            pRunspace.Dispose()
            pRunspace = Nothing

            Commun.Journal("Suppression de la boite mail Réussie: " & login)
        Catch e As Exception
            Commun.Journal("ERREUR : Suppression de compte : Suppression du compte mail: " & e.Message & " : " & login, True)
        End Try


    End Sub
End Class
Public Class PwShell
    Dim runSpace As Runspace
    Dim pipeLine As Pipeline

End Class
