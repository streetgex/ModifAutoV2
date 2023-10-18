Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Collections.ObjectModel


Public Class Pws

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
    Shared Sub commandePWSMailbox(ByVal login As String, ByVal db As String)
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
            Dim pCommand2 As PSCommand
            Dim pResult2 As Collection(Of PSObject)

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
                .AddCommand("Enable-mailbox")
                .AddParameter("identity", login)
                .AddParameter("alias", login)
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
                .AddParameter("identity", login)
                .AddParameter("ActiveSyncEnabled", True)
                .AddParameter("ImapEnabled", False)
                .AddParameter("PopEnabled", False)
                .AddParameter("DomainController", ctrlDomain)
            End With

            pShell.Commands = pCommand1
            pResult1 = pShell.Invoke

            pCommand2 = New PSCommand

            With pCommand2
                .AddCommand("set-MailboxCalendarConfiguration")
                .AddParameter("identity", login)
                .AddParameter("FirstWeekOfYear", "FirstFourDayWeek")
                '.AddParameter("WeatherUnit", "Celsius")
                .AddParameter("DomainController", ctrlDomain)
            End With

            pShell.Commands = pCommand2
            pResult2 = pShell.Invoke

            pRunspace.Close()
            pRunspace.Dispose()
            pRunspace = Nothing

            Commun.Journal("Creation de la boite mail Réussie: " & login)
        Catch e As Exception
            Commun.Journal("ERREUR : Creation de compte : Creation du compte mail: " & e.Message & " : " & login, True)
        End Try
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

    Sub New()
        Dim pCredential As PSCredential
        Dim pConnectionInfo As WSManConnectionInfo
        'Dim pRunspace As Runspace
        pCredential = DirectCast(Nothing, PSCredential)
        'pCredential = New PSCredential("IGBMC\userprog", CreateSecurePasswordString("FV,k,~?qa3 8ESYjYF9%"))

        '-- set connection info
        Dim server As String = ini.ReadValue("MODIFAUTO", "CasExchangeServer")
        pConnectionInfo = New WSManConnectionInfo(New Uri("http://" & server & "/powershell"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", pCredential)

        '-- create remote runspace
        runSpace = RunspaceFactory.CreateRunspace(pConnectionInfo)
        'runSpace.InitialSessionState.LanguageMode = PSLanguageMode.FullLanguage
        runSpace.Open()


        '' runspace
        'runSpace = RunspaceFactory.CreateRunspace()
        '' open it
        'runSpace.Open()
        'exec("Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin")
    End Sub

    Function exec(ByVal cmd As String) As Collection(Of PSObject)
        pipeLine = runSpace.CreatePipeline()
        pipeLine.Commands.AddScript(cmd)

        Dim outobjects = ""
        Dim psObjCollection As Collection(Of PSObject) = pipeLine.Invoke


        'Dim result(,) As String
        'Dim i As Integer = 0
        'For Each item As PSObject In psObjCollection
        '    'result = item.Members("Identity").Value
        '    ReDim Preserve result(1, 0)
        '    result(0, 0) = item.("Name").Value
        '    result(1, 0) = item.Members("Value").Value
        'Next
        Return psObjCollection
    End Function
    Function exec(ByVal cmd As String, ByVal propertieToReturn As String) As String
        pipeLine = runSpace.CreatePipeline()
        pipeLine.Commands.AddScript(cmd)

        Dim outobjects = ""
        Dim psObjCollection As Collection(Of PSObject) = pipeLine.Invoke
        'Dim outerror = pipeLine.Error.ReadToEnd
        'If outerror.Count > 0 Then
        '    Dim allerrors As String = ""
        '    For Each elem As PSObject In outerror
        '        Dim erreur = DirectCast(elem.BaseObject, Management.Automation.ErrorRecord)
        '        allerrors &= erreur.Exception.Message

        '    Next
        '    Throw New Exception(allerrors)
        'End If
        Dim result As String = Nothing
        For Each item As PSObject In psObjCollection
            'result = item.Members("Identity").Value
            result = item.Members(propertieToReturn).Value
        Next
        Return result
    End Function

    Sub close()
        runSpace.Close()
    End Sub
End Class