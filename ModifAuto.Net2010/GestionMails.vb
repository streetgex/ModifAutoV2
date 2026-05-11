Module GestionMails
    ''' <summary>
    ''' Vérifie si un mail doit être envoyé aux assistants de prévention,
    ''' puis déclenche l'envoi si nécessaire.
    ''' </summary>
    ''' <param name="usrPrenom">Prénom de la personne concernée.</param>
    ''' <param name="usrNom">Nom de la personne concernée.</param>
    ''' <param name="usrID">Identifiant de la personne dans IGBMCSERVICES / GDPI.</param>
    ''' <param name="EqDescr">Libellé du service ou de l'équipe.</param>
    ''' <param name="usrLogin">Login de la personne concernée.</param>
    ''' <param name="finContrat">Date de fin de contrat à transmettre dans le mail.</param>
    ''' <param name="createdUser">
    ''' Indique le contexte du contrôle.
    ''' <c>True</c> pour une création d'utilisateur ;
    ''' <c>False</c> pour une mise à jour.
    ''' </param>
    ''' <remarks>
    ''' La décision d'envoi repose sur <c>GetContractsLenght</c>.
    ''' L'envoi effectif est réalisé via <c>SendMailAP</c>.
    ''' </remarks>
    Public Sub EnvoyerMailAPSiNecessaire(
                    ByVal usrPrenom As String,
                    ByVal usrNom As String,
                    ByVal usrID As String,
                    ByVal eqDescr As String,
                    ByVal usrLogin As String,
                    ByVal finContrat As String,
                    ByVal checkMode As Boolean
    )
        Try
            Dim ctrlEnvoiMailAP As Boolean = GetContractsLenght(usrID, checkMode)

            If ctrlEnvoiMailAP = True Then
                SendMailAP(usrPrenom, usrNom, usrID, eqDescr, usrLogin, finContrat)
            End If
        Catch ex As Exception
            Commun.Journal("ERREUR : Envoi de mail Assistants de prévention : " & usrLogin & " : " & ex.Message, True)
        End Try
    End Sub
    ''' <summary>
    ''' Vérifie si un mail doit être envoyé aux assistants de prévention
    ''' dans le cadre d'une modification de date de fin de contrat, puis déclenche l'envoi si nécessaire.
    ''' </summary>
    ''' <param name="userRH">Utilisateur issu des données RH.</param>
    ''' <param name="userAD">État actuel de l'utilisateur dans Active Directory.</param>
    ''' <remarks>
    ''' Aucun mail n'est envoyé si une date de fin de contrat apparaît pour la première fois
    ''' alors qu'aucune date n'était encore présente dans l'AD.
    ''' La décision d'envoi repose sur <c>GetContractsLenght</c>.
    ''' L'envoi effectif est réalisé via <c>SendMailAP</c>.
    ''' </remarks>
    Public Sub EnvoyerMailAPSiNecessaire(ByVal userRH As UtilisateurRH, ByVal userAD As UtilisateurADIndex)
        If userRH.extensionAttribute1_finDeContrat <> "" AndAlso If(userAD Is Nothing, "", userAD.extensionAttribute1) = "" Then
            Exit Sub
        End If

        Dim ctrlEnvoiMailAP As Boolean = GetContractsLenght(userRH.employeeID_id, False)
        If ctrlEnvoiMailAP = False Then
            Exit Sub
        End If

        SendMailAP(
            userRH.prenom_givenName,
            userRH.nom_sn,
            userRH.employeeID_id,
            userRH.department_destinationNomLong,
            userRH.login_samAccountName,
            userRH.extensionAttribute1_finDeContrat
        )
    End Sub
    Public Sub SendMailAP(ByVal usrPrenom As String, ByVal usrNom As String, ByVal usrID As String, ByVal EqDescr As String, ByVal usrLogin As String, ByVal finContrat As String)
        Dim corpmailAssistentsPrévention =
                                                 vbCrLf & "Nom : " & usrPrenom & " " & usrNom &
                                                 vbCrLf & "Identifiant GDPI : " & usrID &
                                                 vbCrLf & "Service : " & EqDescr &
                                                 vbCrLf & "Mail : " & usrLogin & "@igbmc.fr" &
                                                 vbCrLf & "Date de fin de contrat : " & finContrat

        Commun.SendEmail("administrateur@igbmc.fr", "assistants-de-prevention@igbmc.fr;Bcc:steph@igbmc.fr", "(Mail automatique) Nouvel entrant", corpmailAssistentsPrévention)
    End Sub

    Public Sub AjouterNotificationOfficierOrienteur(
    ByVal userRH As UtilisateurRH,
    ByVal oldDate As String,
    ByVal newDate As String,
    ByVal habilitationsExpirées As String,
    ByRef ctrlMailOOrienteurs As Boolean,
    ByRef corpmailOOrienteurs As String
)

        If habilitationsExpirées = "" Then
            Exit Sub
        End If

        If ctrlMailOOrienteurs = False Then
            corpmailOOrienteurs =
            "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" &
            "<html>" &
            "<head>" &
            "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" &
            "<title>ARRET DU COMPTE</title>" &
            "</head>" &
            "<body>" &
            "Attention, la date de fin de contrat de cette/ces personne(s) a changée : <BR><BR>"
        End If

        ctrlMailOOrienteurs = True
        corpmailOOrienteurs &= ConstruireBlocMailOfficierOrienteur(userRH, oldDate, newDate, habilitationsExpirées)

        Commun.Journal(
        vbTab & "ajout de modification pour les Officiers orienteurs pour un changement de fin de contrat : " & userRH.login_samAccountName
    )

    End Sub

    Private Function ConstruireBlocMailOfficierOrienteur(
    ByVal userRH As UtilisateurRH,
    ByVal oldDate As String,
    ByVal newDate As String,
    ByVal habilitationsExpirées As String
) As String

        Return vbCrLf & "Nom : " & userRH.displayName & "<BR>" &
           vbCrLf & "Matricule : " & userRH.employeeID_id & "<BR>" &
           vbCrLf & "Service : " & userRH.department_destinationNomLong & "<BR>" &
           vbCrLf & "Mail : " & userRH.login_samAccountName & "@igbmc.fr" & "<BR>" &
           vbCrLf & "Ancienne date de fin de contrat : " & oldDate & "<BR>" &
           vbCrLf & "Nouvelle date de fin de contrat : " & newDate & "<BR>" &
           vbCrLf & "Habilitation(s) concernée(s) : " & habilitationsExpirées & "<BR>" &
           vbCrLf & "Certaines habilitations de cette personne n'iront pas jusqu'a la fin de son contrat." & "<BR>" &
           vbCrLf & "____________________________________________________________________________________________<BR>" & vbCrLf

    End Function

End Module
