Imports System.DirectoryServices

Public Class UtilisateurRH
    Public Property nom_sn As String
    Public Property prenom_givenName As String
    Public Property login_samAccountName As String
    Public Property departmentNumber_destinationNomCourt As String
    Public Property department_destinationNomLong As String
    Public Property tempsTravail As String

    Public Property equipeInfo As String
    Public Property equipeInfoDico As EquipeInfo
    Public Property manager_IDNplus1 As String = ""
    Public Property employeeID_id As String
    Public Property msExchExtensionAttribute16_aliasMailLong As String
    Public Property listesDiffusions As String
    Public Property division_organism As String
    Public Property extensionAttribute2_genre As String
    Public Property thumbnailPhoto_diffusionPhoto As String
    Public Property title_unité As String
    Public Property telephoneNumber_telPrincipal As String
    'Public Property ipPhone As String
    Public Property otherTelephone_telSecondaire As String()
    Public Property mobile_telPortable As String
    Public Property employeeNumber_employeeNumber As String
    Public Property serialNumber_serialNumber As String()
    Public Property uidNumber_uidNumber As String
    Public Property jpegPhoto_jpegPhoto As Byte()
    Public Property memberOf_groupeRH As String()
    Public Property destinationsRH As String()
    Private _physicalDeliveryOfficeName_bureaux As String = ""
    Private _extensionAttribute1_finDeContrat As String = ""
    Private _accountDeactivationDT_finDeContrat As Date?
    Private _accountDeletionDate_finDeContratPlus3Mois As String = ""
    Private _DNNplus1 As String
    Private _accountDeletionDT_finDeContratPlus3Mois As Date?
    Public Property accountDeletionDT_finDeContratPlus3Mois As Date?
        Get
            Return _accountDeletionDT_finDeContratPlus3Mois
        End Get
        Set(value As Date?)
            _accountDeletionDT_finDeContratPlus3Mois = value
        End Set
    End Property
    Public Property accountDeletionDate_finDeContratPlus3Mois As String
        Get
            Return _accountDeletionDate_finDeContratPlus3Mois
        End Get
        Set(value As String)
            _accountDeletionDate_finDeContratPlus3Mois = If(value, "").Trim()
        End Set
    End Property
    Public Property extensionAttribute1_finDeContrat As String
        Get
            Return _extensionAttribute1_finDeContrat
        End Get
        Set(value As String)
            Dim valeurNettoyee As String = If(value, "").Trim()
            _extensionAttribute1_finDeContrat = valeurNettoyee

            Dim dateFin As Date
            If Date.TryParseExact(valeurNettoyee,
                              "dd/MM/yyyy",
                              System.Globalization.CultureInfo.InvariantCulture,
                              System.Globalization.DateTimeStyles.None,
                              dateFin) Then
                accountDeactivationDT_finDeContrat = dateFin.Date
                accountDeletionDate_finDeContratPlus3Mois = dateFin.AddMonths(3).ToString("dd/MM/yyyy")
                accountDeletionDT_finDeContratPlus3Mois = dateFin.AddMonths(3).Date
            Else
                accountDeactivationDT_finDeContrat = Nothing
                accountDeletionDate_finDeContratPlus3Mois = ""
                accountDeletionDT_finDeContratPlus3Mois = Nothing
            End If
        End Set
    End Property

    Public Property accountDeactivationDT_finDeContrat As Date?
        Get
            Return _accountDeactivationDT_finDeContrat
        End Get
        Set(value As Date?)
            _accountDeactivationDT_finDeContrat = value
        End Set
    End Property
    Public Property DNNplus1 As String
        Get
            Return _DNNplus1
        End Get
        Set(value As String)
            Dim valeurNettoyee As String

            If value Is Nothing Then
                valeurNettoyee = ""
            Else
                valeurNettoyee = value.Trim()
            End If

            If valeurNettoyee <> "" Then
                If Me.login_samAccountName = directeurLogin Then
                    valeurNettoyee = ""
                End If
            End If

            _DNNplus1 = valeurNettoyee
        End Set
    End Property
    Public Property physicalDeliveryOfficeName_bureaux As String
        Get
            Return _physicalDeliveryOfficeName_bureaux
        End Get
        Set(value As String)
            Dim valeurNettoyee As String

            If value Is Nothing Then
                valeurNettoyee = ""
            Else
                valeurNettoyee = value.Trim()
            End If

            If valeurNettoyee <> "" Then
                valeurNettoyee = Microsoft.VisualBasic.Replace(valeurNettoyee, "BATIMENT", "", 1, -1, CompareMethod.Text).Trim()
                valeurNettoyee = Microsoft.VisualBasic.Replace(valeurNettoyee, "E.S.B.S.", "ESBS", 1, -1, CompareMethod.Text).Trim()
                valeurNettoyee = Microsoft.VisualBasic.Replace(valeurNettoyee, "; ", ";", 1, -1, CompareMethod.Text).Trim()
            End If


            Dim tabBureaux As String() = Split(valeurNettoyee, ";")
            Dim tabBureauxUniques As String() = Nothing

            For Each bureau As String In tabBureaux
                Dim bureauNettoye As String = bureau.Trim()
                If bureauNettoye <> "" Then
                    tabBureauxUniques.Add(bureauNettoye)
                End If
            Next

            If tabBureauxUniques Is Nothing Then
                valeurNettoyee = ""
            Else
                valeurNettoyee = Join(tabBureauxUniques, ";")
            End If


            _physicalDeliveryOfficeName_bureaux = valeurNettoyee
        End Set
    End Property
    Public ReadOnly Property mail_mailPrincipal As String
        Get
            Return Me.login_samAccountName & "@igbmc.fr"
        End Get
    End Property
    Public ReadOnly Property displayName As String
        Get
            Return prenom_givenName & " " & nom_sn
        End Get
    End Property

    Public ReadOnly Property displayNamePrintable As String
        Get
            Return nom_sn & " " & prenom_givenName
        End Get
    End Property
    Public ReadOnly Property cn As String
        Get
            Return prenom_givenName & " " & nom_sn
        End Get
    End Property
    Public ReadOnly Property unixHomeDirectoryAD_unixHomeDirectoryAD As String
        Get
            Return LCase("/shared/home/") & Me.login_samAccountName
        End Get
    End Property
    Public ReadOnly Property uid_uid As String
        Get
            Return Me.login_samAccountName
        End Get
    End Property
    Public ReadOnly Property gidNumberAD_gidNumber As String
        Get
            If equipeInfoDico Is Nothing Then Return ""
            Return equipeInfoDico.id_equipeinfo
        End Get
    End Property

End Class

Public Class UtilisateurADIndex
    Public Property employeeID As String
    Public Property departmentNumber As String
    Public Property sn As String
    Public Property givenName As String
    Public Property mail As String
    Public Property mailNickname As String
    Public Property physicalDeliveryOfficeName As String
    Public Property telephoneNumber As String
    Public Property manager As String
    Public Property division As String
    Public Property department As String
    Public Property title As String
    Public Property extensionAttribute1 As String
    Public Property extensionAttribute2 As String
    Public Property company As String = "IGBMC"
    Public Property otherTelephone As String()
    Public Property mobile As String
    Public Property whenCreated As Date?
    Public Property uidNumber As String
    Public Property samAccountName As String
    Public Property loginShell As String
    Public Property uid As String
    Public Property gidNumber As String
    Public Property unixHomeDirectory As String
    Public Property cn As String
    Public Property msExchExtensionAttribute16 As String
    Public Property displayNamePrintable As String
    Public Property displayName As String
    'Public Property memberOf_listPhdPostdoc As String
    Public Property accountActivationDT As Date?
    Public Property accountDeactivationDT As Date?
    Public Property accountDeletionDT As Date?
    Public Property accountDeletionDate As String
    Public Property employeeNumber As String
    Public Property serialNumber As String()
    Public Property comment As String
    Public Property description As String
    Public Property distinguishedName As String
    Public Property jpegPhoto As Byte()
    Public Property memberOf As String()
    Public Property memberOf_groupesGerables As String()

    Public ReadOnly Property parentDistinguishedName As String
        Get
            If String.IsNullOrWhiteSpace(distinguishedName) Then Return ""

            Dim pos As Integer = distinguishedName.IndexOf(","c)
            If pos = -1 OrElse pos >= distinguishedName.Length - 1 Then Return ""

            Return distinguishedName.Substring(pos + 1)
        End Get
    End Property
    Public ReadOnly Property isInOUActifs As Boolean
        Get
            Return String.Equals(parentDistinguishedName, OUUtilisateursActifs, StringComparison.OrdinalIgnoreCase)
        End Get
    End Property
    Public ReadOnly Property isInOUExceptions As Boolean
        Get
            Return String.Equals(parentDistinguishedName, OUUtilisateursExceptions, StringComparison.OrdinalIgnoreCase)
        End Get
    End Property

    Public ReadOnly Property isInOUDesactives As Boolean
        Get
            Return String.Equals(parentDistinguishedName, OUUtilisateursDesactives, StringComparison.OrdinalIgnoreCase)
        End Get
    End Property
    Public ReadOnly Property isInOUSortis As Boolean
        Get
            Return String.Equals(parentDistinguishedName, OUUtilisateursSortis, StringComparison.OrdinalIgnoreCase)
        End Get
    End Property
End Class

Public Class UserRaw
    Public Property lastname As String
    Public Property firstname As String
    Public Property login As String
    Public Property Dest_short_name As String
    Public Property dest_name As String
    Public Property time_rate As String = "100"
    Public Property BatimentUser As String
    Public Property TelUser As String
    Public Property equipeUser As String
    Public Property Nplus1ID As String
    Public Property IDuser As String
    Public Property aliasMail As String
    Public Property ld As String
    Public Property organism As String
    Public Property finContrat As String
    Public Property genre As String
    Public Property diffusionPhotoInterne As String
    Public Property uniteName As String
    Public Property destinationsRH As String()

    Public ReadOnly Property TimeRateValue As Integer
        Get
            Dim v As Integer = 0
            Integer.TryParse(time_rate, v)
            Return v
        End Get
    End Property
End Class
Public Class ChangementAttributAD
    Public Property Attribut As String
    Public Property AncienneValeur As String
    Public Property NouvelleValeur As String
End Class

'             0                 1               2                   3                   4               5               6                   7               8              
'lineJson = lastname & "," & firstname & "," & login & "," & Dest_short_name & "," & dest_name & "," & "100" & "," & BatimentUser & "," & TelUser & "," & equipeUser & "," & 
'    9              10              11              12          13                  14              15                    16                              17
'Nplus1ID & "," & IDuser & "," & aliasMail & "," & ld & "," & organism & "," & finContrat & "," & genre & "," & diffusionPhotoInterne.ToString & "," & uniteName
