Imports System.DirectoryServices
Imports System.IO
Imports System.Web.Script.Serialization

Public Class ctrlMS
    Shared Sub recupUserAD()
        Try
            Kill("c:\temp\MSrapport.csv")
        Catch
        End Try
        File.AppendAllText("c:\temp\MSrapport.csv", "matricule,nom,present AD,present MS,actif MS,date de Sortie IGBMC,badge Edité,photo,photo à refaire,plusieurs badges,date fin de validité MS")

        Dim dateRef As String = "20190101000000.0Z"

        Dim countAD As Integer = 0
        Dim countMS As Integer = 0
        Dim countActiveMS As Integer = 0
        Dim countPhotoAFaire As Integer = 0
        Dim countMSACreer As Integer = 0
        Dim countMSASuppr As Integer = 0
        Dim CountMultiBadge As Integer = 0
        Dim countBadgeEdit As Integer = 0
        Dim countPhotoARefaire As Integer = 0
        Dim tabMatriculeMS As String()
        Dim tabMatriculeAD As String()

        For k = 0 To 10
            'Dim totalMS = jsonMS.MakeRequest("GET", "/users?fields=matricule&filter=matricule&limit=100&offset=" & k * 100)
            Dim totalMS = jsonMS.MakeRequest("GET", "/users?fields=matricule&filter=matricule&limit=100&offset=" & k * 100)
            Dim matriculeResponseData = New JavaScriptSerializer().Deserialize(Of Object)(totalMS)
            If matriculeResponseData("data").length = 0 Then
                Exit For
            Else
                Dim m As Integer = -1
                For Each matriculeR In matriculeResponseData("data")
                    m += 1
                    Dim matri As String = matriculeResponseData("data")(m)("matricule")
                    If Left(matri, 1) <> "E" Then
                        tabMatriculeMS.Add(matri)
                    End If
                Next
            End If
        Next

        Using AD As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateurs]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            'Dim filtre As SearchResultCollection = Commun.SearchFilterAll(AD, "(&(EmployeeID=*)(objectCategory=person)(|((&(accountDeactivationDT>=20200101000000.0Z)(msDS-parentdistname=OU=Out,OU=Comptes Désactivés,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr))(&(accountDeactivationDT>=20200101000000.0Z)(msDS-parentdistname=OU=Comptes Désactivés,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr))(msDS-parentdistname=OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr))))", SearchScope.Subtree, "employeeID")
            Dim filtre As SearchResultCollection = Commun.SearchFilterAll(AD, "(&(objectCategory=person)(employeeID=*)(|((&(accountDeactivationDT>=" & dateRef & ")(msDS-parentdistname=OU=Out,OU=Comptes Désactivés,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr))(&(accountDeactivationDT>=" & dateRef & ")(msDS-parentdistname=OU=Comptes Désactivés,OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr))(msDS-parentdistname=OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr))))", SearchScope.Subtree)
            Dim i As Integer = filtre.Count
            For Each result As SearchResult In filtre

                tabMatriculeAD.Add(result.Properties("employeeID")(0))

            Next
        End Using

        Dim diffItems() As String = tabMatriculeMS.Where(Function(i) Not tabMatriculeAD.Contains(i)).ToArray()
        'Array.ForEach(diffItems, Sub(i) Console.WriteLine(i))
        Dim tabTotal As String() = tabMatriculeAD.Union(diffItems).ToArray





        For Each employeeid In tabTotal

            Dim presentMS As Boolean = True
            Dim activeMS As Boolean = False
            Dim badgeEdit As Boolean = False
            Dim photo As Boolean = False
            Dim multibadge As Boolean = False
            Dim faceDetected As Boolean = True
            Dim photoARefaire As Boolean = False
            Dim presentAD As Boolean = False
            Dim matricule As String = ""
            Dim nom As String = ""
            Dim dateSortieTxtAD As String = ""
            Dim endDateMSTxt As String = ""

            matricule = employeeid

            Dim fiche = jsonMS.MakeRequest("GET", "/users?prettyPrint&fields=id,lastName,firstName,valid,validityStartDate,validityEndDate,isCurrentlyValid,photo.root&filter=matricule=" & matricule)
            Dim ficheResponseData = New JavaScriptSerializer().Deserialize(Of Object)(fiche)
            If ficheResponseData("data").length = 0 Then
                presentMS = False
            End If



            If presentMS = True Then

                'If ficheResponseData("data")(0)("valid") = "true" Then


                '    Dim dateDebut As String = ficheResponseData("data")(0)("validityStartDate")
                '    Dim dateFin As String = ficheResponseData("data")(0)("validityEndDate")
                '    If dateDebut = "" Then dateDebut = "1900-01-01"

                '    If dateFin <> "" Then
                '        activeMS = IsValidMS(dateDebut, dateFin)
                '    Else
                '        activeMS = True
                '    End If
                'End If

                Dim dateFin As String = ficheResponseData("data")(0)("validityEndDate")
                If dateFin <> "" Then
                    Dim endDateMS As Date = Convert.ToDateTime(dateFin)
                    endDateMSTxt = endDateMS.ToString("dd/MM/yyyy")

                End If

                activeMS = Convert.ToBoolean(ficheResponseData("data")(0)("isCurrentlyValid"))
                Dim donneephoto = ficheResponseData("data")(0)("photo")
                photo = (Not donneephoto Is Nothing)

                If photo = True Then
                    faceDetected = ficheResponseData("data")(0)("photo")("facesDetected")
                    photoARefaire = Not faceDetected
                End If

            End If

            Dim filter1 As SearchResult = Commun.SearchFilterOne(RecupDataini.RecupVar("[OUUtilisateurs]"), "(&(employeeID=" & employeeid & "))", SearchScope.Subtree, "employeeID")
            If filter1 IsNot Nothing Then
                Using user As DirectoryEntry = New DirectoryEntry(filter1.Path)

                    presentAD = (user.Parent.Path = "LDAP://OU=Utilisateurs,DC=igbmc,DC=u-strasbg,DC=fr")

                    dateSortieTxtAD = user.Properties("extensionAttribute1").Value
                    Dim dateSortie As DateTime
                    If dateSortieTxtAD <> "" Then
                        dateSortie = DateTime.ParseExact(dateSortieTxtAD, "dd/MM/yyyy", Nothing)
                        If dateSortie < #1/1/2020 12:00:00 AM# Then Continue For
                    End If

                    nom = user.Properties("sn").Value & " " & user.Properties("givenName").Value

                    If presentMS = True Then
                        'Dim idMS As String = ficheResponseData("data")(0)("id") 'deviceCreate("id")
                        'Dim credbadge = jsonMS.MakeRequest("GET", "/users/" & id & "?fields=credentials.TECHNO_01.csn")
                        badgeEdit = user.Properties.Contains("serialNumber")
                        If badgeEdit = True Then
                            If user.Properties("serialNumber").Count > 1 Then
                                multibadge = True
                            End If
                        End If


                    End If

                End Using
            Else
                nom = ficheResponseData("data")(0)("lastName") & " " & ficheResponseData("data")(0)("firstName")

            End If


            'comptage
            If presentAD = True Then countAD += 1
            If presentMS = True Then countMS += 1
            If activeMS = True Then countActiveMS += 1
            If presentAD = True And activeMS = False Then countMSACreer += 1
            If presentAD = False And activeMS = True Then countMSASuppr += 1
            If presentAD = True And activeMS = True And photo = False Then countPhotoAFaire += 1
            If multibadge = True Then CountMultiBadge += 1
            If activeMS = True And presentAD = True And photo = True And badgeEdit = False Then countBadgeEdit += 1
            If presentAD = True And activeMS = True And photo = True And faceDetected = False Then
                countPhotoARefaire += 1
                countPhotoAFaire += 1
            End If
            File.AppendAllText("c:\temp\MSrapport.csv", vbCrLf & matricule & "," & nom & "," & presentAD & "," & presentMS & "," & activeMS & "," & dateSortieTxtAD & "," & badgeEdit & "," & photo & "," & photoARefaire & "," & multibadge & "," & endDateMSTxt)

        Next

        Dim compteurPresent = jsonMS.MakeRequest("GET", "/areas/2?fields=usersCounter")
        Dim compteurPresentResponseData = New JavaScriptSerializer().Deserialize(Of Object)(compteurPresent)
        Dim counter As Integer = compteurPresentResponseData("usersCounter")

        Dim corpMail As String = "Personnes Totales presentes dans MicroSesame : " & countMS & vbCrLf & vbCrLf _
            & "Personnes presentes dans l'AD : " & countAD & vbCrLf & vbCrLf _
            & "Personnes actives dans MicroSesame: " & countActiveMS & vbCrLf & vbCrLf _
            & "Personnes à supprimer dans MicroSesame : " & countMSASuppr & vbCrLf & vbCrLf _
            & "Personnes à creer dans MicroSesame (ou Désactivées): " & countMSACreer & vbCrLf & vbCrLf _
            & "Photo à faire dans MicroSesame : " & countPhotoAFaire & " (dont " & countPhotoARefaire & " à refaire)" & vbCrLf & vbCrLf _
            & "Badge à éditer dans MicroSesame : " & countBadgeEdit & vbCrLf & vbCrLf _
            & "Personnes avec plusieurs badges déclarés : " & CountMultiBadge & vbCrLf & vbCrLf & vbCrLf _
            & "Personnes présentes à " & Format(Now, "hh:mm") & " : " & counter

        File.Copy("c:\temp\MSrapport.csv", Form1.nomFichierRapportMS)
        Commun.SendEmail("administrateur@igbmc.fr", "officiersorienteurs@igbmc.fr", "Fichier de controle MicroSesame  (" & Now & ")", corpMail, Form1.nomFichierRapportMS) 'kolb@igbmc.fr;

    End Sub
    Shared Function IsValidMS(ByVal startDateTxt As String, ByVal endDateTxt As String) As Boolean
        Dim result As Boolean = False
        Dim startDate As Date = Convert.ToDateTime(startDateTxt)
        Dim endDate As Date = Convert.ToDateTime(endDateTxt)
        Dim aaa = Now
        If Now > startDate And Now < DateAdd("d", 1, endDate) Then
            result = True
        End If
        Return result
    End Function

    Shared Sub adddatedefin()
        Using AD As DirectoryEntry = New DirectoryEntry("LDAP://" & RecupDataini.RecupVar("[OUUtilisateursDesactives]"), Commun.admin, Commun.passwd, AuthenticationTypes.SecureSocketsLayer + AuthenticationTypes.Secure)
            Dim filtre As SearchResultCollection = Commun.SearchFilterAll(AD, "(&(objectCategory=person)(!(extensionAttribute1=*)))", SearchScope.Subtree)
            For Each result As SearchResult In filtre
                Using user As DirectoryEntry = New DirectoryEntry(result.Path)

                    Dim employeeID As String = user.Properties("employeeid").Value
                    Dim dataContract As String = Json.SendJson("", "persons/" & employeeID & "/contracts", "AD", "GET")
                    'Dim contractRunning As Boolean = Form1.ContratEnCours(dataContract)
                    'Dim dataContract As String = Json.SendJson("", "persons/" & IDuser & "/contracts?current_contract=true", "AD", "GET")

                    Dim contracts = Json.DeserializeJson(dataContract, "contracts")
                    Dim finContrat As String = Form1.DateDeFinDeContract(dataContract, employeeID)
                    Dim dateactu As String = user.Properties("extensionAttribute1").Value
                    If finContrat = "" Then

                        If dateactu <> finContrat Then
                            user.Properties("extensionAttribute1").Value = finContrat
                            user.CommitChanges()
                            Commun.Journal("Mise a jour d'une date de fin de contrat : " & employeeID & " : " & user.Properties("cn").Value & " : " & finContrat & " --> " & dateactu)
                        End If
                    Else
                        Dim findecontratDate As Date = DateTime.ParseExact(finContrat, "dd/MM/yyyy", Nothing)

                        If findecontratDate <> "01/01/0001" Then
                            If dateactu <> finContrat Then

                                user.Properties("extensionAttribute1").Value = finContrat
                                user.CommitChanges()
                                Commun.Journal("Mise a jour d'une date de fin de contrat : " & employeeID & " : " & user.Properties("cn").Value & " : " & finContrat & " --> " & dateactu)
                            End If
                        End If
                    End If
                End Using
            Next
        End Using
    End Sub
End Class
