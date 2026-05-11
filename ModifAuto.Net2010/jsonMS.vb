Imports System.IO
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Text
Imports System.Security.Cryptography.X509Certificates
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


Public Class jsonMS
    'Shared serveurMS As String = "serv-ca-formation.igbmc.u-strasbg.fr"
    'Shared token As String = "{bddb6c80-98c3-4a38-9ee7-fffddfd04343}"
    Shared serveurMS As String = ini.ReadValue("MS", "serveur") '"serv-ca.igbmc.u-strasbg.fr:443"
    Shared token As String = ini.ReadValue("MS", "token_encrypted") '"{6ec09464-6b1f-4f5e-a39e-a7db30e15338}"
    Shared Function MakeRequest(method As String, requestAPI As String, Optional ByVal data As String = "") As String
        Dim fullUrl As String = "https://" & serveurMS & "/api" & requestAPI
        'Dim fullUrl As String = "https://serv-ca.igbmc.u-strasbg.fr:81/api" & requestAPI
        Dim responseFromServer As String = ""
        Try

            ' ignore ssl certificate
            ServicePointManager.ServerCertificateValidationCallback = AddressOf clsSSL.AcceptAllCertifications
            'ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
            'ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            'Dim request As HttpWebRequest = DirectCast(WebRequest.Create(New Uri(fullUrl)), HttpWebRequest)
            Dim request As HttpWebRequest = HttpWebRequest.Create(fullUrl)

            request.Accept = "application/json"
            request.Method = method

            request.KeepAlive = True
            request.Headers.Add("X-API-KEY", token)
            request.Headers.Add("X-PROFILE", "1")

            If method = "PUT" Or method = "POST" Then
                request.ContentType = "application/json; charset=utf-8"
                Dim jsonDataBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(data)
                request.ContentLength = jsonDataBytes.Length
                Dim stream = request.GetRequestStream()
                stream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
                stream.Close()
                stream.Dispose()
            ElseIf method = "GET" Or method = "DELETE" Then
                request.ContentType = "application/json"
            Else
                Exit Function
            End If

            ServicePointManager.UseNagleAlgorithm = False
            ServicePointManager.Expect100Continue = False

            Dim startTime As Date = Now
            Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)


            If response.StatusCode <> HttpStatusCode.OK Then
                Throw New Exception([String].Format("Server error (HTTP {0}: {1}).", response.StatusCode, response.StatusDescription))
                Commun.Journal("ERREUR : Connexion à MS : " & response.StatusCode & " " & response.StatusDescription, True)
                sendJournalError()
                End
            End If


            ' Display the status.
            Dim responseStatus As String = DirectCast(response, HttpWebResponse).StatusDescription
            If responseStatus = "OK" Then
                'ConsoleWrite("Connection to server : " & responseStatus, ConsoleColor.Green)
            Else
                'ConsoleWrite("Connexion to server failed : " & responseStatus & " : exit", ConsoleColor.Red)
            End If

            ' Get the stream containing content returned by the server.  
            Dim dataStream As Stream = response.GetResponseStream()
            ' Open the stream using a StreamReader for easy access.  
            Dim reader As New StreamReader(dataStream)
            ' Read the content.  
            responseFromServer = reader.ReadToEnd()
            ' Display the content. 
            Dim endTime As Date = Now
            If responseFromServer <> "" Then
                'ConsoleWrite("Response from server in " & DateDiff(DateInterval.Second, startTime, endTime) & " seconds : OK", ConsoleColor.Green)
            Else
                Commun.Journal("Response is empty : exit", True)
                sendJournalError()
                End
            End If
            ' Clean up the streams.  
            reader.Close()
            dataStream.Close()
            response.Close()

        Catch e As Exception
            'MsgBox(e.Message)
            'ConsoleWrite(e.ToString(), ConsoleColor.Red)
        End Try

        Return responseFromServer


    End Function

    Shared Function ChargerBadgesParEmployeeId() As Dictionary(Of String, String())
        Dim result As New Dictionary(Of String, String())(StringComparer.OrdinalIgnoreCase)

        Dim offset As Integer = 0
        Dim limit As Integer = 100

        Do
            Dim url As String =
            "/credentials/?filter=idTechno=1,csn,holder.matricule,status=k_valid" &
            "&limit=" & limit &
            "&offset=" & offset &
            "&fields=holder.matricule,status,csn,holder.mail"

            Dim fiche As String = jsonMS.MakeRequest("GET", url)
            Dim responseData = New JavaScriptSerializer().Deserialize(Of Object)(fiche)

            If responseData Is Nothing OrElse Not responseData.ContainsKey("data") Then
                Exit Do
            End If

            Dim data = responseData("data")
            If data Is Nothing OrElse data.length = 0 Then
                Exit Do
            End If

            For Each cred In data
                Dim matricule As String = ""
                Dim csn As String = ""

                If Not cred("csn") Is Nothing Then
                    csn = CStr(cred("csn"))
                End If

                If cred.ContainsKey("holder") AndAlso Not cred("holder") Is Nothing Then
                    Dim holder = cred("holder")
                    If holder.ContainsKey("matricule") AndAlso Not holder("matricule") Is Nothing Then
                        matricule = CStr(holder("matricule"))
                    End If
                End If

                If matricule = "" OrElse csn = "" Then
                    Continue For
                End If

                If Not result.ContainsKey(matricule) Then
                    result.Add(matricule, New String() {})
                End If

                result(matricule).Add(csn)
            Next

            If data.length < limit Then
                Exit Do
            End If

            offset += limit
        Loop

        For Each matricule As String In result.Keys.ToList()
            result(matricule) = TrierTableau(result(matricule))
        Next

        Return result
    End Function
    Shared Function GetIdMS(ByVal matricule As Integer) As String
        'recuperation de l'id MS
        Dim reqUser As String = "/users?fields=id&filter=matricule=" & matricule
        Dim dataUser As String = jsonMS.MakeRequest("GET", reqUser)
        Dim responseUser = New JavaScriptSerializer().Deserialize(Of Object)(dataUser)
        Dim idMS As Integer = responseUser("data")(0)("id")
        Return idMS
    End Function
    Shared Sub SetMSEndValidity(ByVal idMS As String, ByVal finDeContrat As String)
        If finDeContrat = "Aucune" Then finDeContrat = "01/01/2050"



        Dim dateObj As DateTime

        DateTime.TryParseExact(finDeContrat, "dd/MM/yyyy", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, dateObj)

        Dim dateFormatee As String = dateObj.ToString("yyyy-MM-dd")

        Dim data As String = "{""validityEndDate"":  """ & dateFormatee & """}"
        Dim dataresponse As String = jsonMS.MakeRequest("PUT", "/users/" & idMS, data)
    End Sub
    Shared Function GetMSAccreditation(ByVal idMS As String, ByVal finDeContrat As String) As String
        GetMSAccreditation = ""
        'https://serv-ca.igbmc.u-strasbg.fr/api/userClearances?fields=all&filter=user.id=7724
        If finDeContrat = "Aucune" Then finDeContrat = "01/01/2050"
        ' Construction de la requête pour récupérer les utilisateurs présents dans la zone

        Dim dateFindeContrat As DateTime

        DateTime.TryParseExact(finDeContrat, "dd/MM/yyyy", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, dateFindeContrat)

        Dim dateFormatee As String = dateFindeContrat.ToString("yyyy-MM-dd")

        '$"Contrôleur de domaine sélectionné : {SelectedDomainController}"
        'Dim reqClearance As String = "/userClearances?fields=clearance[name,id],endDate&filter=user.matricule=" & usrID & ",endDate<=" & dateFormatee
        Dim reqAccreditation As String = $"/users/{idMS}/accreditations?fields=accreditation.name,endDate&filter=endDate<{dateFormatee}&limit=100"
        Dim dataAccreditation As String = jsonMS.MakeRequest("GET", reqAccreditation)

        ' Vérification si la réponse est nulle (échec de la requête)
        If dataAccreditation Is Nothing Or dataAccreditation = "" Then Exit Function


        ' Désérialisation de la réponse JSON en un objet
        Dim responseClearance = New JavaScriptSerializer().Deserialize(Of Object)(dataAccreditation)

        ' Boucle sur les utilisateurs présents dans la zone
        For Each accreditation In responseClearance("data")
            Dim accreditationName As String = accreditation("accreditation")("name")
            Dim endDateTxt As String = accreditation("endDate")
            Dim endDate As Date
            DateTime.TryParseExact(endDateTxt, "yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, endDate)

            GetMSAccreditation += accreditationName & " (" & endDate.ToString("dd/MM/yyyy") & "), "
        Next
        If GetMSAccreditation <> "" Then
            GetMSAccreditation = Left(GetMSAccreditation, GetMSAccreditation.Length - 2)
        End If
        Return GetMSAccreditation
    End Function
End Class

Public Class clsSSL
    Public Shared Function AcceptAllCertifications(ByVal sender As Object, ByVal certification As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
        Return True
    End Function

End Class