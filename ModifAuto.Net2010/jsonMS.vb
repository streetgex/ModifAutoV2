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
    Shared serveurMS As String = "serv-ca.igbmc.u-strasbg.fr:437"
    Shared token As String = "{6ec09464-6b1f-4f5e-a39e-a7db30e15338}"
    Shared Function MakeRequest(method As String, requestAPI As String) As String
        Dim fullUrl As String = "https://" & serveurMS & "/api" & requestAPI
        Dim responseFromServer As String = ""
        Try
            ' ignore ssl certificate
            ServicePointManager.ServerCertificateValidationCallback = AddressOf clsSSL.AcceptAllCertifications
            'ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls

            Dim request As HttpWebRequest = DirectCast(WebRequest.Create(New Uri(fullUrl)), HttpWebRequest)

            request.Accept = "application/json"
            request.ContentType = "application/json"
            request.Method = method
            'request.UserAgent = userAgent

            ServicePointManager.UseNagleAlgorithm = False
            ServicePointManager.Expect100Continue = False


            'request.CookieContainer = savedCookies
            request.KeepAlive = False
            request.Headers.Add("X-API-KEY", token)
            request.Headers.Add("X-PROFILE", "1")



            Dim startTime As Date = Now
            Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)


            If response.StatusCode <> HttpStatusCode.OK Then
                Throw New Exception([String].Format("Server error (HTTP {0}: {1}).", response.StatusCode, response.StatusDescription))
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
                'ConsoleWrite("Response is empty : exit", ConsoleColor.Red)
                End
            End If
            ' Clean up the streams.  
            reader.Close()
            dataStream.Close()
            response.Close()

        Catch e As Exception
            'ConsoleWrite(e.ToString(), ConsoleColor.Red)
        End Try

        Return responseFromServer


    End Function
    Shared Function AllBadgeNumber(ByVal matricule As String) As String()

        Dim fiche = jsonMS.MakeRequest("GET", "/users?prettyPrint&fields=id&filter=matricule=" & matricule)
        Dim ficheResponseData = New JavaScriptSerializer().Deserialize(Of Object)(fiche)

        If ficheResponseData("data").length = 0 Then
            Exit Function
        End If
        Dim id As String = ficheResponseData("data")(0)("id") 'deviceCreate("id")
        Dim credbadge = jsonMS.MakeRequest("GET", "/users/" & id & "?fields=credentials.TECHNO_01.csn")
        Dim credbadgeResponseData = New JavaScriptSerializer().Deserialize(Of Object)(credbadge)
        Dim techno1 = credbadgeResponseData("credentials")("TECHNO_01")

        Dim i As Integer = -1
        For Each badge In techno1
            i += 1
            Dim numeroBadge As String = badge("csn")
            If numeroBadge <> "" Then
                AllBadgeNumber.Add(numeroBadge)
            End If
        Next badge
        Return AllBadgeNumber
    End Function
    Shared Function AllCred()
        Dim fiche = jsonMS.MakeRequest("GET", "/credentials?fields=holder.matricule,holder.id,csn,code&offset=500")
        Dim ficheResponseData = New JavaScriptSerializer().Deserialize(Of Object)(fiche)
        'Dim data = ficheResponseData("data")
        'Dim ficheResponseData = New JavaScriptSerializer().Deserialize(Of Object)(dataResponseData)
        Dim data = ficheResponseData("data")
        Dim max = UBound(Data)
        Dim id = ficheResponseData("data")(0)("holder") 'deviceCreate("id")
        Dim credbadge = jsonMS.MakeRequest("GET", "/users/" & id & "?fields=credentials.TECHNO_01.code")
        Dim credbadgeResponseData = New JavaScriptSerializer().Deserialize(Of Object)(credbadge)
        Dim techno1 = credbadgeResponseData("credentials")("TECHNO_01")

        Dim i As Integer = -1
        For Each badge In techno1
            i += 1
            Dim numeroBadge As String = badge("code")
            ' AllBadgeNumber.Add(numeroBadge)
        Next badge
        'Return AllBadgeNumber()
    End Function
End Class

Public Class clsSSL
    Public Shared Function AcceptAllCertifications(ByVal sender As Object, ByVal certification As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
        Return True
    End Function

End Class