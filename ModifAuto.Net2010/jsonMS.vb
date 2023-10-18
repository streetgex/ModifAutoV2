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
    Shared Function MakeRequest1(method As String, requestAPI As String) As String
        Dim fullUrl As String = "https://" & serveurMS & "/api" & requestAPI
        Dim responseFromServer As String = ""
        'Try
        ServicePointManager.ServerCertificateValidationCallback = AddressOf clsSSL.AcceptAllCertifications
        'We this to make an HTTP web request
        Dim request As Net.HttpWebRequest = Net.WebRequest.Create(fullUrl)
        request.KeepAlive = False
        request.Headers.Add("X-API-KEY", token)
        request.Headers.Add("X-PROFILE", "1")
        request.ServerCertificateValidationCallback = AddressOf clsSSL.AcceptAllCertifications
        'Make the web request and get the response
        Dim response As Net.WebResponse = request.GetResponse

        response.Close()
            Dim stream As System.IO.Stream = response.GetResponseStream

            'Prepare buffer for reading from stream
            Dim buffer As Byte() = New Byte(1000) {}

            'Data read from stream is gathered here
            Dim data As New List(Of Byte)

            'Start reading stream
            Dim bytesRead = stream.Read(buffer, 0, buffer.Length)

            Do Until bytesRead = 0
                For i = 0 To bytesRead - 1
                    data.Add(buffer(i))
                Next

                bytesRead = stream.Read(buffer, 0, buffer.Length)
            Loop


            'Gets the JSON data
            Debug.WriteLine(System.Text.Encoding.UTF8.GetString(data.ToArray))

            response.Close()
            stream.Close()


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