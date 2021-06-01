Imports System
Imports System.Collections.Generic
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security

Public Class FTPP

    Public UserName As String = String.Empty
    Public Password As String = String.Empty
    Public KeepAlive As Boolean = False
    Public UseSSL As Boolean = False
    Public m_FTPSite As String = String.Empty

    Public Property FTPSite() As String
        Get
            Return m_FTPSite
        End Get
        Set(ByVal value As String)
            m_FTPSite = value
            If Not m_FTPSite.EndsWith("/") Then m_FTPSite += "/"
        End Set
    End Property

    Private m_CurDir As String = String.Empty
    Public Property CurrentDirectory() As String
        Get
            Return m_CurDir
        End Get
        Set(ByVal value As String)
            m_CurDir = value
            If Not m_CurDir.EndsWith("/") And m_CurDir = "" Then
                m_CurDir += "/"
                m_CurDir = m_CurDir.TrimStart("/".ToCharArray())
            End If
        End Set
    End Property

    Public Sub New()
    End Sub

    Public Sub New(ByVal sFTPSite As String, ByVal sUserName As String, ByVal sPassword As String)
        UserName = sUserName
        Password = sPassword
        FTPSite = sFTPSite
    End Sub

    Public Function ValidateServerCertificate(ByVal sender As Object, ByVal certificate As X509Certificate, ByVal chain As X509Chain, ByVal sslPolicyErrors As SslPolicyErrors) As Boolean
        If sslPolicyErrors = sslPolicyErrors.RemoteCertificateChainErrors Then
            Return False
        ElseIf sslPolicyErrors = sslPolicyErrors.RemoteCertificateNameMismatch Then
            Dim z As System.Security.Policy.Zone = System.Security.Policy.Zone.CreateFromUrl(CType(sender, HttpWebRequest).RequestUri.ToString)
            If z.SecurityZone = System.Security.SecurityZone.Intranet Or z.SecurityZone = System.Security.SecurityZone.MyComputer Then
                Return True
            End If
            Return False
        End If
        Return True
    End Function
    Public Function GetFileList(ByVal CurDirectory As String, ByVal StartsWith As String, ByVal EndsWith As String) As List(Of String)
        CurrentDirectory = CurDirectory
        Return GetFileList(StartsWith, EndsWith)
    End Function

    Public Function GetFileList(ByVal StartsWith As String, ByVal EndsWith As String) As List(Of String)
        Dim oFTP As FtpWebRequest = CType(FtpWebRequest.Create(FTPSite & CurrentDirectory), FtpWebRequest)
        oFTP.Credentials = New NetworkCredential(UserName, Password)
        oFTP.KeepAlive = KeepAlive
        oFTP.EnableSsl = UseSSL
        If UseSSL Then ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf ValidateServerCertificate)
        oFTP.Method = WebRequestMethods.Ftp.ListDirectory
        Dim response As FtpWebResponse = CType(oFTP.GetResponse, FtpWebResponse)
        Dim sr As StreamReader = New StreamReader(response.GetResponseStream)
        Dim str As String = sr.ReadLine
        Dim oList As New List(Of String)
        While str IsNot Nothing
            str = Replace(str, Replace(CurrentDirectory, "/", "") & "/", "")
            str = Replace(str, "/", "")
            If str.StartsWith(StartsWith) And str.EndsWith(EndsWith) Then
                oList.Add(str)
            End If
            str = sr.ReadLine
        End While

        sr.Close()
        response.Close()
        oFTP = Nothing
        Return oList
    End Function

    Public Function GetFile(ByVal Name As String, ByVal DestFile As String) As Boolean


        Dim oFTP As FtpWebRequest = CType(FtpWebRequest.Create(FTPSite & Name), FtpWebRequest)
        oFTP.Credentials = New NetworkCredential(UserName, Password)
        oFTP.Method = WebRequestMethods.Ftp.DownloadFile
        oFTP.KeepAlive = KeepAlive
        oFTP.EnableSsl = UseSSL
        If UseSSL Then ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf ValidateServerCertificate)
        oFTP.UseBinary = True
        Dim response As FtpWebResponse = CType(oFTP.GetResponse, FtpWebResponse)
        Dim responseStream As Stream = response.GetResponseStream
        Dim fs As New FileStream(DestFile, FileMode.Create)
        Dim buffer(2047) As Byte
        Dim read As Integer = 1
        While read <> 0
            read = responseStream.Read(buffer, 0, buffer.Length)
            fs.Write(buffer, 0, read)
        End While
        responseStream.Close()
        fs.Flush()
        fs.Close()
        responseStream.Close()
        response.Close()
        oFTP = Nothing
        Return True


    End Function




    Public Function UploadFile(ByVal fileLoc As String, ByVal fileServer As String) As Boolean
        Dim oFile As New FileInfo(fileLoc)
        Dim ftpRequest As FtpWebRequest
        Dim ftpResponse As FtpWebResponse
        Try
            ftpRequest = CType(FtpWebRequest.Create(FTPSite + fileServer), FtpWebRequest)
            ftpRequest.Method = WebRequestMethods.Ftp.UploadFile
            ftpRequest.Proxy = Nothing
            ftpRequest.UseBinary = True
            ftpRequest.Credentials = New NetworkCredential(UserName, Password)
            ftpRequest.KeepAlive = KeepAlive
            ftpRequest.EnableSsl = UseSSL
            If UseSSL Then ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf ValidateServerCertificate)
            Dim fileContents(oFile.Length) As Byte
            Using fr As FileStream = oFile.OpenRead
                fr.Read(fileContents, 0, Convert.ToInt32(oFile.Length))
            End Using
            Using writer As Stream = ftpRequest.GetRequestStream
                writer.Write(fileContents, 0, fileContents.Length)
            End Using
            ftpResponse = CType(ftpRequest.GetResponse, FtpWebResponse)
            ftpResponse.Close()
            ftpRequest = Nothing
            Return True
        Catch ex As WebException
            Return False
        End Try
    End Function

    Public Function DeleteFile(ByVal filServ As String) As Boolean
        Dim oFTP As FtpWebRequest
        oFTP = CType(FtpWebRequest.Create(FTPSite & filServ), FtpWebRequest)
        oFTP.Credentials = New NetworkCredential(UserName, Password)
        oFTP.Method = WebRequestMethods.Ftp.DeleteFile
        oFTP.KeepAlive = KeepAlive
        oFTP.EnableSsl = UseSSL
        If UseSSL Then ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf ValidateServerCertificate)
        oFTP.UseBinary = True
        Dim response As FtpWebResponse = CType(oFTP.GetResponse, FtpWebResponse)
        Dim oStat As FtpStatusCode = response.StatusCode
        response.Close()
        oFTP = Nothing
        Return True
    End Function
End Class

