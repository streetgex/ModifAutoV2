Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Security.Principal

Public Class IdentitePartagePST
    Private Const LogonNewCredentials As Integer = 9
    Private Const LogonProviderWinNT50 As Integer = 3

    <DllImport("advapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function LogonUser(
        username As String,
        domain As String,
        password As String,
        logonType As Integer,
        logonProvider As Integer,
        ByRef token As IntPtr
    ) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function CloseHandle(handle As IntPtr) As Boolean
    End Function

    Public Shared Sub Executer(action As Action)
        Dim login As String = If(Commun.ini.ReadValue("MODIFAUTO", "exchangeLogin"), "").Trim()
        Dim password As String = Commun.ini.DecryptPassword(
            Commun.ini.ReadValue("MODIFAUTO", "exchangePassword_encrypted"))
        Dim domaine As String = "IGBMC"
        Dim token As IntPtr = IntPtr.Zero

        If String.IsNullOrWhiteSpace(login) OrElse String.IsNullOrWhiteSpace(password) Then
            Throw New Exception("Les identifiants MODIFAUTO/exchangeLogin ou exchangePassword_encrypted ne sont pas renseignes.")
        End If

        If login.Contains("\") Then
            Dim elements As String() = login.Split("\"c)
            domaine = elements(0)
            login = elements(1)
        End If

        If Not LogonUser(login, domaine, password, LogonNewCredentials, LogonProviderWinNT50, token) Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If

        Try
            Dim contexte As WindowsImpersonationContext = WindowsIdentity.Impersonate(token)
            Try
                action()
            Finally
                contexte.Undo()
            End Try
        Finally
            CloseHandle(token)
        End Try
    End Sub
End Class