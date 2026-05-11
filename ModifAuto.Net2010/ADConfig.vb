Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.Linq

Public Class ADHelper

    Public Shared dcs As String = ini.ReadValue("MODIFAUTO", "dcs")
    Public Shared Sub InitialiserDC()
        If Not String.IsNullOrWhiteSpace(Commun.DCName) Then
            Exit Sub
        End If

        Commun.DCName = ChoisirDcAuDemarrage()
        Commun.Journal("DC global choisi : " & Commun.DCName)
    End Sub

    Private Shared Function ChoisirDcAuDemarrage() As String

        If String.IsNullOrWhiteSpace(dcs) Then
            Throw New Exception("La variable globale 'dcs' est vide.")
        End If

        Dim dcList As String() = Split(dcs, ",")
        Dim mesures As New List(Of Tuple(Of String, Long))()

        For Each dc As String In dcList
            dc = Trim(dc)

            If dc = "" Then
                Continue For
            End If

            Dim sw As Stopwatch = Stopwatch.StartNew()

            Try
                Using root As New DirectoryEntry("LDAP://" & dc & "/rootDSE", Commun.admin, Commun.passwd, auth)
                    Dim namingContext As String = CStr(root.Properties("defaultNamingContext").Value)

                    If Not String.IsNullOrWhiteSpace(namingContext) Then
                        sw.Stop()
                        mesures.Add(Tuple.Create(dc, sw.ElapsedMilliseconds))
                    End If
                End Using
            Catch ex As Exception
                Debug.WriteLine("DC ignoré : " & dc & " - " & ex.Message)
            End Try
        Next

        If mesures.Count = 0 Then
            Throw New Exception("Aucun contrôleur de domaine disponible.")
        End If

        Return mesures.OrderBy(Function(x) x.Item2).First().Item1
    End Function

End Class


