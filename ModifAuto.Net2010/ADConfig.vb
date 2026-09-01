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
                Using entry As New DirectoryEntry(
            "LDAP://" & dc & "/DC=igbmc,DC=u-strasbg,DC=fr",
            Commun.admin,
            Commun.passwd,
            auth)

                    entry.RefreshCache(New String() {"distinguishedName"})

                    sw.Stop()
                    mesures.Add(Tuple.Create(dc, sw.ElapsedMilliseconds))
                End Using

            Catch ex As Exception
                Commun.Journal("ERREUR : DC ignore : " & dc & " : " & ex.Message, True)
                Debug.WriteLine("DC ignoré : " & dc & " - " & ex.Message)
            End Try
        Next

        If mesures.Count = 0 Then
            Throw New Exception("Aucun contrôleur de domaine disponible.")
        End If

        Return mesures.OrderBy(Function(x) x.Item2).First().Item1
    End Function

End Class


