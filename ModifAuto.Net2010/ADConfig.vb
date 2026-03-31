Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.Linq

Public Class ADHelper

    Public Shared Sub InitialiserDC()
        If Not String.IsNullOrWhiteSpace(Commun.DCName) Then
            Exit Sub
        End If

        Commun.DCName = ChoisirDcAuDemarrage("igbmc.u-strasbg.fr")
        Debug.WriteLine("DC global choisi : " & Commun.DCName)
    End Sub

    Private Shared Function ChoisirDcAuDemarrage(domainName As String) As String
        If String.IsNullOrWhiteSpace(domainName) Then
            Throw New ArgumentException("domainName vide.")
        End If

        Dim ctx As New DirectoryContext(DirectoryContextType.Domain, domainName)
        Dim domain As Domain = Domain.GetDomain(ctx)
        Dim mesures As New List(Of Tuple(Of String, Long))()

        For Each dc As DomainController In domain.FindAllDiscoverableDomainControllers()
            Dim sw As Stopwatch = Stopwatch.StartNew()

            Try
                Using root As New DirectoryEntry("LDAP://" & dc.Name & "/rootDSE")
                    Dim namingContext = CStr(root.Properties("defaultNamingContext").Value)
                    If Not String.IsNullOrWhiteSpace(namingContext) Then
                        sw.Stop()
                        mesures.Add(Tuple.Create(dc.Name, sw.ElapsedMilliseconds))
                    End If
                End Using
            Catch ex As Exception
                Debug.WriteLine("DC ignoré : " & dc.Name & " - " & ex.Message)
            End Try
        Next

        If mesures.Count = 0 Then
            Throw New Exception("Aucun contrôleur de domaine disponible.")
        End If

        Return mesures.OrderBy(Function(x) x.Item2).First().Item1
    End Function

End Class

Public Module ExempleADHelper

    Public Sub InitialisationApplication()
        ADHelper.InitialiserDC()
        Debug.WriteLine("Application connectée au DC : " & Commun.DCName)
    End Sub

End Module
