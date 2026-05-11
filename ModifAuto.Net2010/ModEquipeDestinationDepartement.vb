Imports System
Imports System.Collections.Generic
Imports System.DirectoryServices

Public Class EquipeInfo
    Public Property id_equipeinfo As String
    Public Property nom_court_equipeinfo As String
    Public Property nom_long_equipeinfo As String
    Public Property dn_equipeinfo_eq_ad As String
End Class

Public Class DepartementInfo
    Public Property id_dept As String
    Public Property nom_court_dept As String
    Public Property nom_long_dept As String
    Public Property dn_dept_ad As String
End Class

Public Class DestinationInfo
    Public Property id_dest As String
    Public Property nom_court_dest As String
    Public Property nom_long_dest As String
    Public Property id_equipeinfo As String
    Public Property id_dept As String
    Public Property login_responsable_dest As String
    Public Property dn_dest_ad As String
End Class

Public Module ModEquipeDestinationDepartement

    Public DicoEquipesInfoRefRH As New Dictionary(Of String, EquipeInfo)(StringComparer.OrdinalIgnoreCase)
    Public DicoDepartementsRH As New Dictionary(Of String, DepartementInfo)(StringComparer.OrdinalIgnoreCase)
    Public DicoDestinationsRH As New Dictionary(Of String, DestinationInfo)(StringComparer.OrdinalIgnoreCase)
    Public DicoGroupesDiffusionRH As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

    Public Sub ViderDictionnaires()
        DicoEquipesInfoRefRH.Clear()
        DicoDepartementsRH.Clear()
        DicoDestinationsRH.Clear()
        DicoGroupesDiffusionRH.Clear()
    End Sub


    Public Function ObtenirIdEquipeInfoDepuisNomCourt(nomCourtEquipeInfo As String) As String
        If String.IsNullOrWhiteSpace(nomCourtEquipeInfo) Then
            Return ""
        End If

        For Each item As EquipeInfo In DicoEquipesInfoRefRH.Values
            If String.Equals(item.nom_court_equipeinfo, nomCourtEquipeInfo, StringComparison.OrdinalIgnoreCase) Then
                Return item.id_equipeinfo
            End If
        Next

        Return ""
    End Function


    Public Function ChargerDnDepartementsDepuisAd() As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)

        Using ouDepart As New DirectoryEntry("LDAP://" & Commun.LdapPath("OU=Departements,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr"), Commun.admin, Commun.passwd, auth)
            Using searcher As New DirectorySearcher(ouDepart)
                searcher.Filter = "(&(objectClass=group)(languageCode=*))"
                searcher.SearchScope = SearchScope.OneLevel
                searcher.PropertiesToLoad.Clear()
                searcher.PropertiesToLoad.Add("languageCode")
                searcher.PropertiesToLoad.Add("distinguishedName")

                Dim results As SearchResultCollection = searcher.FindAll()

                For Each sr As SearchResult In results
                    If sr.Properties.Contains("languageCode") AndAlso sr.Properties.Contains("distinguishedName") Then
                        Dim idDept As String = sr.Properties("languageCode")(0).ToString().Trim()
                        Dim dnDept As String = sr.Properties("distinguishedName")(0).ToString()

                        If idDept <> "" AndAlso Not result.ContainsKey(idDept) Then
                            result.Add(idDept, dnDept)
                        End If
                    End If
                Next
            End Using
        End Using

        Return result
    End Function

    Public Sub ChargerInfosEquipeDepuisDestinationNomCourt(destinationNomCourt As String,
                                                            ByRef equipeInfoNomCourt As String,
                                                            ByRef equipeInfoNomLong As String,
                                                            ByRef equipeInfoID As String,
                                                            ByRef dnDestAd As String,
                                                            ByRef dnEquipeInfoEqAd As String)
        equipeInfoNomCourt = ""
        equipeInfoNomLong = ""
        equipeInfoID = ""
        dnDestAd = ""
        dnEquipeInfoEqAd = ""

        If String.IsNullOrWhiteSpace(destinationNomCourt) Then
            Exit Sub
        End If

        Dim racineEquipes As String = Commun.LdapPrefix & "OU=Equipes,DC=igbmc,DC=u-strasbg,DC=fr"
        Dim racineEquipesDest As String = Commun.LdapPrefix & "OU=Gestion_Destinations,OU=Equipes,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr"
        Dim racineEquipesInfos As String = Commun.LdapPrefix & "OU=Equipes,OU=EMC Celerra,DC=igbmc,DC=u-strasbg,DC=fr"

        Using entreeEquipes As New DirectoryEntry(racineEquipes)
            Using rechercheurDestination As New DirectorySearcher(entreeEquipes)
                rechercheurDestination.Filter = "(&(objectClass=group)(cn=" & destinationNomCourt & " grp))"
                rechercheurDestination.SearchScope = SearchScope.OneLevel
                rechercheurDestination.PropertiesToLoad.Add("memberOf")
                rechercheurDestination.PropertiesToLoad.Add("distinguishedName")

                Dim resultatDestination As SearchResult = rechercheurDestination.FindOne()
                If resultatDestination Is Nothing Then
                    Exit Sub
                End If

                dnDestAd = LirePropriete(resultatDestination, "distinguishedName")
                Dim dnEquipeDest As String = TrouverDnEquipeDestDepuisMemberOf(resultatDestination)
                If String.IsNullOrWhiteSpace(dnEquipeDest) Then
                    Exit Sub
                End If

                Using entreeEquipesDest As New DirectoryEntry(racineEquipesDest)
                    Using rechercheurEquipeDest As New DirectorySearcher(entreeEquipesDest)
                        rechercheurEquipeDest.Filter = "(&(objectClass=group)(distinguishedName=" & dnEquipeDest & "))"
                        rechercheurEquipeDest.SearchScope = SearchScope.Subtree
                        rechercheurEquipeDest.PropertiesToLoad.Add("cn")

                        Dim resultatEquipeDest As SearchResult = rechercheurEquipeDest.FindOne()
                        If resultatEquipeDest Is Nothing Then
                            Exit Sub
                        End If

                        equipeInfoNomCourt = LirePropriete(resultatEquipeDest, "cn")
                        If equipeInfoNomCourt.EndsWith("_dest", StringComparison.OrdinalIgnoreCase) Then
                            equipeInfoNomCourt = equipeInfoNomCourt.Substring(0, equipeInfoNomCourt.Length - 5)
                        End If
                    End Using
                End Using

                If String.IsNullOrWhiteSpace(equipeInfoNomCourt) Then
                    Exit Sub
                End If

                Using entreeEquipesInfos As New DirectoryEntry(racineEquipesInfos)
                    Using rechercheurEquipeInfo As New DirectorySearcher(entreeEquipesInfos)
                        rechercheurEquipeInfo.Filter = "(&(objectClass=group)(cn=" & equipeInfoNomCourt & "_eq))"
                        rechercheurEquipeInfo.SearchScope = SearchScope.OneLevel
                        rechercheurEquipeInfo.PropertiesToLoad.Add("displayName")
                        rechercheurEquipeInfo.PropertiesToLoad.Add("gidNumber")
                        rechercheurEquipeInfo.PropertiesToLoad.Add("distinguishedName")

                        Dim resultatEquipeInfo As SearchResult = rechercheurEquipeInfo.FindOne()
                        If resultatEquipeInfo Is Nothing Then
                            Exit Sub
                        End If

                        equipeInfoNomLong = LirePropriete(resultatEquipeInfo, "displayName")
                        equipeInfoID = LirePropriete(resultatEquipeInfo, "gidNumber")
                        dnEquipeInfoEqAd = LirePropriete(resultatEquipeInfo, "distinguishedName")
                    End Using
                End Using
            End Using
        End Using
    End Sub

    Private Function TrouverDnEquipeDestDepuisMemberOf(resultatDestination As SearchResult) As String
        If resultatDestination Is Nothing OrElse Not resultatDestination.Properties.Contains("memberOf") Then
            Return ""
        End If

        For Each groupeParent As Object In resultatDestination.Properties("memberOf")
            Dim dnParent As String = groupeParent.ToString()
            Dim cnParent As String = ExtraireCnDepuisDn(dnParent)

            If cnParent.EndsWith("_dest", StringComparison.OrdinalIgnoreCase) Then
                Return dnParent
            End If
        Next

        Return ""
    End Function

    Private Function LirePropriete(resultat As SearchResult, nomPropriete As String) As String
        If resultat IsNot Nothing AndAlso resultat.Properties.Contains(nomPropriete) AndAlso resultat.Properties(nomPropriete).Count > 0 Then
            Return resultat.Properties(nomPropriete)(0).ToString()
        End If

        Return ""
    End Function

    Private Function ExtraireCnDepuisDn(distinguishedName As String) As String
        If String.IsNullOrWhiteSpace(distinguishedName) Then
            Return ""
        End If

        If Not distinguishedName.StartsWith("CN=", StringComparison.OrdinalIgnoreCase) Then
            Return ""
        End If

        Dim resultat As New System.Text.StringBuilder()
        Dim i As Integer = 3
        Dim caractereEchappe As Boolean = False

        While i < distinguishedName.Length
            Dim caractereCourant As Char = distinguishedName(i)

            If caractereEchappe Then
                resultat.Append(caractereCourant)
                caractereEchappe = False
            ElseIf caractereCourant = "\"c Then
                caractereEchappe = True
            ElseIf caractereCourant = ","c Then
                Exit While
            Else
                resultat.Append(caractereCourant)
            End If

            i += 1
        End While

        Return resultat.ToString()
    End Function
    Public Class ReferentielDestinationsDepartements
        Public Property Destinations As Dictionary(Of String, DestinationInfo)
        Public Property Departements As Dictionary(Of String, DepartementInfo)
        Public Property EquipesInfo As Dictionary(Of String, EquipeInfo)
    End Class
End Module
