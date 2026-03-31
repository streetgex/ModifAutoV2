Imports System.DirectoryServices

Public Class UtilisateurRH
    Public Property nom As String
    Public Property prenom As String
    Public Property samAccountName As String
    Public Property destinationNomCourt As String
    Public Property destinationNomLong As String
    Public Property tempsTravail As String
    Public Property bureaux As String
    Public Property telephones As String
    Public Property equipeInfo As String
    Public Property IDNplus1 As String
    Public Property DNNplus1 As String
    Public Property ID As String
    Public Property aliasMailLong As String
    Public Property listesDiffusions As String
    Public Property organism As String
    Public Property finDeContrat As String
    Public Property genre As String
    Public Property diffusionPhoto As String
    Public Property unité As String
    Public Property destNomLong As String

    Public Function TrouverDN() As String
        'Pas besoin de mettre samAccountName en argument puisque l'appel se fait par UtilisateurMD.TrouverDN() et samAccountName est connu par l'instanciation de utilisateurMD
        Dim sam As String = Me.samAccountName
        Using Ldap As DirectoryEntry = New DirectoryEntry("LDAP://" & Commun.LdapPath(OUUtilisateurs), AdminScriptLogin, AdminScriptPassword, auth)

            Using searcher As DirectorySearcher = New DirectorySearcher(Ldap)
                searcher.Filter = "(&(SAMAccountName=" & sam & "))"
                Dim result As SearchResult = searcher.FindOne()
                If Not result Is Nothing Then
                    Dim pathCN As String = Replace(result.Path, "LDAP://" & Commun.LdapServerPrefix(), "")
                    TrouverDN = pathCN
                End If
                result = Nothing
            End Using

        End Using
        Return TrouverDN
    End Function

End Class

