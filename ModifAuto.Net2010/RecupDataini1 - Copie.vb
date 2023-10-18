Imports System.IO
Imports System.Security.Cryptography



Public Class ini
    Private ReadOnly _filePath As String
    Shared keyBase64 As String = "GR7x6ODFJnOQrDiqoGE3xY5eRcJZWZJP4ptD9EoTMUQ="
    Shared ivBase64 As String = "g0PMLlmunjuAV1a9sCVD6A=="
    Shared valuesDictionary As New Dictionary(Of String, Dictionary(Of String, String))

    Public Shared tabVar(1, 0) As String

    Shared Sub recupDataConnect()

        'Using sr As New StreamReader("\\igbmc.u-strasbg.fr\SYSVOL\igbmc.u-strasbg.fr\Scripts\ScriptSteph1.ini")
        '    Dim currentSection As String = ""

        '    While Not sr.EndOfStream
        '        Dim line As String = sr.ReadLine().Trim()

        '        If line.StartsWith("[") AndAlso line.EndsWith("]") Then
        '            currentSection = line.Substring(1, line.Length - 2).Trim()
        '            If currentSection = section Then
        '                foundSection = True
        '            Else
        '                foundSection = False
        '            End If
        '            Continue While
        '        End If

        '        If Not String.IsNullOrEmpty(currentSection) Then
        '            Dim parts As String() = line.Split(New Char() {"="c}, 2)
        '            If parts.Length = 2 Then
        '                Dim key = parts(0).Trim()
        '                Dim value = parts(1).Trim()

        '                If Not valuesDictionary(currentSection).ContainsKey(key) Then
        '                    valuesDictionary(currentSection)(key) = value
        '                End If
        '            End If
        '        End If
        '    End While
        'End Using
    End Sub














    Shared Function RecupVar(section As String, key As String, Optional defaultValue As String = "") As String
        Dim value As String = defaultValue


        If valuesDictionary.ContainsKey(section) AndAlso valuesDictionary(section).ContainsKey(key) Then
            value = valuesDictionary(section)(key)
        Else


            Try
                Using sr As New StreamReader(Form1.iniFilePath)
                    Dim currentSection As String = ""
                    Dim foundSection As Boolean = False

                    While Not sr.EndOfStream
                        Dim line As String = sr.ReadLine().Trim()

                        ' Ignore les lignes de commentaire commençant par ';' ou '#'
                        If line.StartsWith(";") OrElse line.StartsWith("#") Then
                            Continue While
                        End If

                        ' Vérifie si la ligne contient une section [section]
                        If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                            currentSection = line.Substring(1, line.Length - 2).Trim()
                            If currentSection = section Then
                                foundSection = True

                            Else
                                foundSection = False
                            End If
                            Continue While
                        End If

                        ' Si nous sommes dans la section appropriée, vérifiez si la ligne contient la clé recherchée
                        If foundSection Then
                            Dim parts As String() = line.Split(New Char() {"="c}, 2)
                            If parts.Length = 2 AndAlso parts(0).Trim() = key Then
                                ' La ligne contient la clé recherchée
                                ' Examinez parts(1) pour voir s'il y a un commentaire en fin de ligne et supprimez-le
                                Dim valueWithComment As String = parts(1).Trim()
                                Dim commentIndex As Integer = valueWithComment.LastIndexOf("#"c) ' Recherche le dernier symbole de commentaire
                                If commentIndex <> -1 Then
                                    ' Un commentaire a été trouvé, supprimez uniquement ce qui est après le dernier '#'
                                    valueWithComment = valueWithComment.Substring(0, commentIndex).Trim()
                                End If
                                value = valueWithComment ' La valeur sans commentaire en fin de ligne
                                Exit While
                            End If
                        End If
                    End While
                End Using
            Catch ex As Exception
                ' Gestion des erreurs de lecture du fichier INI
                Console.WriteLine("Erreur de lecture du fichier INI : " & ex.Message)
            End Try
            If Right(key, 10) = "_encrypted" Then
                value = DecryptPassword(value)
            Else
            End If

            If Not valuesDictionary.ContainsKey(section) Then
                valuesDictionary(section) = New Dictionary(Of String, String)
            End If


            valuesDictionary(section)(key) = value
        End If

        If value = "" Then
            Commun.Journal("ERREUR : RecupVar : Cette donnée n'existe pas : {" & section & ":" & key & "} : ", True)
            Form1.sendJournalError()
            End
        End If
        Return value

    End Function

    Shared Function DecryptPassword(key_encrypt As String) As String
        Try
            Dim aesAlg As New AesCryptoServiceProvider()

            ' Lisez les clés de chiffrement du fichier INI
            'key = Replace(key, "_encrypted", "")
            Dim encryptedPasswordBase64 As String = key_encrypt

            If String.IsNullOrWhiteSpace(keyBase64) OrElse String.IsNullOrWhiteSpace(ivBase64) OrElse String.IsNullOrWhiteSpace(encryptedPasswordBase64) Then
                Return ""
            End If

            ' Convertissez les clés à partir de Base64
            aesAlg.Key = Convert.FromBase64String(keyBase64)
            aesAlg.IV = Convert.FromBase64String(ivBase64)

            ' Convertissez le mot de passe chiffré à partir de Base64
            Dim encryptedPasswordBytes As Byte() = Convert.FromBase64String(encryptedPasswordBase64)

            ' Créez un déchiffreur AES
            Using decryptor As ICryptoTransform = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV)
                ' Déchiffrez le mot de passe
                Dim decryptedPasswordBytes As Byte() = decryptor.TransformFinalBlock(encryptedPasswordBytes, 0, encryptedPasswordBytes.Length)
                Return System.Text.Encoding.UTF8.GetString(decryptedPasswordBytes)
            End Using
        Catch ex As Exception
            ' Gestion des erreurs de déchiffrement ou de lecture dans le fichier
            Console.WriteLine("Erreur lors de la lecture du mot de passe : " & ex.Message)
            Return ""
        End Try
    End Function

    Shared Sub ConsoleWrite(ByVal texte As String, ByVal couleur As ConsoleColor)
        'voir les couleurs ici : https://msdn.microsoft.com/fr-fr/library/system.consolecolor(v=vs.110).aspx
        Console.ForegroundColor = couleur
        Console.WriteLine(texte)
        Console.ResetColor()
    End Sub

End Class
