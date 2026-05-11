Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.DirectoryServices
Imports System.IO
Imports System.Web.Script.Serialization



Public Class Thumbn


    Shared fichierThumbTemp As String = "c:\temp\thumbTemp.jpg"

    'Public Shared srv As String = "Space2"
    'Public Shared pathPhoto As String = "\\" & Thumbn.srv & "\photos-AD\AD\"
    'Public Shared pathThumbs As String = "\\" & Thumbn.srv & "\photos-AD\AD\Thumb\"
    'Public Shared srv As String = "Labo4"
    'Public Shared pathPhoto As String = "\\" & Thumbn.srv & "\photos-RH\"
    'Public Shared pathThumbs As String = "\\" & Thumbn.srv & "\photos-RH\Thumbnails\"

    'Shared Sub ComparePhoto(ByVal cn As String, ByVal matricule As String, ByVal objDirEnt As DirectoryEntry)
    Shared Sub ComparePhoto(ByVal photo64 As String, ByVal autorisationDiffInterne As String, ByVal objDirEnt As DirectoryEntry)
        Try


            If autorisationDiffInterne = False Then
                If (Not objDirEnt.Properties("jpegPhoto").Value Is Nothing Or Not objDirEnt.Properties("thumbnailPhoto").Value Is Nothing) Then
                    objDirEnt.Properties("jpegPhoto").Value = Nothing
                    objDirEnt.Properties("thumbnailPhoto").Value = Nothing
                    Commun.AppliquerChangement(objDirEnt)
                    Commun.Journal("Nettoyage des attributs Photos Réussi : " & objDirEnt.Properties("sAMAccountName").Value)
                End If
                Exit Sub
            End If


            Dim imageOriginale As String = photo64
            imageOriginale = imageOriginale.Replace("data:image/jpeg;base64,", "")

            'Dim aaaa = imageOriginale.Length

            Dim testCompare As Boolean = False

            If Not objDirEnt.Properties("jpegPhoto").Value Is Nothing Then
                Dim imageAD As Byte() = DirectCast(objDirEnt.Properties("jpegPhoto")(0), Byte())
                'testCompare = CompareImageBytesIdentique(imageAD, imageOriginale) '
                testCompare = CompareImageBytes_base64Identique(imageAD, imageOriginale)
                imageAD = Nothing

            End If

            If testCompare = False Then

                objDirEnt.Properties("jpegPhoto").Value = Nothing
                objDirEnt.Properties("thumbnailPhoto").Value = Nothing
                Commun.AppliquerChangement(objDirEnt)

                Dim ImageOriginaleByte As Byte() = Convert.FromBase64String(imageOriginale)
                Dim imageByteTh As Byte() = CreateThumb1(Convert.FromBase64String(imageOriginale), objDirEnt.Properties("employeeID").Value)

                objDirEnt.Properties("jpegPhoto").Insert(0, ImageOriginaleByte)
                objDirEnt.Properties("thumbnailPhoto").Insert(0, imageByteTh)
                Commun.AppliquerChangement(objDirEnt)
                Commun.Journal("Modification des attributs Photos Réussi : " & objDirEnt.Properties("sAMAccountName").Value)

            End If
            testCompare = Nothing

        Catch ex As Exception
            Commun.Journal("ERREUR : Comparaison Thumbnail : " & Replace(objDirEnt.Path, "LDAP://" & Commun.LdapServerPrefix(), "") & " : " & ex.Message, True)
            Return
        End Try
    End Sub

    Shared Function CreateThumb1(ByVal imageBytes As Byte(), ByVal matricule As String) As Byte()

        Dim sFile As New System.IO.FileInfo(fichierThumbTemp)
        Dim testFichier As Boolean = sFile.Exists
        sFile = Nothing
        If testFichier Then
            Kill(fichierThumbTemp)
        End If

        Dim image As Image = Nothing
        Dim imgThumb As Image = Nothing
        Dim imageThBytes As Byte() = imageBytes

        Try


            Dim myImageCodecInfo As ImageCodecInfo
            Dim myEncoder As Encoder
            Dim myEncoderParameter As EncoderParameter
            Dim myEncoderParameters As EncoderParameters
            myImageCodecInfo = GetEncoderInfo("image/jpeg")
            myEncoder = Encoder.ColorDepth
            myEncoderParameters = New EncoderParameters(1)
            myEncoderParameter = New EncoderParameter(myEncoder, CType(24L, Int32))
            myEncoderParameters.Param(0) = myEncoderParameter

            Using MS As New MemoryStream(imageBytes, 0, imageBytes.Length)

                ' Convert byte[] to Image
                MS.Write(imageBytes, 0, imageBytes.Length)
                image = Image.FromStream(MS, True)
                image.Save(fichierThumbTemp, myImageCodecInfo, myEncoderParameters)
            End Using

            'Dim fichierPhoto As String = pathThumbs & matricule & ".jpg"
            Dim tailleFichierThumb As Long = Nothing

            image.Dispose()
            image = Nothing


            'File.WriteAllBytes(fichierThumbTemp, myImageCodecInfo, image)



            While imageThBytes.Length > 10240 'Or tailleFichierThumb = -1
                Using ms As New IO.MemoryStream(CType(imageThBytes, Byte()))
                    image = Image.FromStream(ms)
                End Using

                imgThumb = image.GetThumbnailImage(image.Width / 2, image.Height / 2, Nothing, New IntPtr())
                image.Dispose()
                image = Nothing
                Using ms = New MemoryStream()
                    imgThumb.Save(ms, myImageCodecInfo, myEncoderParameters)
                    'ecriture du fichier thumbnails
                    'imgThumb.Save(pathThumbs & "\" & matricule & ".jpg", myImageCodecInfo, myEncoderParameters)
                    imageThBytes = ms.ToArray
                End Using

                imgThumb = Nothing
                'tailleFichierThumb = GetFileLength(fichierThumbTemp)
            End While
            'File.Copy(fichierThumbTemp, pathThumbs & "\" & matricule & ".jpg")
            myEncoderParameter.Dispose()
            myEncoderParameter = Nothing
            myEncoderParameters.Dispose()
            myEncoderParameters = Nothing
            myImageCodecInfo = Nothing
            myEncoder = Nothing
            GC.Collect()
        Catch ex As Exception
            Commun.Journal("ERREUR : Création Thumbnail : " & matricule & " : " & ex.Message, True)
        End Try
        'If sFile.Exists Then
        '    Kill(fichierThumbTemp)
        'End If
        Return imageThBytes

    End Function

    Shared Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
        Dim j As Integer
        Dim encoders() As ImageCodecInfo
        encoders = ImageCodecInfo.GetImageEncoders()

        j = 0
        While j < encoders.Length
            If encoders(j).MimeType = mimeType Then
                Return encoders(j)
            End If
            j += 1
        End While
        encoders = Nothing
        Return Nothing

    End Function


    Shared Function CompareImageBytes_base64Identique(ByVal imageAD As Byte(), ByVal image64 As String) As Boolean
        Dim imageADbase64 As String = Convert.ToBase64String(imageAD, 0, imageAD.Length)
        If imageADbase64 <> image64 Then
            Return False
        Else
            Return True
        End If

        imageAD = Nothing
    End Function

End Class
