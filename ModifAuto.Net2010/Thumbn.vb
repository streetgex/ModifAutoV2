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
                    Commun.Journal("Nettoyage des attributs Photos Réussi : " & objDirEnt.Properties("SAMAccountName").Value)
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
                Dim imageByteTh As Byte() = CreateThumb1(Convert.FromBase64String(imageOriginale), objDirEnt.Properties("EmployeeID").Value)

                objDirEnt.Properties("jpegPhoto").Insert(0, ImageOriginaleByte)
                objDirEnt.Properties("thumbnailPhoto").Insert(0, imageByteTh)
                Commun.AppliquerChangement(objDirEnt)
                Commun.Journal("Modification des attributs Photos Réussi : " & objDirEnt.Properties("SAMAccountName").Value)

            End If
            testCompare = Nothing

        Catch ex As Exception
            Commun.Journal("ERREUR : Comparaison Thumbnail : " & Replace(objDirEnt.Path, "LDAP://", "") & " : " & ex.Message, True)
            Return
        End Try
    End Sub
    Function ConvertBase64ToByte(ByVal base64String As String) As Byte()
        Dim result As Byte() = Convert.FromBase64String(base64String)
        Return result

    End Function
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

            'If tailleFichierThumb = -1 Then
            '    'File.Copy(fichierPhoto, fichierThumbTemp)

            '    image = Image.FromFile(fichierPhoto)
            '    Dim original As Bitmap = image
            '    image = original.Clone
            '    original.Dispose()
            '    original = Nothing
            '    image.Save(fichierThumbTemp, myImageCodecInfo, myEncoderParameters)
            'End If

            ''DECOUPAGE PAYSAGE

            'If image.Width > image.Height Then
            '    image = Image.FromFile(fichierThumbTemp)
            '    Dim focusRectangle As New Rectangle()
            '    Dim original As Bitmap = image
            '    focusRectangle.X = (original.Width - original.Height * 0.9) / 2
            '    focusRectangle.Y = 0
            '    focusRectangle.Height = original.Height - 1
            '    focusRectangle.Width = (original.Height * 0.9)
            '    image = original.Clone(focusRectangle, PixelFormat.DontCare)
            '    original.Dispose()
            '    original = Nothing
            '    image.Save(fichierThumbTemp, myImageCodecInfo, myEncoderParameters)
            'End If
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
    'Shared Sub CreateThumb(ByVal matricule As String)

    '    Dim sFile As New System.IO.FileInfo(fichierThumbTemp)
    '    Dim testFichier As Boolean = sFile.Exists
    '    sFile = Nothing
    '    If testFichier Then
    '        Kill(fichierThumbTemp)
    '    End If

    '    Try

    '        Dim image As Image = Nothing
    '        Dim imgThumb As Image = Nothing

    '        Dim myImageCodecInfo As ImageCodecInfo
    '        Dim myEncoder As Encoder
    '        Dim myEncoderParameter As EncoderParameter
    '        Dim myEncoderParameters As EncoderParameters
    '        myImageCodecInfo = GetEncoderInfo("image/jpeg")
    '        myEncoder = Encoder.ColorDepth
    '        myEncoderParameters = New EncoderParameters(1)
    '        myEncoderParameter = New EncoderParameter(myEncoder, CType(24L, Int32))
    '        myEncoderParameters.Param(0) = myEncoderParameter


    '        Dim fichierPhoto As String = pathThumbs & matricule & ".jpg"
    '        Dim tailleFichierThumb As Long = GetFileLength(fichierThumbTemp)


    '        If tailleFichierThumb = -1 Then
    '            'File.Copy(fichierPhoto, fichierThumbTemp)

    '            image = Image.FromFile(fichierPhoto)
    '            Dim original As Bitmap = image
    '            image = original.Clone
    '            original.Dispose()
    '            original = Nothing
    '            image.Save(fichierThumbTemp, myImageCodecInfo, myEncoderParameters)
    '        End If

    '        'DECOUPAGE PAYSAGE

    '        If image.Width > image.Height Then
    '            image = Image.FromFile(fichierThumbTemp)
    '            Dim focusRectangle As New Rectangle()
    '            Dim original As Bitmap = image
    '            focusRectangle.X = (original.Width - original.Height * 0.9) / 2
    '            focusRectangle.Y = 0
    '            focusRectangle.Height = original.Height - 1
    '            focusRectangle.Width = (original.Height * 0.9)
    '            image = original.Clone(focusRectangle, PixelFormat.DontCare)
    '            original.Dispose()
    '            original = Nothing
    '            image.Save(fichierThumbTemp, myImageCodecInfo, myEncoderParameters)
    '        End If
    '        image.Dispose()
    '        image = Nothing


    '        tailleFichierThumb = GetFileLength(fichierThumbTemp)

    '        While tailleFichierThumb > 10240 'Or tailleFichierThumb = -1

    '            image = Image.FromFile(fichierThumbTemp)
    '            imgThumb = image.GetThumbnailImage(image.Width / 2, image.Height / 2, Nothing, New IntPtr())
    '            image.Dispose()
    '            image = Nothing
    '            imgThumb.Save(fichierThumbTemp, myImageCodecInfo, myEncoderParameters)
    '            'ecriture du fichier thumbnails
    '            'imgThumb.Save(pathThumbs & "\" & matricule & ".jpg", myImageCodecInfo, myEncoderParameters)
    '            imgThumb.Dispose()
    '            imgThumb = Nothing
    '            tailleFichierThumb = GetFileLength(fichierThumbTemp)
    '        End While
    '        myEncoderParameter.Dispose()
    '        myEncoderParameter = Nothing
    '        myEncoderParameters.Dispose()
    '        myEncoderParameters = Nothing
    '        myImageCodecInfo = Nothing
    '        myEncoder = Nothing
    '        GC.Collect()
    '    Catch e As Exception
    '        Commun.Journal("ERREUR : Création Thumbnail : " & matricule, True)
    '        Return
    '    End Try
    '    'If sFile.Exists Then
    '    '    Kill(fichierThumbTemp)
    '    'End If
    'End Sub
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
    Shared Function GetFileLength(ByVal sPathFile As String) As Long
        Dim sFile As New System.IO.FileInfo(sPathFile)
        Dim lRet As Long
        If sFile.Exists Then lRet = sFile.Length Else lRet = -1
        sFile = Nothing
        Return lRet
    End Function
    Shared Function CompareImageBytesIdentique(ByVal imageAD As Byte(), ByVal imageF As Byte()) As Boolean
        If UBound(imageF) = UBound(imageAD) Then
            For i = 0 To UBound(imageF) - 1
                If imageF(i) <> imageAD(i) Then
                    Return False
                    imageF = Nothing
                    imageAD = Nothing
                    Exit Function
                End If
            Next i
        Else
            Return False
            imageF = Nothing
            imageAD = Nothing
            Exit Function
        End If
        Return True
        imageF = Nothing
        imageAD = Nothing
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
    Shared Function LireImage(ByVal imageFile As String) As Byte()
        Try
            Dim fi As FileInfo = New FileInfo(imageFile)
            fi.IsReadOnly = False
            Dim fs As New FileStream(imageFile, FileMode.Open)
            Dim r As New BinaryReader(fs)
            r.BaseStream.Seek(0, SeekOrigin.Begin)
            Dim imageF As Byte() = New Byte(r.BaseStream.Length - 1) {}
            imageF = r.ReadBytes(CInt(r.BaseStream.Length))
            fs.Close()
            fs.Dispose()
            fs = Nothing
            r.Close()
            r = Nothing
            Return imageF
        Catch ex As Exception
            Commun.Journal("ERREUR : Erreur de lecture du fichier image : " & ex.Message & " : " & imageFile, True)
        End Try
        imageFile = Nothing
    End Function

End Class
