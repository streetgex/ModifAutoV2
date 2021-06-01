Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.DirectoryServices
Imports System.IO



Public Class Thumbn


    Shared fichierThumbTemp As String = "c:\temp\thumbTemp.jpg"

    'Public Shared srv As String = "Space2"
    'Public Shared pathPhoto As String = "\\" & Thumbn.srv & "\photos-AD\AD\"
    'Public Shared pathThumbs As String = "\\" & Thumbn.srv & "\photos-AD\AD\Thumb\"
    Public Shared srv As String = "Labo4"
    Public Shared pathPhoto As String = "\\" & Thumbn.srv & "\photos-RH\"
    Public Shared pathThumbs As String = "\\" & Thumbn.srv & "\photos-RH\Thumbnails\"

    'Shared Sub ComparePhoto(ByVal cn As String, ByVal matricule As String, ByVal objDirEnt As DirectoryEntry)
    Shared Sub ComparePhoto(ByVal matricule As String, ByVal objDirEnt As DirectoryEntry)
        Try
            Dim fichierPhoto As String = pathPhoto & matricule & ".jpg"
            'Dim sFile As New System.IO.FileInfo(fichierPhoto)
            'Dim testFichier As Boolean = sFile.Exists
            'sFile = Nothing
            'If testFichier Then

            'File.Copy(fichierPhoto, "S:\Mes bibliothèques\Photos\" & objDirEnt.Properties("displayNamePrintable").Value & "-" & matricule & ".jpg")

            Dim imageOriginale As Byte() = LireImage(fichierPhoto)

            Dim testCompare As Boolean = False

            If Not objDirEnt.Properties("jpegPhoto").Value Is Nothing Then
                Dim imageAD As Byte() = DirectCast(objDirEnt.Properties("jpegPhoto")(0), Byte())
                testCompare = CompareImageBytesIdentique(imageAD, imageOriginale)
                imageAD = Nothing
            End If

            If testCompare = False Then

                objDirEnt.Properties("jpegPhoto").Value = Nothing
                objDirEnt.Properties("thumbnailPhoto").Value = Nothing
                Commun.AppliquerChangement(objDirEnt)

                CreateThumb(matricule)
                Dim imageFile As String = fichierThumbTemp
                Dim imageTh As Byte() = LireImage(imageFile)
                objDirEnt.Properties("jpegPhoto").Insert(0, imageOriginale)
                objDirEnt.Properties("thumbnailPhoto").Insert(0, imageTh)
                Commun.AppliquerChangement(objDirEnt)
                Commun.Journal("Modification des attributs Photos Réussi : " & objDirEnt.Properties("SAMAccountName").Value)
                imageTh = Nothing
                Kill(fichierThumbTemp)
            End If
            testCompare = Nothing
            imageOriginale = Nothing

            'Else
            '    If Not objDirEnt.Properties("jpegPhoto").Value Is Nothing Or Not objDirEnt.Properties("thumbnailPhoto").Value Is Nothing Then
            '        'Si pas de photo dans le dossier, on efface l'attribut AD
            '        objDirEnt.Properties("jpegPhoto").Value = Nothing
            '        objDirEnt.Properties("thumbnailPhoto").Value = Nothing
            '        Commun.AppliquerChangement(objDirEnt)
            '        Commun.Journal("Nettoyage des attributs Photos Réussi : " & objDirEnt.Properties("SAMAccountName").Value)
            '    End If
            'End If

        Catch e As Exception
            Commun.Journal("ERREUR : Comparaison Thumbnail : " & Replace(objDirEnt.Path, "LDAP://", ""), True)
            Return
        End Try
    End Sub
    Shared Sub CreateThumb(ByVal matricule As String)

        Dim sFile As New System.IO.FileInfo(fichierThumbTemp)
        Dim testFichier As Boolean = sFile.Exists
        sFile = Nothing
        If testFichier Then
            Kill(fichierThumbTemp)
        End If

        Try

            Dim image As Image = Nothing
            Dim imgThumb As Image = Nothing

            Dim myImageCodecInfo As ImageCodecInfo
            Dim myEncoder As Encoder
            Dim myEncoderParameter As EncoderParameter
            Dim myEncoderParameters As EncoderParameters
            myImageCodecInfo = GetEncoderInfo("image/jpeg")
            myEncoder = Encoder.ColorDepth
            myEncoderParameters = New EncoderParameters(1)
            myEncoderParameter = New EncoderParameter(myEncoder, CType(24L, Int32))
            myEncoderParameters.Param(0) = myEncoderParameter


            Dim fichierPhoto As String = pathThumbs & matricule & ".jpg"
            Dim tailleFichierThumb As Long = GetFileLength(fichierThumbTemp)


            If tailleFichierThumb = -1 Then
                'File.Copy(fichierPhoto, fichierThumbTemp)

                image = image.FromFile(fichierPhoto)
                Dim original As Bitmap = image
                image = original.Clone
                original.Dispose()
                original = Nothing
                image.Save(fichierThumbTemp, myImageCodecInfo, myEncoderParameters)
            End If

            'DECOUPAGE PAYSAGE

            If image.Width > image.Height Then
                image = image.FromFile(fichierThumbTemp)
                Dim focusRectangle As New Rectangle()
                Dim original As Bitmap = image
                focusRectangle.X = (original.Width - original.Height * 0.9) / 2
                focusRectangle.Y = 0
                focusRectangle.Height = original.Height - 1
                focusRectangle.Width = (original.Height * 0.9)
                image = original.Clone(focusRectangle, PixelFormat.DontCare)
                original.Dispose()
                original = Nothing
                image.Save(fichierThumbTemp, myImageCodecInfo, myEncoderParameters)
            End If
            image.Dispose()
            image = Nothing


            tailleFichierThumb = GetFileLength(fichierThumbTemp)

            While tailleFichierThumb > 10240 'Or tailleFichierThumb = -1

                image = image.FromFile(fichierThumbTemp)
                imgThumb = image.GetThumbnailImage(image.Width / 2, image.Height / 2, Nothing, New IntPtr())
                image.Dispose()
                image = Nothing
                imgThumb.Save(fichierThumbTemp, myImageCodecInfo, myEncoderParameters)
                'ecriture du fichier thumbnails
                'imgThumb.Save(pathThumbs & "\" & matricule & ".jpg", myImageCodecInfo, myEncoderParameters)
                imgThumb.Dispose()
                imgThumb = Nothing
                tailleFichierThumb = GetFileLength(fichierThumbTemp)
            End While
            myEncoderParameter.Dispose()
            myEncoderParameter = Nothing
            myEncoderParameters.Dispose()
            myEncoderParameters = Nothing
            myImageCodecInfo = Nothing
            myEncoder = Nothing
            GC.Collect()
        Catch e As Exception
            Commun.Journal("ERREUR : Création Thumbnail : " & matricule, True)
            Return
        End Try
        'If sFile.Exists Then
        '    Kill(fichierThumbTemp)
        'End If
    End Sub
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
    Shared Function controlServeurZoneLabo()
        'controle si les zones labo sont disponibles
        Dim ctrlZonesLabo As Boolean = Directory.Exists("\\Labo4\CtrlZones")
        Dim ctrlZonesNxprod As Boolean = Directory.Exists("\\Nxprod\CtrlZones")
        If ctrlZonesLabo = True Then
            Thumbn.srv = "Labo4"
            ctrlZonesLabo = True
        ElseIf ctrlZonesNxprod = True Then
            Thumbn.srv = "Nxprod"
            ctrlZonesLabo = True
        Else
            ctrlZonesLabo = False
        End If

        Return ctrlZonesLabo
    End Function
End Class
