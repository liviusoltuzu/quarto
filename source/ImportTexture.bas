Attribute VB_Name = "LoadTextures"
Option Explicit

Public Texture(6) As GLuint

Public Function LoadGLTextures() As Boolean
On Error GoTo err:
' Load Bitmaps And Convert To Textures
    Dim Status As Boolean
    Dim i As Byte
    Dim w As Long
    Dim h As Long
    Dim ids(6) As Long
    Dim TextureImage() As GLbyte
    
    ZStart.lblDoing = "Loading textures..."
    ZStart.lblDoing.Refresh
    
    ids(0) = 101 'piese albe
    ids(1) = 102 'tabla si piese negre
    ids(2) = 103 'cercuri
    ids(3) = 104 'highlight
    ids(4) = 105
    
    glGenTextures 7, Texture(0)
    Status = True
    For i = 0 To 4
        DoEvents
        If Not Status Then Exit For
        Status = False
        If LoadBMP(ids(i), TextureImage(), h, w) Then
            glBindTexture glTexture2D, Texture(i)
            glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR
            glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR
            glTexImage2D glTexture2D, 0, 3, w, h, 0, tiRGB, GL_UNSIGNED_BYTE, TextureImage(0, 0, 0)
            Erase TextureImage
            Status = True
        End If
    Next
    
    Dim bmpHeight As GLfloat, bmpWidth As GLfloat
    Dim i2 As Long, j As Long, k As Long
    Dim bmpImage() As GLubyte
    Dim X As Long


    ids(5) = 106 'viewport mask
    ids(6) = 107 'viewport
    If Status Then
        For i = 5 To 6
            If Not Status Then Exit For
            Status = False


            Stage2.Picture1.Picture = LoadResPicture(ids(i), 0)
            bmpWidth = Stage2.Picture1.ScaleWidth
            bmpHeight = Stage2.Picture1.ScaleHeight

            ReDim bmpImage(bmpHeight * bmpWidth * 3 - 1)
            X = 0
            For i2 = 0 To bmpWidth - 1
                For j = 0 To bmpHeight - 1
                    k = Stage2.Picture1.Point(i2, j)

                    bmpImage(X) = (k And &HFF&)
                    bmpImage(X + 1) = ((k And &HFF00&) / &H100&)
                    bmpImage(X + 2) = ((k And &HFF0000) / &H10000)

                    X = X + 3
                Next
            Next

            glBindTexture glTexture2D, Texture(i)
            glPixelStorei pxsUnpackAlignment, 1
            glTexEnvf tetTextureEnv, tenTextureEnvMode, tepModulate
            glTexParameterf glTexture2D, tpnTextureMagFilter, GL_LINEAR
            glTexParameterf glTexture2D, tpnTextureMinFilter, GL_LINEAR
            glTexParameterf glTexture2D, tpnTextureWrapS, GL_REPEAT
            glTexParameterf glTexture2D, tpnTextureWrapT, GL_REPEAT

            glTexImage2D glTexture2D, 0, 3, bmpWidth, bmpHeight, 0, GL_RGB, GL_UNSIGNED_BYTE, bmpImage(0)
            Erase bmpImage()
            Status = True
        Next
    End If

err:
If err.Number <> 0 Or Not Status Then
    ErrPrg = True
    ErrNumb = 102
    MsgBox LoadResString(ErrNumb), vbOKOnly + vbCritical, "Quarto!"
    If LogF <> 0 Then Print #LogF, GetTime & "Error " & ErrNumb & ": " & LoadResString(ErrNumb)
    If LogF <> 0 Then Print #LogF, GetTime & "Exit"
    End
End If
End Function

Private Function LoadBMP(ByVal id, ByRef Texture() As GLuint, ByRef Height As Long, ByRef Width As Long) As Boolean
On Error GoTo err:
    'Dim intFileHandle As Integer
    'Dim bitmapheight As Long
    'Dim bitmapwidth As Long

    ' Open a file.
    ' The file should be BMP with pictures 64x64,128x128,256x256 .....
  
    'If UCase(Right(Filename, 3)) = "BMP" Then
        Stage2.Picture1.Picture = LoadResPicture(id, 0)
        CreateTextureMapFromImage Stage2.Picture1, Texture(), Height, Width
    'ElseIf UCase(Right(Filename, 3)) = "MOT" Then
    '    intFileHandle = FreeFile
    '    Open Filename For Binary Access Read Lock Read Write As intFileHandle
    '    Get #intFileHandle, , Width
    '    Get #intFileHandle, , Height
    '    ReDim bitmapImage(2, Height - 1, Width - 1)
    '    Get #intFileHandle, , Texture
    '    Close intFileHandle
    'End If
    LoadBMP = True
err:
If err.Number <> 0 Then
    ErrPrg = True
    ErrNumb = 102
    MsgBox LoadResString(ErrNumb), vbOKOnly + vbCritical, "Quarto!"
    If LogF <> 0 Then Print #LogF, GetTime & "Error " & ErrNumb & ": " & LoadResString(ErrNumb)
    If LogF <> 0 Then Print #LogF, GetTime & "Exit"
    End
End If
End Function


Private Sub CreateTextureMapFromImage(pict As PictureBox, ByRef TextureImg() As GLbyte, ByRef Height As Long, ByRef Width As Long)
On Error GoTo err:

    ' Create the array as needed for the image.
    pict.ScaleMode = 3                  ' Pixels
    Height = pict.ScaleHeight
    Width = pict.ScaleWidth
    
    ReDim TextureImg(2, Height - 1, Width - 1)
    
    ' Fill the array with the bitmap data...  This could take
    ' a while...
    
    Dim X As Long, Y As Long
    Dim C As Long
    
    Dim yloc As Long
    For X = 0 To Width - 1
        For Y = 0 To Height - 1
            C = pict.Point(X, Y)                ' Returns in long format.
            yloc = Height - Y - 1
            TextureImg(0, X, yloc) = C And 255
            TextureImg(1, X, yloc) = (C And 65280) \ 256
            TextureImg(2, X, yloc) = (C And 16711680) \ 65536
        Next Y
    Next X
err:
If err.Number <> 0 Then
    ErrPrg = True
    ErrNumb = 102
    MsgBox LoadResString(ErrNumb), vbOKOnly + vbCritical, "Quarto!"
    If LogF <> 0 Then Print #LogF, GetTime & "Error " & ErrNumb & ": " & LoadResString(ErrNumb)
    If LogF <> 0 Then Print #LogF, GetTime & "Exit"
    End
End If
   
End Sub

