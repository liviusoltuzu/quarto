Attribute VB_Name = "Text"
Option Explicit
Public FontBase As GLuint
Public FontBaseTip1 As GLuint
Public FontBaseTip2 As GLuint

Public Sub BuildFont(P As PictureBox, Optional Tip As Byte = 0)
If Tip <> 0 And Tip <> 1 And Tip <> 2 Then Exit Sub
    
    Dim hfont As Long
If Tip = 0 Then
    FontBase = glGenLists(96)

    hfont = CreateFont(-85, 0, 0, 0, FW_BOLD, True, False, False, _
            ANSI_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS, _
            ANTIALIASED_QUALITY, FF_DONTCARE Or DEFAULT_PITCH, "times new roman")
            

    SelectObject P.hDC, hfont

    wglUseFontBitmaps P.hDC, 32, 96, FontBase
ElseIf Tip = 1 Then
    FontBaseTip1 = glGenLists(128)

    hfont = CreateFont(-8, 0, 0, 0, FW_NORMAL, False, False, False, _
            ANSI_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS, _
            ANTIALIASED_QUALITY, FF_DONTCARE Or DEFAULT_PITCH, "ms sans serif")
            

    SelectObject P.hDC, hfont

    wglUseFontBitmaps P.hDC, 0, 127, FontBaseTip1
ElseIf Tip = 2 Then
    FontBaseTip2 = glGenLists(128)

    hfont = CreateFont(-18, 0, 0, 0, FW_NORMAL, False, False, False, _
            ANSI_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS, _
            ANTIALIASED_QUALITY, FF_DONTCARE Or DEFAULT_PITCH, "ms sans serif")
            

    SelectObject P.hDC, hfont

    wglUseFontBitmaps P.hDC, 0, 127, FontBaseTip2
End If
End Sub


Public Sub glPrint(ByVal s As String, Optional Tip As Byte = 0)
If Tip <> 0 And Tip <> 1 And Tip <> 2 Then Exit Sub
s = s & Space(1)
    Dim b() As Byte
    Dim i As Integer
    If Len(s) > 0 Then
        ReDim b(Len(s))
        For i = 1 To Len(s)
            b(i) = Asc(Mid$(s, i, 1))
        Next
        b(Len(s)) = 0
        glPushAttrib amListBit
        If Tip = 0 Then
            glListBase (FontBase - 32)
        ElseIf Tip = 1 Then
            glListBase (FontBaseTip1)
        ElseIf Tip = 2 Then
            glListBase (FontBaseTip2)
        End If
        b(0) = 0
        glCallLists Len(s) - 1, GL_UNSIGNED_BYTE, b(1)
        glPopAttrib
    End If

End Sub
