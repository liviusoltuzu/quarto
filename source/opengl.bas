Attribute VB_Name = "Graphics"
Option Explicit
Public aflLightPosition(4) As GLfloat

Private Type BITMAP           '14 bytes
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Long
  bmBitsPixel As Long
  bmBits As Long
End Type

Public Type BITMAPINFOHEADER  '40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Public Type RGBQUAD           '4 bytes
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

'viewport variables
Global MainW As Long
Global MainH As Long
Global MWMinWOp2 As GLfloat '(mainw-widthoptions)\2
Global MHMinHOp2 As GLfloat '(mainh-heightoptions)\2
Public URW As Long 'top-right viewport width
Public URH As Long 'top-right viewport height
Type ViewportXY
    X As GLdouble
    Y As GLdouble
End Type
Public Const KtMenu = 4
Public ViewXY() As ViewportXY 'coords la patratele negre din lista de "givepiece"
Public MenuXY(1 To KtMenu) As ViewportXY 'coordonatele la cele ktmenu meniuri
Public KTpiese As Byte 'cate piese au mai ramas in Piese (lista cu cele 16 piese)
Public StartView As Byte 'de la ce pozitie incepe sa afiseze in lista de givepiece
Public NrSqr As Integer 'numarul de patrate care intra pe width fereastra

'gameplay variables
Public DelayPutGivenP As Double 'cat sta ca sa puna piesa/sa dea piesa

Public KtRolledMenu As Integer
Type AltTip
    Items As String
    Rolled As Boolean
End Type
Public Menus(1 To KtMenu) As AltTip
Public KeepMenu As Boolean
Public ShowOptions As Boolean
Public ShowMessage As Boolean
'Public ShowAbout As Boolean
'Public AboutDelay As GLfloat


'KEYBOARD
Const KtKeyCtrl = 4
Public KeyControl As Byte '1:picklist 2:tabla 3:menu 4:options
Public PosPickList As Byte
Public PosTabla As Byte


'rotation variables
Public xRot As Double
Public yRot As Double
Public zRot As Double
Private UpRZ As Double 'viewport "up" rotation Z
Private ZoomUpRZ As Double
Private ZoomUpScale As Boolean 'true=plus false=minus
Public Zoom As Double 'z-ul pentru tabla, main vieport
Public ZoomRate As Integer 'cu cat se impart coord obiectelor

Type UnType
    coord As Double
    bool As Boolean 'doar pentru 1-16 ca sa stie cand creste si cand scade
    bool2 As Boolean 'prima oara creste cu 0.02 apoi cu 0.01 si ca sa stie daca a crescut o data
End Type
Public XYZCoord(1 To 21) As UnType '1 to 16 pentru picklist, 17 pentru tabla,18 font marime,19 x=lumina
                                        ' 20 y=lumina 21 z=lumina
Public oneRendered As Boolean

'Type CreditsTip
'    nume As String
'    Cx As GLfloat
'    Cy As GLfloat
'End Type
'Dim CreditDelay As Single
'Dim CreditX As GLfloat, CreditY As GLfloat
'Dim CeCrd As Long
'Dim CredFin As Long
'Dim Credits(15) As CreditsTip
'Public ShowCredits As Boolean

Public ShowHtext As Boolean

Public UseDblClkFnc As Boolean
Public TblDblClk(1 To 4, 1 To 4) As Boolean 'doubleclicked pentru tabla
Public LstDblClk(1 To 16) As Boolean 'doubleclicked pentru lista

Dim FPS As Integer, fps2 As Integer
Dim Sec As Double
Dim CeFPSAsStr As String
'Private Transp(0 To 16) As Boolean
'Private TranspPiece As Byte
'Private TranspRate As Double

Private hRC As Long
Public HitsList(1 To 50) As Boolean
Public HitsListConst As Byte 'cate vals am in hitslist
Public MenuHits(1 To 50) As Boolean
Public MenuHitsConst As Byte 'cate vals am menuhits
Public MsgBoxHits(1 To 50) As Boolean
Public MsgBoxHitsConst As Byte
Public OptionsHits(1 To 50) As Boolean
Public Const OptionsHitsConst = 15 'cate valori am hitslist, cate casute/butoane am in options
Public MouseX As Single, MouseY As Single
Public Hits As GLint

'tipuri de mesaje
'1: doar cu ok
'2: cu yes, no si cancel
Public TipMesaj As Byte
Public glPrompt As String
Public glSett As String


Public Sub ResizeGLScene(w As GLsizei, h As GLsizei)
    If h = 0 Then
      h = 1
  End If

  glViewport 0, 0, w, h
  
  glMatrixMode mmProjection
  glLoadIdentity

  'Calculate The Aspect Ratio Of The Window
  gluPerspective 45, w / h, 0.1, 1000#

  'Select The Modelview Matrix and reset it
  glMatrixMode mmModelView
  glLoadIdentity

End Sub
Function CreateGL(P As PictureBox)

Dim bRet As Boolean

Dim pfd As PIXELFORMATDESCRIPTOR
Dim lPixelFormat As Long
Dim lRet As Long
    
  pfd.nVersion = 1
  pfd.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
  pfd.iLayerType = PFD_MAIN_PLANE
  pfd.iPixelType = PFD_TYPE_RGBA
  pfd.cColorBits = 16
  pfd.cDepthBits = 16
  pfd.nSize = Len(pfd)
  
  'Request the desired pixel format
  lPixelFormat = ChoosePixelFormat(P.hDC, pfd)
  If (lPixelFormat = 0) Then
    'Failed to get desired pixel format, so bail out
    ErrPrg = True
    ErrNumb = 103
    'MsgBox "Could not get desired Pixel Format"
    Exit Function
  End If
  
    lRet = SetPixelFormat(P.hDC, lPixelFormat, pfd)
  If (lRet = 0) Then
    'Failed to set pixel format, so bail out
    'MsgBox "Could not set desired Pixel Format"
    ErrPrg = True
    ErrNumb = 103
    Exit Function
  End If

hRC = wglCreateContext(P.hDC)
  If (hRC = 0) Then
    'Failed to create rendering context, so bail out
    ErrPrg = True
    ErrNumb = 103
    Exit Function
  End If

  lRet = wglMakeCurrent(P.hDC, hRC)
  If (lRet = 0) Then
    'Failed to activate the RC, so bail out
    ErrPrg = True
    ErrNumb = 103
    
    Exit Function
  End If

  'ResizeGLScene P.ScaleWidth, P.ScaleHeight
  bRet = InitGL()
  If (bRet = False) Then
    'Could not initialize OpenGL, so bail out
    ErrPrg = True
    ErrNumb = 103
    MsgBox LoadResString(ErrNumb), vbOKOnly + vbCritical, "Quarto!"
    If LogF <> 0 Then Print #LogF, GetTime & "Error " & ErrNumb & ": " & LoadResString(ErrNumb)
    If LogF <> 0 Then Print #LogF, GetTime & "Exit"
    End
    'CreateGLWindow = False
    Exit Function
  End If

End Function
Public Function InitGL() As Boolean
' All Setup For OpenGL Goes Here

  'Load the texture(s)
  If LogF <> 0 Then Print #LogF, GetTime & "creating lists..."
      CreateGLLists
If LogF <> 0 Then Print #LogF, GetTime & "done."
If LogF <> 0 Then Print #LogF, GetTime & "loading textures..."
  LoadGLTextures
  If LogF <> 0 Then Print #LogF, GetTime & "done."
  If ErrPrg Or (ErrNumb <> 0) Then
    InitGL = False
    Exit Function
  End If


  glEnable glcTexture2D               ' Enable Texture Mapping
  glShadeModel smSmooth                ' Enables Smooth Shading

  glClearColor 0#, 0#, 0#, 1           ' Black Background
  glClearDepth 1#                     ' Depth Buffer Setup
   
  glEnable glcDepthTest               ' Enables Depth Testing
  glDepthFunc cfLEqual                ' The Type Of Depth Test To Do
  'glHint htLineSmoothHint, hmNicest
  'glHint htPolygonSmoothHint, hmNicest
  'glHint htPointSmoothHint, hmNicest
  glHint htPerspectiveCorrectionHint, hmNicest    ' Really Nice Perspective Calculations

    glEnable glcLighting
    'Set the light settings
    Dim aflLightAmbient(4) As GLfloat
    Dim aflLightDiffuse(4) As GLfloat
    'Dim aflLightPosition(4) As GLfloat
  
  'Ambient settings
  aflLightAmbient(0) = 0.8
  aflLightAmbient(1) = 0.8
  aflLightAmbient(2) = 0.8
  aflLightAmbient(3) = 1#
  'Diffuse settings
  aflLightDiffuse(0) = 1#
  aflLightDiffuse(1) = 1#
  aflLightDiffuse(2) = 1#
  aflLightDiffuse(3) = 1#
  'Position settings
  aflLightPosition(0) = 0#
  aflLightPosition(1) = 1#
  aflLightPosition(2) = 1#
  aflLightPosition(3) = 1#
  
  'Now set up the light in OpenGL
  glLightfv ltLight1, lpmAmbient, aflLightAmbient(0)
  glLightfv ltLight1, lpmDiffuse, aflLightDiffuse(0)
  glLightfv ltLight1, lpmPosition, aflLightPosition(0)
  'aflLightPosition(0) = 0#
  'aflLightPosition(1) = 0#
  'aflLightPosition(2) = 0#
  'aflLightPosition(3) = 1#
  'glLightfv ltLight2, lpmPosition, aflLightPosition(0)
  'aflLightPosition(0) = 0#
  'aflLightPosition(1) = 0#
  'aflLightPosition(2) = 1#
  'aflLightPosition(3) = 0#
  'glLightfv ltLight3, lpmPosition, aflLightPosition(0)
  'aflLightPosition(0) = 0#
  'aflLightPosition(1) = 1#
  'aflLightPosition(2) = 0#
  'aflLightPosition(3) = 0#
  'glLightfv ltLight4, lpmPosition, aflLightPosition(0)

  glEnable glcLight1

  'And enable the light
  glColor4f 1, 1, 1, 0.5
  glBlendFunc sfSrcAlpha, dfOne

  'Initialization Went OK
  InitGL = True
End Function


Public Sub Render(P As PictureBox)
Dim bDone As Boolean
bDone = False


ZoomUpRZ = -0.5
ZoomUpScale = True
StartView = 1
xRot = 36
zRot = 8
'Stage2.scrZoom.Value = 3
Menus(1).Items = "File;New game;|;Exit"
Menus(3).Items = "Options;Gameplay;|;Save options"
Menus(2).Items = "Pieces;Light and dark;  Light;  Dark;|;Short and tall;  Short;  Tall;|;" & _
                    "Round and square;  Round;  Square;|;Hollow and solid;  Hollow;  Solid;|;|;Reset all"
Menus(4).Items = "Help;Contents"

'ShowCredits = False
'CeCrd = 0
'CredFin = 0
'Credits(1).nume = "Ilici"
'Credits(2).nume = "Johnny"
'Credits(3).nume = "Moondy"
fps2 = 30

Dim tz As Byte
For tz = 1 To KtMenu
    Menus(tz).Rolled = False
Next
'frmQuarto.Caption = "Quarto " & Str$(Timer - dU)
If LogF <> 0 Then Print #LogF, GetTime & "Rendering..."
While Not bDone
    DoEvents
    
    If (DrawGLScene() = False) Then
      'Quit the app
      ErrPrg = True
      ErrNumb = 104
      MsgBox LoadResString(ErrNumb), vbOKOnly + vbCritical, "Quarto!"
      If LogF <> 0 Then Print #LogF, GetTime & "Error " & ErrNumb & ": " & LoadResString(ErrNumb)
      If LogF <> 0 Then Print #LogF, GetTime & "Exit"
      End
      bDone = True
    Else
      'We're not quitting, so update the screen and process events
      SwapBuffers P.hDC
      DoEvents
    End If
    If Not oneRendered Then
        oneRendered = True
    End If
Wend
End Sub

Public Function DrawGLScene()
'If CineMutaAcum Then
'    Stage2.Frame2.Visible = False
'Else
'    Stage2.Frame2.Visible = True
'End If
'Se calc FPS-ul
FPS = FPS + 1
's = Timer

If Timer - Sec >= 1 Then
    CeFPSAsStr = "FPS: " & CStr(Format((FPS / (Timer - Sec)), "0.00"))
    fps2 = FPS
    FPS = 0
    Sec = Timer
End If
'Pana aici


'dU = Timer

If (DelayPutGivenP < Timer) And CineMutaAcum And Not TbSaDeaPiesa And Not Ilose And Not Youlose And Not Remiza Then
    MutaCalc
    CineMutaAcum = False
End If

'On Error GoTo err:
    glClear clrColorBufferBit
    
    
    glClear GL_DEPTH_BUFFER_BIT
    'GoTo zy
                                                                                    'MAIN VIEWPORT
    glViewport 0, 0, MainW, MainH
    glMatrixMode mmProjection
    glLoadIdentity
    gluPerspective 45, MainW / MainH, 0.1, 1000#
    glMatrixMode mmModelView
    glLoadIdentity
    
    glTranslatef 0#, 0#, Zoom
    glRotatef -90, 1, 0, 0
    'glRotatef -8, 0, 0, 1
    'glRotatef 36, 1, 0, 0
    
    glRotatef xRot, 1, 0, 0
    glRotatef yRot, 0, 1, 0
    glRotatef zRot, 0, 0, 1

    glEnable glcTexture2D
    glEnable glcLighting
    glDisable glcBlend
    glEnable glcDepthTest
    
    glHint htLineSmoothHint, hmNicest
    
    glBindTexture glTexture2D, Texture(1)
    glCallList 9
    
    glBindTexture glTexture2D, Texture(2)
    glCallList 10
    If Not CineMutaAcum And Not TbSaDeaPiesa And Not Ilose And Not Youlose And Not Remiza Then
        If (MouseY > (HeightMenu + HeightPickList)) And Not KeepMenu And Not ShowOptions Then Selection
        
        Dim i As Byte
        glBindTexture glTexture2D, Texture(3)
        For i = 1 To 16
            If HitsList(i) And Tabla(Mid(Lists(10 + i).ObjName, 2, 1), Mid(Lists(10 + i).ObjName, 3, 1)).ESett = 0 Then
                glCallList i + 10
                HitsList(17) = False
            End If
        Next
    ElseIf Youlose Or Ilose And Not Remiza Then
        Dim i1 As Integer, i2 As Integer
        For i = 1 To 19
            glBindTexture glTexture2D, Texture(3)
            If ArrayVicts(i).ACuloare <> 2 Or ArrayVicts(i).BMarime <> 2 Or ArrayVicts(i).CForma <> 2 Or ArrayVicts(i).DSoliditate <> 2 Then
                Select Case i
                Case 1 To 4 'linii h
                    For i1 = 1 To 4
                        glCallList 10 + (i - 1) * 4 + i1
                    Next
                Case 5 To 8 'linii v
                    For i1 = 1 To 4
                        glCallList 10 + (i1 - 1) * 4 + i - 4
                    Next
                Case 9 'daig princ
                    For i1 = 1 To 4
                        glCallList 10 + (i1 - 1) * 5 + 1
                    Next
                Case 10 'diag sec
                    For i1 = 1 To 4
                        glCallList 10 + i1 * 3 + 1
                    Next
                Case 11 To 19
                    For i1 = 1 To 2
                        For i2 = 1 To 2
                            If (i - 10) Mod 3 <> 0 Then
                                glCallList i + ((i - 10) \ 3) + (i1 - 1) * 4 + i2 - 1
                            Else
                                glCallList i + ((i - 10) \ 3 - 1) + (i1 - 1) * 4 + i2 - 1
                            End If
                        Next
                    Next
                End Select
            End If
        Next
    End If

    PutPieces
    DoEvents
  
    
'zy:    'GoTo z
    glClear GL_DEPTH_BUFFER_BIT
                                                                        'VIEWPORT ORTHO: BACKGROUND
    glViewport 0, 0, MainW, MainH
    glMatrixMode mmProjection
    glLoadIdentity
    glOrtho 0#, MainW, MainH, 0#, -1#, 1#
    glMatrixMode mmModelView
    glLoadIdentity
    
    
    glEnable glcTexture2D
    glDisable glcLighting
    glEnable glcBlend
    glDisable glcDepthTest
    
    'numai daca cursorul este in zona de sus apeleza selectia
    If (MouseY <= (HeightMenu + HeightPickList)) And (MouseY > HeightMenu) And Not KeepMenu And Not ShowOptions Then SelectionOrtho
    
    Dim TransX As GLdouble, TransY As GLdouble 'retine cu cat se deplaseaza pe x si y cand deseneza figurile
    Dim k As GLdouble 'in ce dist tb sa se incadreze cele nrsqr patrate
    TransX = 0
    TransY = 0
    k = WidthPickList - 2 * WidthPickListButton - 2 * 10
    NrSqr = k \ (HWPiece + 10)
    If KTpiese < NrSqr Then NrSqr = KTpiese
    If StartView + NrSqr - 1 > KTpiese And StartView <> 1 Then
        StartView = KTpiese - NrSqr + 1
    End If
    glBindTexture glTexture2D, Texture(4)
        glColor4f 1, 1, 1, 0.8
        glBlendFunc sfSrcAlpha, dfOne
    'glCallList 39
    'glTranslatef 0, HeightMenu, 0
    'menu
    glCallList 31
        glColor4f 1, 1, 1, 0.4
    'picklist
    glTranslatef (MainW - WidthPickList) / 2, HeightMenu, 0
    TransX = TransX + (MainW - WidthPickList) / 2
    TransY = TransY + HeightMenu
    glCallList 27
        '"not" options
        'glTranslatef 0, HeightPickList, 0
            'glColor4f 1, 1, 1, 0.8
            'glBlendFunc sfSrcAlpha, dfOne
        'glCallList 31
        'glTranslatef 0, -HeightPickList, 0
         '   glColor4f 1, 1, 1, 0.4
    'leftbutton
    glTranslatef 10, (HeightPickList - HeightPickListButton) / 2, 0
    TransX = TransX + 10
    TransY = TransY + (HeightPickList - HeightPickListButton) / 2
    glCallList 30 'deseneaza butonul (din stanga)
        glBlendFunc sfSrcAlpha, dfOneMinusDstAlpha
    'triunghiul negru din button; numai daca se poate face stanga
    If StartView > 1 Then
        If HitsList(20) Then
            glColor4f 1, 1, 1, 0.1
        Else
            glColor4f 1, 1, 1, 0.6
        End If
    glTranslatef (WidthPickListButton - WidthPLBTriangle) / 2, (HeightPickListButton - HeightPLBTriangle) / 2, 0
    If NrSqr <> 0 Then glCallList 32 'deseneza triunghiul negru
    glTranslatef -(WidthPickListButton - WidthPLBTriangle) / 2, -(HeightPickListButton - HeightPLBTriangle) / 2, 0
    End If
    
       'glColor4f 1, 1, 1, 0.2
    If NrSqr > 0 Then 'in caz ca exista cel putin un patrat care intra in width
            glTranslatef WidthPickListButton, -10, 0
            TransX = TransX + WidthPickListButton
            TransY = TransY - 10
    
        glTranslatef (k - (NrSqr * (HWPiece + 10))) / 2, 0, 0
        TransX = TransX + (k - (NrSqr * (HWPiece + 10))) / 2
        ReDim ViewXY(NrSqr + 1) '1-ul este pentru giveNpiece (viewport dreapta sus/jos)
        Dim DkEsteSgnClkPs As Boolean 'daca a fost dat click pe o pozitie din lista
        Dim MouseDsp As Boolean 'mouse over
        If UseDblClkFnc Then
        For i = 1 To 16
            If LstDblClk(i) Then DkEsteSgnClkPs = True: Exit For
        Next
        End If
        For i = 1 To CByte(NrSqr)
            glTranslatef 5, 0, 0
            TransX = TransX + 5: ViewXY(i).X = TransX: ViewXY(i).Y = TransY 'se incarca viewxy
            
            MouseDsp = (MouseX > ViewXY(i).X) And (MouseX < (ViewXY(i).X + HWPiece)) And _
                (MouseY > ViewXY(i).Y) And (MouseY < (ViewXY(i).Y + HWPiece))
            If (MouseDsp Or (LstDblClk(AvPieces(i + StartView - 1))) Or (Not UseDblClkFnc And MouseDsp)) And _
                Not CineMutaAcum And TbSaDeaPiesa And Not KeepMenu And Not ShowOptions And Not ShowMessage _
                 Then
                    glColor4f 1, 1, 1, 0.3
                    'aici cand mouseul este deasupra piesei
                    If XYZCoord(i + StartView - 1).bool Then
                        If XYZCoord(i + StartView - 1).bool2 Then
                            XYZCoord(i + StartView - 1).coord = XYZCoord(i + StartView - 1).coord + 0.01
                        Else
                            XYZCoord(i + StartView - 1).coord = XYZCoord(i + StartView - 1).coord + 0.02
                        End If
                        If XYZCoord(i + StartView - 1).coord > CDbl(-0.42) Then
                            XYZCoord(i + StartView - 1).bool = False
                            XYZCoord(i + StartView - 1).bool2 = True
                        End If
                        
                    Else
                        XYZCoord(i + StartView - 1).coord = XYZCoord(i + StartView - 1).coord - CDbl(0.01)
                        If XYZCoord(i + StartView - 1).coord < CDbl(-0.5) Then
                            XYZCoord(i + StartView - 1).bool = True
                        End If
                    End If
            Else
                    'daca mouseul e pe alta piesa, marimea curentei scade
                    If Abs(XYZCoord(i + StartView - 1).coord) <= Abs(-0.5) Then
                        XYZCoord(i + StartView - 1).coord = XYZCoord(i + StartView - 1).coord - CDbl(0.02)
                        XYZCoord(i + StartView - 1).bool2 = False
                    End If
                    glColor4f 1, 1, 1, 0.2
            End If
            glCallList 29 'deseneaza patratul
            glTranslatef HWPiece + 5, 0, 0
            TransX = TransX + HWPiece + 5
        Next
        glTranslatef (k - (NrSqr * (HWPiece + 10))) / 2, 10, 0
        TransX = TransX + (k - (NrSqr * (HWPiece + 10))) / 2
        TransY = TransY + 10
            glColor4f 1, 1, 1, 0.4
            glBlendFunc sfSrcAlpha, dfOne
        glCallList 30 'deseneaza butonul (din dreapta)
        If HitsList(21) Then
            glColor4f 1, 1, 1, 0.1
        Else
            glColor4f 1, 1, 1, 0.6
        End If
            glBlendFunc sfSrcAlpha, dfOneMinusDstAlpha
        glTranslatef (WidthPickListButton - WidthPLBTriangle) / 2, (HeightPickListButton - HeightPLBTriangle) / 2, 0
        If StartView + NrSqr - 1 <> KTpiese Then
            glCallList 33 'deseneaza triunghiul
        End If
        glTranslatef -(WidthPickListButton - WidthPLBTriangle) / 2, -(HeightPickListButton - HeightPLBTriangle) / 2, 0
    
    Else 'dk nu exista nici un patrat care poate fi desenat in width
        glTranslatef WidthPickListButton, 0, 0
        TransX = TransX + WidthPickListButton
        glTranslatef k, 0, 0
        TransX = TransX + k
            glColor4f 1, 1, 1, 0.4
            glBlendFunc sfSrcAlpha, dfOne
        glCallList 30 'deseneaza butonul (din dreapta)
        If HitsList(21) Then
            glColor4f 1, 1, 1, 0.1
        Else
            glColor4f 1, 1, 1, 0.6
        End If
            glBlendFunc sfSrcAlpha, dfOneMinusDstAlpha
        glTranslatef (WidthPickListButton - WidthPLBTriangle) / 2, (HeightPickListButton - HeightPLBTriangle) / 2, 0
        'If StartView + NrSqr - 1 <> KTpiese Then
        'glCallList 33 'deseneaza triunghiul
        'End If
        glTranslatef -(WidthPickListButton - WidthPLBTriangle) / 2, -(HeightPickListButton - HeightPLBTriangle) / 2, 0
    End If
    
        glColor4f 1, 1, 1, 0.4
        glBlendFunc sfSrcAlpha, dfOne
    glTranslatef -TransX, -TransY, 0
    TransX = 0: TransY = 0
    ViewXY(NrSqr + 1).X = MainW - HWGivenPiece - 30
    If CineMutaAcum Then
        ViewXY(NrSqr + 1).Y = HeightMenu + HeightPickList + 25
    Else
        ViewXY(NrSqr + 1).Y = MainH - HWGivenPiece - 25 'HeightMenu + HeightPickList + 25
    End If
    glTranslatef ViewXY(NrSqr + 1).X, ViewXY(NrSqr + 1).Y, 0
        glColor4f 1, 1, 1, 0.2
    If ((Not CineMutaAcum And Not TbSaDeaPiesa) Or (CineMutaAcum)) And (Not Ilose And Not Youlose And Not Remiza) Then glCallList 34
    glTranslatef -ViewXY(NrSqr + 1).X, -ViewXY(NrSqr + 1).Y, 0
    
'    glLoadIdentity
'    glTranslatef 0, MainH - HeightMenu, 0
'        glBindTexture glTexture2D, Texture(4)
'        glColor4f 1, 1, 1, 0.8
'
'    glCallList 31
'    glLoadIdentity
    
    'glDisable glcBlend
    'glCallList 36
    'glTranslatef 550, 300, 0
    'GLCreateWindowsShape 1, 60, 40, ""
    'glEnable glcBlend
    'On Error Resume Next
    
    If Youlose Or Ilose Or Remiza Then
        glColor4f 1, 1, 1, 0.5
        On Error Resume Next
        If Format$(XYZCoord(20).coord, "0.00") <> Format$(-1, "0.00") And _
            Format$(XYZCoord(21).coord, "0.00") <> Format$(0.05, "0.00") Then
            XYZCoord(20).coord = XYZCoord(20).coord - 0.05
            XYZCoord(21).coord = XYZCoord(21).coord - 0.05
        End If
        aflLightPosition(0) = 0
        aflLightPosition(1) = 1
        aflLightPosition(2) = XYZCoord(20).coord
        aflLightPosition(3) = XYZCoord(21).coord
  
      glLightfv ltLight1, lpmPosition, aflLightPosition(0)
Else
'reset lights
aflLightPosition(0) = 0#
aflLightPosition(1) = 1#
aflLightPosition(2) = 1#
aflLightPosition(3) = 1#
  
glLightfv ltLight1, lpmPosition, aflLightPosition(0)
'end lights

    End If
    
    glEnable glcLighting
    glDisable glcBlend
    glEnable glcDepthTest
    'GoTo ti
    DoEvents
                                                                         
                                                                         
                                                                         'VIEWPORTS FROM PICKLIST
    Dim l As Byte
    l = StartView
                
    For i = 1 To NrSqr
        glClear clrDepthBufferBit
        glViewport ViewXY(i).X, MainH - ViewXY(i).Y - HWPiece, URW, URH
        glMatrixMode mmProjection
        glLoadIdentity
        gluPerspective 25, URW / URH, 0.1, 100#
        glMatrixMode mmModelView
        glLoadIdentity
        
        glTranslatef -0.01!, -0.07!, XYZCoord(i + l - 1).coord  '-0.41!
        'glTranslatef -0.01!, -0.06!, -0.5!
        
        glRotatef -50, 1, 0, 0
        glRotatef 3, 0, 1, 0
        
        If (MouseX > ViewXY(i).X) And (MouseX < (ViewXY(i).X + HWPiece)) And _
            (MouseY > ViewXY(i).Y) And (MouseY < (ViewXY(i).Y + HWPiece)) Then
            'Stage2.Stage.MousePointer = 99
        End If

        'l = 0
        If Piese(AvPieces(i + l - 1)).ACuloare = 0 Then
            glBindTexture glTexture2D, Texture(0)
        Else
            glBindTexture glTexture2D, Texture(1)
        End If
        glCallList Piese(AvPieces(i + l - 1)).DSoliditate * 2 ^ 0 + _
            Piese(AvPieces(i + l - 1)).CForma * 2 ^ 1 + _
            Piese(AvPieces(i + l - 1)).BMarime * 2 ^ 2 + 1
            'l = l + 1
    Next




                                                                'VIEWPORT TOP-RIGHT/BOTTOM-RIGHT
    If Not Ilose And Not Youlose And Not Remiza Then

    glClear GL_DEPTH_BUFFER_BIT
    glViewport ViewXY(NrSqr + 1).X, MainH - ViewXY(NrSqr + 1).Y - URH, URW, URH
    glMatrixMode mmProjection
    glLoadIdentity
    gluPerspective 25, URW / URH, 0.1, 100#
    glMatrixMode mmModelView
    glLoadIdentity
    
    glTranslatef 0#, -0.06!, ZoomUpRZ
    'glTranslatef 0#, 0#, ZoomUpRZ
    
    glRotatef -50, 1, 0, 0
    glRotatef 3, 0, 1, 0
    UpRZ = UpRZ + 1
    glRotatef UpRZ, 0, 0, 1

    If Not CineMutaAcum And Not TbSaDeaPiesa Then
        If Piese(PiesaForU).ACuloare = 0 Then
            glBindTexture glTexture2D, Texture(0)
        Else
            glBindTexture glTexture2D, Texture(1)
        End If
        glCallList Piese(PiesaForU).DSoliditate * 2 ^ 0 + _
        Piese(PiesaForU).CForma * 2 ^ 1 + Piese(PiesaForU).BMarime * 2 ^ 2 + 1
    ElseIf CineMutaAcum Then
        If Piese(PiesaForC).ACuloare = 0 Then
            glBindTexture glTexture2D, Texture(0)
        Else
            glBindTexture glTexture2D, Texture(1)
        End If
        glCallList Piese(PiesaForC).DSoliditate * 2 ^ 0 + _
        Piese(PiesaForC).CForma * 2 ^ 1 + Piese(PiesaForC).BMarime * 2 ^ 2 + 1
    End If
End If
 
    
    glEnable glcTexture2D
    glDisable glcLighting
    glEnable glcBlend
    glDisable glcDepthTest
    glBindTexture glTexture2D, Texture(4)

                                                                        'PIECES (TEXT)
    
'    glClear GL_DEPTH_BUFFER_BIT
'
'    glViewport 0, 0, MainW, MainH
'    glMatrixMode mmProjection
'    glLoadIdentity
'    glOrtho 0#, MainW, MainH, 0#, -1#, 1#
'    glMatrixMode mmModelView
'    glLoadIdentity
'
'    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
'    glColor4f 1, 1, 1, 1
'
'    Dim RasterLeft As GLfloat
'    Dim RasterTop As GLfloat
'    RasterTop = HeightMenu + HeightPickList
'    Dim s As String
'    RasterLeft = 20
'    glBegin bmLineStrip
'        glVertex2d RasterLeft - 2 + 8, RasterTop + 12
'        glVertex2d RasterLeft - 2, RasterTop + 12
'        glVertex2d RasterLeft - 2, RasterTop + 15 * 5 + 5
'        glVertex2d RasterLeft - 2 + 160, RasterTop + 15 * 5 + 5
'        glVertex2d RasterLeft - 2 + 160, RasterTop + 12
'        glVertex2d RasterLeft - 2 + 80, RasterTop + 12
'    glEnd
'    glLoadIdentity
'    glClear GL_DEPTH_BUFFER_BIT
'    glRasterPos2f RasterLeft, RasterTop + 15
'    s = "   Pieces: " & KTpiese & "/" & AllPieces
'    glPrint s, 1
'    glLoadIdentity
'    glClear GL_DEPTH_BUFFER_BIT
'    glRasterPos2f RasterLeft, RasterTop + 15 * 2
'    s = ""
'    If ktLightA <> 0 Then s = Space(5) & "Light: " & ktLight & "/" & ktLightA & ""
'    If ktDarkA <> 0 Then s = s & Space(5) & "Dark: " & ktDark & "/" & ktDarkA & ""
'    glPrint s, 1
'    glLoadIdentity
'    glClear GL_DEPTH_BUFFER_BIT
'    glRasterPos2f RasterLeft, RasterTop + 15 * 3
'    s = ""
'    If ktShortA <> 0 Then s = Space(5) & "Short: " & ktShort & "/" & ktShortA & ""
'    If ktTallA <> 0 Then s = s & Space(5) & "Tall: " & ktTall & "/" & ktTallA & ""
'    glPrint s, 1
'    glLoadIdentity
'    glClear GL_DEPTH_BUFFER_BIT
'    glRasterPos2f RasterLeft, RasterTop + 15 * 4
'    s = ""
'    If ktRoundA <> 0 Then s = Space(5) & "Round: " & ktRound & "/" & ktRoundA & ""
'    If ktSquareA <> 0 Then s = s & Space(5) & "Square: " & ktSquare & "/" & ktSquareA & ""
'    glPrint s, 1
'    glLoadIdentity
'    glClear GL_DEPTH_BUFFER_BIT
'    glRasterPos2f RasterLeft, RasterTop + 15 * 5
'    s = ""
'    If ktHollowA <> 0 Then s = Space(5) & "Hollow: " & ktHollow & "/" & ktHollowA & ""
'    If ktSolidA <> 0 Then s = s & Space(5) & "Solid: " & ktSolid & "/" & ktSolidA & ""
'    glPrint s, 1
   

    
    
    glClear GL_DEPTH_BUFFER_BIT

                                                                        'VIEWPORT MENU
    glViewport 0, 0, MainW, MainH
    glMatrixMode mmProjection
    glLoadIdentity
    glOrtho 0#, MainW, MainH, 0#, -1#, 1#
    glMatrixMode mmModelView
    glLoadIdentity
       
    If ((MouseY <= HeightMenu) Or KeepMenu) And Not ShowOptions Then SelectionMenu
    
    Dim KtSep As Byte, KtPsiV As Byte, e As Byte 'e=counter; ktsep=cati separatori; ktpsiv=cate ";"
    Dim LungRolledMenu As Integer
    
    'Menus(1).Rolled = 0: Menus(2).Rolled = 0: Menus(3).Rolled = 0
    glLoadIdentity
    glTranslatef 0, (HeightMenu - HeightMenuTitle) / 2, 0
    TransX = 0
    TransY = (HeightMenu - HeightMenuTitle) / 2
    For i = 1 To KtMenu
        glTranslatef 10, 0, 0
            TransX = TransX + 10
        If Menus(i).Rolled And KeepMenu Then
            glTranslatef -1, -1, 0
                TransX = TransX - 1
                TransY = TransY - 1
            
            KtSep = 0
            KtPsiV = 0
            For e = 1 To Len(Menus(i).Items)
                Select Case Mid(Menus(i).Items, e, 1)
                Case ";"
                    KtPsiV = KtPsiV + 1
                Case "|"
                    KtSep = KtSep + 1
                End Select
            Next
            LungRolledMenu = (KtPsiV - KtSep) * (HeightRolledMenu + 1) + KtSep * 2
            KtRolledMenu = KtPsiV - KtSep
                glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
                glColor4f 1, 0, 0, 1
                
            glBegin bmPolygon
                glVertex2i 0, HeightMenu + LungRolledMenu
                glVertex2i WidthRolledMenu + 2, HeightMenu + LungRolledMenu
                glVertex2i WidthRolledMenu + 2, HeightMenu - 1
                glVertex2i WidthMenuTitle + 2, HeightMenu - 1
                glVertex2i WidthMenuTitle + 2, 0
                glVertex2i 0, 0
                'glVertex2i 0, 0
            glEnd
            
            glTranslatef 1, 1, 0
                TransX = TransX + 1
                TransY = TransY + 1
                MenuXY(i).X = TransX: MenuXY(i).Y = TransY
            GLCreateMenu i, , , True
            glTranslatef WidthMenuTitle, 0, 0
            TransX = TransX + WidthMenuTitle
        Else
            If LungRolledMenu <> 0 Then
                glTranslatef 0, -LungRolledMenu + 7, 0
                    TransY = TransY - LungRolledMenu + 7
                LungRolledMenu = 0
            End If
                MenuXY(i).X = TransX: MenuXY(i).Y = TransY
            GLCreateMenu i, , , False
            glTranslatef WidthMenuTitle, 0, 0
            TransX = TransX + WidthMenuTitle
        End If
    Next
        
                                                                    

                                                                                     'MENU TEXT
    glClear GL_DEPTH_BUFFER_BIT

    glViewport 0, 0, MainW, MainH
    glMatrixMode mmProjection
    glLoadIdentity
    glOrtho 0#, MainW, MainH, 0#, -1#, 1#
    glMatrixMode mmModelView
    glLoadIdentity

    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glColor4f 1, 1, 1, 1
    Dim j As Byte, Lcm As Byte 'la ce rolled menu a ajuns
    Dim RasterX As Integer, RasterY As Integer

    'title file
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f 28, 14
    glPrint "File", 1
    'title pieces
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f 79, 14
    glPrint "Pieces", 1
    'title options
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f 137, 14
    glPrint "Options", 1
    'title help
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f 204, 14
    glPrint "Help", 1
    
    '& Space(11) & "Options" & Space(9) & "Pieces" & Space(11) & "Help", 1
    
    'Menus(3).Rolled = True
    'KeepMenu = True
    'glLoadIdentity
    'glClear GL_DEPTH_BUFFER_BIT

    RasterX = 0: RasterY = 14
    For i = 1 To KtMenu
        If Menus(i).Rolled And KeepMenu Then
            
            Select Case i
            Case 1
                RasterX = 28
            Case 2
                RasterX = 79
            Case 3
                RasterX = 137
            Case 4
                RasterX = 204
            End Select
            'Stage2.Check1.Visible = False
            Dim a() As String
            a = Split(Menus(i).Items, ";")
            Lcm = 0
            For j = 1 To UBound(a)
                If a(j) <> "|" Then
                    Lcm = Lcm + 1
                    If j > 1 Then
                        RasterY = RasterY + 3
                    End If
                    RasterY = RasterY + 23
                    glLoadIdentity
                    glClear GL_DEPTH_BUFFER_BIT
                    glRasterPos2f RasterX, RasterY
                    If a(0) = "Pieces" Then
                    If MenuPieces(Lcm) And Lcm <> 13 Then
                        glPrint "* " & a(j), 1
                    Else
                        If Lcm <> 13 Then
                            glPrint "  " & a(j), 1
                        Else
                            glPrint a(j), 1
                        End If
                    End If
                    Else
                        glPrint a(j), 1
                    End If
                Else
                    RasterY = RasterY + 2
                End If
            Next
        Exit For
        End If
    Next
'    glLoadIdentity
'    glClear GL_DEPTH_BUFFER_BIT
'    glRasterPos2f 250, 14
'    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
''    glColor4d 1, 1, 1, 0.5
'    glPrint CeFPSAsStr, 1
    
                                                                            'YOU LOSE/YOU HAVE WON
    glClear GL_DEPTH_BUFFER_BIT

    glViewport 0, 0, MainW, MainH
    glMatrixMode mmProjection
    glLoadIdentity
    glOrtho 0#, MainW, MainH, 0#, -1#, 1#
    glMatrixMode mmModelView
    glLoadIdentity
    
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glColor4f 1, 1, 1, 0.7

    'glDisable glcBlend
    glClear GL_DEPTH_BUFFER_BIT
    glLoadIdentity
    If Youlose Then
        glRasterPos2s 50, MainH - 100
        glPrint "You lose!"
    ElseIf Ilose Then
        glRasterPos2s 50, MainH - 100
        glPrint "You win!"
    ElseIf Remiza Then
        glRasterPos2s 50, MainH - 100
        glPrint "Draw!"
    End If
    If Youlose Or Ilose Then
        Dim ktRnd As Integer 'cate randuri tb sa scrie=cate victorii sunt pe tabla
        ktRnd = 0
        For i = 1 To 19
        If ArrayVicts(i).ACuloare <> 2 Then
            glClear GL_DEPTH_BUFFER_BIT
            glRasterPos2s 100, CDbl(MainH - 100 + 25 + ktRnd * 25)
            glPrint ktRnd + 1 & ". " & GetCuvFromNumb(1, ArrayVicts(i).ACuloare) & " pieces", 2
            ktRnd = ktRnd + 1
        End If
        If ArrayVicts(i).BMarime <> 2 Then
            glClear GL_DEPTH_BUFFER_BIT
            glRasterPos2s 100, CDbl(MainH - 100 + 25 + ktRnd * 25)
            glPrint ktRnd + 1 & ". " & GetCuvFromNumb(2, ArrayVicts(i).BMarime) & " pieces", 2
            ktRnd = ktRnd + 1
        End If
        If ArrayVicts(i).CForma <> 2 Then
            glClear GL_DEPTH_BUFFER_BIT
            glRasterPos2s 100, CDbl(MainH - 100 + 25 + ktRnd * 25)
            glPrint ktRnd + 1 & ". " & GetCuvFromNumb(3, ArrayVicts(i).CForma) & " pieces", 2
            ktRnd = ktRnd + 1
        End If
        If ArrayVicts(i).DSoliditate <> 2 Then
            glClear GL_DEPTH_BUFFER_BIT
            glRasterPos2s 100, CDbl(MainH - 100 + 25 + ktRnd * 25)
            glPrint ktRnd + 1 & ". " & GetCuvFromNumb(4, ArrayVicts(i).DSoliditate) & " pieces", 2
            ktRnd = ktRnd + 1
        End If
        Next
    End If

If Not Youlose And Not Ilose And ShowHtext Then
glClear GL_DEPTH_BUFFER_BIT
glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
glColor4f 1, 1, 1, 1

glRasterPos2f 10, MainH - 10
If CineMutaAcum Then
    glPrint "[Computer's turn]", 1
Else
    If Not TbSaDeaPiesa Then
        glPrint "[Your turn] Put the piece on the board.", 1
    Else
        glPrint "[Your turn] Give a piece to the computer.", 1
    End If
End If

'glViewport ViewXY(NrSqr + 1).X, MainH - ViewXY(NrSqr + 1).Y - URH, URW, URH
glClear GL_DEPTH_BUFFER_BIT
If Not TbSaDeaPiesa Then
    glRasterPos2f ViewXY(NrSqr + 1).X, ViewXY(NrSqr + 1).Y - 5
    If CineMutaAcum Then
    glPrint "Computer's piece:", 1
    Else
    glPrint "Your piece:", 1
    End If
End If
End If

If ShowOptions Then
                                                                              'OPTIONS
    glClear GL_DEPTH_BUFFER_BIT

    glViewport 0, 0, MainW, MainH
    glMatrixMode mmProjection
    glLoadIdentity
    glOrtho 0#, MainW, MainH, 0#, -1#, 1#
    glMatrixMode mmModelView
    glLoadIdentity
    If Not ShowMessage Then SelectionOptions

    glLoadIdentity
'  Options
    glTranslatef MWMinWOp2, MHMinHOp2, 0
       ' glBlendFunc sfOneMinusDstAlpha, dfSrcAlpha
        glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
        glColor4f 1, 1, 1, 0.2
        glCallList 37
    glLoadIdentity

    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glColor4f 1, 1, 1, 1
    
    glBegin bmLineStrip
        glVertex2d MWMinWOp2 + WidthOptions \ 2 - 110, MHMinHOp2 + 10
        glVertex2d MWMinWOp2 + 10, MHMinHOp2 + 10
        glVertex2d MWMinWOp2 + 10, MHMinHOp2 + HeightOptions - 10
        glVertex2d MWMinWOp2 + WidthOptions - 10, MHMinHOp2 + HeightOptions - 10
        glVertex2d MWMinWOp2 + WidthOptions - 10, MHMinHOp2 + 10
        glVertex2d MWMinWOp2 + WidthOptions \ 2 - 2, MHMinHOp2 + 10
    glEnd
    'glTranslatef 50, 50, 0

    
    'glBlendFunc sfZero, dfSrcAlpha
    'glColor4f 1, 1, 1, 0

    glBegin bmLines
        glVertex2d MWMinWOp2 + 15, MHMinHOp2 + 15 + 17
        glVertex2d MWMinWOp2 + 64, MHMinHOp2 + 15 + 17
        glVertex2d MWMinWOp2 + 15 + 4 + 7, MHMinHOp2 + 15 + 17
        glVertex2d MWMinWOp2 + 15 + 4 + 7, MHMinHOp2 + 125 + 17
        
    glEnd
    
    glBegin bmLines
        glVertex2d MWMinWOp2 + 150, MHMinHOp2 + 15 + 17
        glVertex2d MWMinWOp2 + 42 + 135, MHMinHOp2 + 15 + 17
        glVertex2d MWMinWOp2 + 15 + 4 + 7 + 135, MHMinHOp2 + 15 + 17
        glVertex2d MWMinWOp2 + 15 + 4 + 7 + 135, MHMinHOp2 + 125 + 17
        
    glEnd
        
    glBegin bmLines
        glVertex2d MWMinWOp2 + 100, MHMinHOp2 + 165
        glVertex2d MWMinWOp2 + 300, MHMinHOp2 + 165
    glEnd
    
    'glBlendFunc sfOneMinusDstAlpha, dfOneMinusSrcAlpha
    'glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
    glColor4f 1, 1, 1, 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + WidthOptions \ 2 - 102, MHMinHOp2 + 10 + 5
    glPrint "Options - Gameplay", 1

    

    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 15, MHMinHOp2 + 15 + 15
    glPrint "First move:", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 150, MHMinHOp2 + 15 + 15
    glPrint "Level:", 1

    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 15 + 17 + 8, MHMinHOp2 + 35 + 17
    glPrint "Alternative", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 15 + 17 + 8, MHMinHOp2 + 35 + 4 + 41 + 17
    glPrint "Player", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 15 + 17 + 8, MHMinHOp2 + 35 + 4 + 41 + 4 + 41 + 17
    glPrint "Computer", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 15 + 17 + 8, MHMinHOp2 + 175 + 17
    glPrint "Use 'two clicks' method", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 15 + 17 + 8 + 175, MHMinHOp2 + 175 + 17
    glPrint "Show instructions", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 150 + 17 + 8, MHMinHOp2 + 35 + 17
    glPrint "Beginner", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 150 + 17 + 8, MHMinHOp2 + 35 + 4 + 41 + 17
    glPrint "Normal", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 150 + 17 + 8, MHMinHOp2 + 35 + 4 + 41 + 4 + 41 + 17
    glPrint "Advanced", 1
If StilDeJoc = 1 Then
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 250 + 17 + 8, MHMinHOp2 + 35 + 17
    glPrint "Light and dark", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 250 + 17 + 8, MHMinHOp2 + 35 + 4 + 26 + 17
    glPrint "Short and tall", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 250 + 17 + 8, MHMinHOp2 + 35 + 4 + 26 + 4 + 26 + 17
    glPrint "Round and square", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 250 + 17 + 8, MHMinHOp2 + 35 + 4 + 26 + 4 + 26 + 4 + 26 + 17
    glPrint "Hollow and solid", 1
End If

    
    
'whostarts
glLoadIdentity
glClear GL_DEPTH_BUFFER_BIT
glTranslatef MWMinWOp2 + 15, MHMinHOp2 + 35, 0
    DrawSquareFade
    DrawOptionBox WhoStarts = 0
glTranslatef 0, 4 + 41, 0
    DrawSquareFade
    DrawOptionBox WhoStarts = 1
glTranslatef 0, 4 + 41, 0
    DrawSquareFade
    DrawOptionBox WhoStarts = 2
'stildejoc
glLoadIdentity
glTranslatef MWMinWOp2 + 150, MHMinHOp2 + 35, 0
    DrawSquareFade
    DrawOptionBox StilDeJoc = 1
glTranslatef 0, 4 + 41, 0
    DrawSquareFade
    DrawOptionBox StilDeJoc = 2
glTranslatef 0, 4 + 41, 0
    DrawSquareFade
    DrawOptionBox StilDeJoc = 3
If StilDeJoc = 1 Then
    glLoadIdentity
    glTranslatef MWMinWOp2 + 250, MHMinHOp2 + 35, 0
        DrawSquareFade
        DrawCheckBox SingleOption(1)
    glTranslatef 0, 4 + 26, 0
        DrawSquareFade
        DrawCheckBox SingleOption(2)
    glTranslatef 0, 4 + 26, 0
        DrawSquareFade
        DrawCheckBox SingleOption(3)
    glTranslatef 0, 4 + 26, 0
        DrawSquareFade
        DrawCheckBox SingleOption(4)
End If

glLoadIdentity
glTranslatef MWMinWOp2 + 15, MHMinHOp2 + 175, 0
DrawSquareFade
DrawOptionBox UseDblClkFnc
glTranslatef 175, 0, 0
DrawSquareFade
'DoEvents
DrawOptionBox ShowHtext
'DoEvents
'glTranslatef 0, 4 + 26, 0
'DrawSquareFade
'DrawOptionBox True


glLoadIdentity
glTranslatef MWMinWOp2, MHMinHOp2 + 225, 0
'glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
glBlendFunc sfSrcAlpha, dfOneMinusDstAlpha
glColor4f 1, 1, 1, 0.1
glTranslatef 10, 0, 0
glTranslatef 35, 0, 0
glCallList 45
glTranslatef 35 + WidthButton, 0, 0
'glCallList 45
glTranslatef 35 + WidthButton, 0, 0
glCallList 45
glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
glColor4f 1, 1, 1, 1

    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 10 + 35 + 22, MHMinHOp2 + 225 + 17
    glPrint "Default", 1
    
    'glLoadIdentity
    'glClear GL_DEPTH_BUFFER_BIT
    'glRasterPos2f MWMinWOp2 + 10 + 35 + WidthButton + 35 + 28, MHMinHOp2 + 225 + 17
    'glPrint "Save", 1
    
    glLoadIdentity
    glClear GL_DEPTH_BUFFER_BIT
    glRasterPos2f MWMinWOp2 + 10 + 35 + WidthButton + 35 + WidthButton + 35 + 25, MHMinHOp2 + 225 + 17
    glPrint "Close", 1
End If 'options

                                                                'MSGBOX
'If ShowMessage Then
'    glClear GL_DEPTH_BUFFER_BIT
'
'    glViewport 0, 0, MainW, MainH
'    glMatrixMode mmProjection
'    glLoadIdentity
'    glOrtho 0#, MainW, MainH, 0#, -1#, 1#
'    glMatrixMode mmModelView
'    glLoadIdentity
'
'    If TipMesaj <> 0 Then
'
'        Select Case TipMesaj
'        Case 1
'            glPrompt = "Saved."
'            glSett = "1"
'            SelectionMsgBox
'            GLMsgBox ' "Saved.", "1"
'        End Select
'        'TipMesaj = 0
'    End If
'
'
'End If

'If ShowCredits Then
'                                                                'CREDITS
'    glClear GL_DEPTH_BUFFER_BIT
'
'    glViewport 0, 0, MainW, MainH
'    glMatrixMode mmProjection
'    glLoadIdentity
'    glOrtho 0#, MainW, MainH, 0#, -1#, 1#
'    glMatrixMode mmModelVie
'    glLoadIdentity
'
'    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
'
'
''On Error Resume Next
'    'CreditDelay = 0
'glColor4f 1, 1, 1, 0.2
'    For i = 1 To CredFin
'            glClear GL_DEPTH_BUFFER_BIT
'            glLoadIdentity
'            glRasterPos2f Credits(i).Cx, Credits(i).Cy
'            glPrint Credits(i).nume
'        Next
'If CeCrd <= 3 Then
'
'    If CeCrd = 0 Then CeCrd = CeCrd + 1
'    If 0 = Format(CreditDelay, "0.00") Then
'        CreditDelay = Timer + 3.4
'        CreditX = CInt((Rnd * (MainW / 2)) + 100)
'        CreditY = CInt((Rnd * (MainH / 2)) + 100)
'    End If
'
'    If CreditDelay - Timer > 0 Then
'        If ((CreditDelay - Timer) / 100 * 15) > 0.2 Then
'            glColor4f 1, 1, 1, (CreditDelay - Timer) / 100 * 15
'
'        Else
'            glColor4f 1, 1, 1, 0.2
'
'        End If
'        glClear GL_DEPTH_BUFFER_BIT
'        glLoadIdentity
'        glRasterPos2f CreditX, CreditY
'        glPrint Credits(CeCrd).nume
'    Else
'        CreditDelay = 0
'       CredFin = CredFin + 1
'             CeCrd = CeCrd + 1
'            Credits(CeCrd - 1).Cx = CreditX
'        Credits(CeCrd - 1).Cy = CreditY
'        'ReDim Preserve Credits(CeCrd - 1)
'        'Credits(CeCrd - 1).nume = Crd(CeCrd)
'
'        'CredFin = CredFin + 1
'    End If
'End If
'End If

'If ShowAbout Then
'    glClear GL_DEPTH_BUFFER_BIT
'
'    glViewport 0, 0, MainW, MainH
'    glMatrixMode mmProjection
'    glLoadIdentity
'    glOrtho 0#, MainW, MainH, 0#, -1#, 1#
'    glMatrixMode mmModelView
'    glLoadIdentity
'    If Not ShowMessage Then SelectionOptions
'
'    glLoadIdentity
'    'glTranslatef MWMinWOp2, MHMinHOp2, 0
'       ' glBlendFunc sfOneMinusDstAlpha, dfSrcAlpha
'        glBlendFunc sfZero, dfOneMinusSrcAlpha
'        If (1 - (AboutDelay - Timer)) >= 1 Then
'            glColor4f 1, 1, 1, 1
'        Else
'            glColor4f 1, 1, 1, (1 - (AboutDelay - Timer))
'        End If
'        'glBegin bmPolygon
'        '    glVertex2d (MainW - 300) \ 2, (MainH - 300) \ 2
'        '    glVertex2d (MainW - 300) \ 2 + 300, (MainH - 300) \ 2
'        '    glVertex2d (MainW - 300) \ 2 + 300, (MainH - 300) \ 2 + 300
'        '    glVertex2d (MainW - 300) \ 2, (MainH - 300) \ 2 + 300
'        'glEnd
'        glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
'        glColor4f 1, 1, 1, 1
'        glClear GL_DEPTH_BUFFER_BIT
'        glRasterPos2f 100, 100
'        glPrint "Quarto!", 1
''        If (1 - (AboutDelay - Timer)) >= 1 Then
''        glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
''        glColor4f 1, 0, 0, 0.5
''        glClear GL_DEPTH_BUFFER_BIT
''        glRasterPos2f 100, 100
''        'glPrint "Quarto!", 0
''
''        glClear GL_DEPTH_BUFFER_BIT
''End If
'End If

'ti:
    DoEvents
    glFlush
    DrawGLScene = True

err:
If err.Number <> 0 Then
    Exit Function
End If
End Function
Private Sub DrawSquareFade()
glBlendFunc sfDstColor, dfZero
glBindTexture glTexture2D, Texture(6)
glCallList 38
glBlendFunc sfZero, dfOne
glBindTexture glTexture2D, Texture(5)
glCallList 38
End Sub

Private Sub DrawOptionBox(Optional optValue As Boolean = False, Optional TranslatefX As GLfloat = 6, Optional TranslatefY As GLfloat = 6)
glTranslatef TranslatefX, TranslatefY, 0
glBindTexture glTexture2D, Texture(4)
        glBlendFunc sfSrcAlpha, dfOneMinusDstAlpha
        glColor4f 1, 1, 1, 0
    glCallList 39
        glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
        glColor4f 1, 1, 1, 1
    glTranslatef 2, 2, 0
    glCallList 40
    If optValue Then
        glTranslatef 2, 2, 0
            glBlendFunc sfSrcAlpha, dfOneMinusDstAlpha
            glColor4f 1, 1, 1, 0
        glCallList 41
        glTranslatef -2, -2, 0
    End If
glTranslatef -TranslatefX - 2, -TranslatefY - 2, 0
End Sub
Private Sub DrawCheckBox(Optional chkValue As Boolean = False, Optional TranslatefX As GLfloat = 6, Optional TranslatefY As GLfloat = 6)
glTranslatef TranslatefX, TranslatefY, 0
glBindTexture glTexture2D, Texture(4)
        glBlendFunc sfSrcAlpha, dfOneMinusDstAlpha
        glColor4f 1, 1, 1, 0
    glCallList 42
        glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
        glColor4f 1, 1, 1, 1
    glTranslatef 2, 2, 0
    glCallList 43
    If chkValue Then
            glBlendFunc sfSrcAlpha, dfOneMinusDstAlpha
            glColor4f 1, 1, 1, 0
        glTranslatef 2, 2, 0
        glCallList 44
        glTranslatef -2, -2, 0
    End If
glTranslatef -TranslatefX - 2, -TranslatefY - 2, 0
End Sub
Public Sub Selection()
Dim i As Integer
Dim Buffer(512) As GLuint
'Dim Hits As GLint
Dim Viewport(4) As GLint
glGetIntegerv GL_VIEWPORT, Viewport(0)
glSelectBuffer 512, Buffer(0)

glRenderMode GL_SELECT
glInitNames
glPushName 0

glMatrixMode mmProjection
glPushMatrix
glLoadIdentity
gluPickMatrix MouseX, Viewport(3) - MouseY, 1#, 1#, Viewport(0)

gluPerspective 45, (Viewport(2) - Viewport(0)) / (Viewport(3) - Viewport(1)), 0.1, 1000#
glMatrixMode mmModelView
DrawSPerspective
glMatrixMode mmProjection
glPopMatrix
glMatrixMode mmModelView
Hits = glRenderMode(rmRender)

If Hits > 0 Then
    Dim choose As GLuint
    Dim depth As GLuint
    choose = Buffer(3)
    depth = Buffer(1)
    For i = 1 To Hits - 1
        If Buffer(4 * i + 1) < depth Then
            choose = Buffer(4 * i + 3)
            depth = Buffer(4 * i + 1)
        End If
    Next
    HitsList(choose) = True
End If

If choose <> 0 Then
    If choose <> 17 Then
        For i = 1 To 16
            HitsList(i) = False
        Next
        HitsList(choose) = True
        i = choose
        If Tabla(Mid(Lists(10 + i).ObjName, 2, 1), Mid(Lists(10 + i).ObjName, 3, 1)).ESett = 0 Then
            'Stage2.Stage.MousePointer = 99
        End If
    Else
        For i = 1 To 17
            HitsList(i) = False
        Next
        If Not ((MouseX > (MainW - HWGivenPiece - 30) And MouseX < (MainW - 30)) And _
                (MouseY > (HeightMenu + HeightPickList + 25) And MouseY < (HeightMenu + HeightPickList + 25 + HWGivenPiece))) Then
            Stage2.Stage.MousePointer = 0
        End If
    End If
End If
    
    
End Sub

Public Sub SelectionMsgBox()
Dim i As Integer
Dim Buffer(512) As GLuint
Dim choose As Integer
Hits = 0
'Dim Hits As GLint
Dim Viewport(4) As GLint
glGetIntegerv GL_VIEWPORT, Viewport(0)
glSelectBuffer 512, Buffer(0)

glRenderMode GL_SELECT
glInitNames
glPushName 0

glMatrixMode mmProjection
glPushMatrix
glLoadIdentity
gluPickMatrix MouseX, Viewport(3) - MouseY, 1#, 1#, Viewport(0)

glOrtho 0#, MainW, MainH, 0#, -1#, 1#
glMatrixMode mmModelView
DrawSMsgBox
glMatrixMode mmProjection
glPopMatrix
glMatrixMode mmModelView
Hits = glRenderMode(rmRender)

For i = 1 To MsgBoxHitsConst
    MsgBoxHits(i) = False
Next

If Hits <> 0 Then
    MsgBoxHits(Buffer(3)) = True
End If

End Sub
Private Sub DrawSPerspective()
Dim i As Byte

'glLoadIdentity
'glTranslatef 0, 0, Zoom
For i = 1 To 16
    glLoadName i
    glPushMatrix
    glCallList 10 + i
    glPopMatrix
Next
glLoadName 17
glPushMatrix
glCallList 9
glPopMatrix
HitsListConst = 17
End Sub
Private Sub SelectionOptions()
Dim i As Integer
Dim Buffer(512) As GLuint
Dim choose As Integer
Hits = 0
'Dim Hits As GLint
Dim Viewport(4) As GLint
glGetIntegerv GL_VIEWPORT, Viewport(0)
glSelectBuffer 512, Buffer(0)

glRenderMode GL_SELECT
glInitNames
glPushName 0

glMatrixMode mmProjection
glPushMatrix
glLoadIdentity
gluPickMatrix MouseX, Viewport(3) - MouseY, 1#, 1#, Viewport(0)

glOrtho 0#, MainW, MainH, 0#, -1#, 1#
glMatrixMode mmModelView
DrawSOptions
glMatrixMode mmProjection
glPopMatrix
glMatrixMode mmModelView
Hits = glRenderMode(rmRender)
    
For i = 1 To OptionsHitsConst
    OptionsHits(i) = False
Next

If Hits <> 0 Then
    OptionsHits(Buffer(3)) = True
    'Debug.Print Hits
End If
End Sub
Public Sub SelectionOrtho()
Dim i As Integer
Dim Buffer(512) As GLuint
Dim choose As Integer
Hits = 0
'Dim Hits As GLint
Dim Viewport(4) As GLint
glGetIntegerv GL_VIEWPORT, Viewport(0)
glSelectBuffer 512, Buffer(0)

glRenderMode GL_SELECT
glInitNames
glPushName 0

glMatrixMode mmProjection
glPushMatrix
glLoadIdentity
gluPickMatrix MouseX, Viewport(3) - MouseY, 1#, 1#, Viewport(0)

glOrtho 0#, MainW, MainH, 0#, -1#, 1#
glMatrixMode mmModelView
DrawSOrtho
glMatrixMode mmProjection
glPopMatrix
glMatrixMode mmModelView
Hits = glRenderMode(rmRender)

For i = 1 To HitsListConst
    HitsList(i) = False
Next

If Hits <> 0 Then
    choose = Buffer(4 * 1 - 1)
    For i = 2 To Hits
        If Buffer(4 * i - 1) > choose Then
            choose = Buffer(4 * i - 1)
        End If
    Next
Select Case choose
Case 19
    Stage2.Stage.MousePointer = 0
Case 20 To HitsListConst
    HitsList(choose) = True
    'Stage2.Stage.MousePointer = 99
End Select
    
End If
End Sub

Public Sub SelectionMenu()
Dim i As Integer
Dim Buffer(512) As GLuint
Dim choose As Integer
Hits = 0
'Dim Hits As GLint
Dim Viewport(4) As GLint
glGetIntegerv GL_VIEWPORT, Viewport(0)
glSelectBuffer 512, Buffer(0)

glRenderMode GL_SELECT
glInitNames
glPushName 0

glMatrixMode mmProjection
glPushMatrix
glLoadIdentity
gluPickMatrix MouseX, Viewport(3) - MouseY, 1#, 1#, Viewport(0)

glOrtho 0#, MainW, MainH, 0#, -1#, 1#
glMatrixMode mmModelView
'If Not KeepMenu Then
'    For i = 1 To KtMenu
'        Menus(i).Rolled = False
'    Next
'End If
DrawSOrthoMenu
'Stage2.BorderStyle = 0
glMatrixMode mmProjection
glPopMatrix
glMatrixMode mmModelView
Hits = glRenderMode(rmRender)

If Hits <> 0 Then
    If Not KeepMenu Then
        For i = 1 To KtMenu
            MenuHits(i) = False
        Next
    MenuHits(Buffer(3)) = True
    Else
        If Buffer(3) <= KtMenu Then
            For i = 1 To KtMenu
                Menus(i).Rolled = False
            Next
            Menus(Buffer(3)).Rolled = True
            
            For i = 1 To KtMenu + KtRolledMenu
                MenuHits(i) = False
            Next
            MenuHits(Buffer(3)) = True
        Else
            For i = (KtMenu + 1) To KtMenu + KtRolledMenu
                MenuHits(i) = False
            Next
            MenuHits(Buffer(3)) = True
        End If
    End If
    
Else
    If Not KeepMenu Then
        For i = 1 To KtMenu
            MenuHits(i) = False
        Next
    Else
        For i = (KtMenu + 1) To KtMenu + KtRolledMenu
            MenuHits(i) = False
        Next
    End If
End If
End Sub

Private Sub DrawSOrtho()
'    glLoadName i
'    glPushMatrix
'    glCallList 10 + i
'    glPopMatrix

    Dim TransX As GLdouble, TransY As GLdouble
    Dim k As GLdouble
    Dim j As Long
    TransX = 0
    TransY = 0
    k = WidthPickList - 2 * WidthPickListButton - 2 * 10
    j = k \ (HWPiece + 10)
    If KTpiese < j Then j = KTpiese

    'menu
    glLoadName 18: glPushMatrix
    glCallList 31
    glPopMatrix
    'picklist
    glTranslatef (MainW - WidthPickList) / 2, HeightMenu, 0:
    TransX = TransX + (MainW - WidthPickList) / 2
    TransY = TransY + HeightMenu
    glLoadName 19: glPushMatrix
    glCallList 27
    glPopMatrix
    'leftbutton
    glTranslatef 10, (HeightPickList - HeightPickListButton) / 2, 0
    TransX = TransX + 10
    TransY = TransY + (HeightPickList - HeightPickListButton) / 2
    glLoadName 20: glPushMatrix
    glCallList 30
    glPopMatrix
        
    If j <> 0 Then
        glTranslatef WidthPickListButton, -10, 0
        TransX = TransX + WidthPickListButton
        TransY = TransY - 10

    glTranslatef (k - (j * (HWPiece + 10))) / 2, 0, 0
    TransX = TransX + (k - (j * (HWPiece + 10))) / 2
    ReDim ViewXY(j)
    Dim i As Byte
    For i = 1 To CByte(j)
        glTranslatef 5, 0, 0
        TransX = TransX + 5: ViewXY(i).X = TransX: ViewXY(i).Y = TransY
        glLoadName 21 + i: glPushMatrix
        glCallList 29
        glPopMatrix
        glTranslatef HWPiece + 5, 0, 0
        TransX = TransX + HWPiece + 5
    Next
    glTranslatef (k - (j * (HWPiece + 10))) / 2, 10, 0
    TransX = TransX + (k - (j * (HWPiece + 10))) / 2
    TransY = TransY + 10
    glLoadName 21: glPushMatrix
    glCallList 30
    glPopMatrix
    Else
    glTranslatef WidthPickListButton, 0, 0
    TransX = TransX + WidthPickListButton
    glTranslatef k, 0, 0
    TransX = TransX + k
    glLoadName 21: glPushMatrix
    glCallList 30
    glPopMatrix
    End If
    
    glTranslatef -TransX, -TransY, 0
    TransX = 0: TransY = 0

HitsListConst = 17 + 2 + 2 + j
End Sub
Private Sub DrawSOptions()
'    glLoadName 20: glPushMatrix
'    glCallList 38: glPopMatrix
    
glLoadIdentity
'glClear GL_DEPTH_BUFFER_BIT
glTranslatef MWMinWOp2 + 15, MHMinHOp2 + 35, 0
    glLoadName 1: glPushMatrix
    glCallList 38: glPopMatrix
glTranslatef 0, 4 + 41, 0
    glLoadName 2: glPushMatrix
    glCallList 38: glPopMatrix
glTranslatef 0, 4 + 41, 0
    glLoadName 3: glPushMatrix
    glCallList 38: glPopMatrix

glLoadIdentity
glTranslatef MWMinWOp2 + 150, MHMinHOp2 + 35, 0
    glLoadName 4: glPushMatrix
    glCallList 38: glPopMatrix
glTranslatef 0, 4 + 41, 0
    glLoadName 5: glPushMatrix
    glCallList 38: glPopMatrix
glTranslatef 0, 4 + 41, 0
    glLoadName 6: glPushMatrix
    glCallList 38: glPopMatrix
If StilDeJoc = 1 Then
glLoadIdentity
glTranslatef MWMinWOp2 + 250, MHMinHOp2 + 35, 0
    glLoadName 7: glPushMatrix
    glCallList 38: glPopMatrix
glTranslatef 0, 4 + 26, 0
    glLoadName 8: glPushMatrix
    glCallList 38: glPopMatrix
glTranslatef 0, 4 + 26, 0
    glLoadName 9: glPushMatrix
    glCallList 38: glPopMatrix
glTranslatef 0, 4 + 26, 0
    glLoadName 10: glPushMatrix
    glCallList 38: glPopMatrix
End If

glLoadIdentity
glTranslatef MWMinWOp2 + 15, MHMinHOp2 + 175, 0
glLoadName 14: glPushMatrix 'use dbl click function
glCallList 38: glPopMatrix
glTranslatef 175, 0, 0
glLoadName 15: glPushMatrix 'use text screen
glCallList 38: glPopMatrix



glLoadIdentity
glTranslatef MWMinWOp2, MHMinHOp2 + 225, 0
glTranslatef 10, 0, 0
glTranslatef 35, 0, 0
glLoadName 11: glPushMatrix
glCallList 45: glPopMatrix
glTranslatef 35 + WidthButton, 0, 0
'glLoadName 12: glPushMatrix
'glCallList 45: glPopMatrix
glTranslatef 35 + WidthButton, 0, 0
glLoadName 13: glPushMatrix
glCallList 45: glPopMatrix


End Sub

Private Sub DrawSMsgBox()
'glLoadIdentity
'glClear GL_DEPTH_BUFFER_BIT
Dim a() As String, i As Integer, SpBet As GLfloat 'spacebetwwen
a = Split(glSett, ";")
SpBet = 30

'carou mare alb
glTranslatef (MainW - WidthMessage) \ 2, (MainH - HeightMessage) \ 2, 0

glTranslatef SpBet, HeightMessage - HeightButton - 15, 0
glTranslatef ((WidthMessage - 2 * SpBet) - ((WidthButton + SpBet) * (UBound(a) + 1))) \ 2 + SpBet \ 2, 0, 0

'buttons
For i = 1 To UBound(a) + 1
    glLoadName i: glPushMatrix
    glCallList 45: glPopMatrix
    glTranslatef SpBet + WidthButton, 0, 0
Next
MsgBoxHitsConst = (UBound(a) + 1)
End Sub
Private Sub DrawSOrthoMenu()
Dim i As Byte
    Dim KtSep As Byte, KtPsiV As Byte, e As Byte 'e=counter; ktsep=cati separatori; ktpsiv=cate ";"
    Dim LungRolledMenu As Integer
    
    glTranslatef 0, ((HeightMenu - HeightMenuTitle) / 2), 0
    For i = 1 To KtMenu
        glTranslatef 10, 0, 0
        If Menus(i).Rolled And KeepMenu Then
            glTranslatef -1, -1, 0
            
            KtSep = 0
            KtPsiV = 0
            For e = 1 To Len(Menus(i).Items)
                Select Case Mid(Menus(i).Items, e, 1)
                Case ";"
                    KtPsiV = KtPsiV + 1
                Case "|"
                    KtSep = KtSep + 1
                End Select
            Next
            LungRolledMenu = (KtPsiV - KtSep) * (HeightRolledMenu + 1) + KtSep * 2
            
            glTranslatef 1, 1, 0
            GLCreateMenuSel i, , , True
            glTranslatef WidthMenuTitle, 0, 0
        Else
            If LungRolledMenu <> 0 Then
                glTranslatef 0, -LungRolledMenu + 7, 0
                LungRolledMenu = 0
            End If
            'glTranslatef 0, ((HeightMenu - HeightMenuTitle) / 2), 0
            GLCreateMenuSel i, , , False
            glTranslatef WidthMenuTitle, 0, 0
        End If
    Next

End Sub
Public Sub PutPieces()
Dim i As Integer, j As Integer
Dim px As Double, py As Double 'pozition x & y unde tb sa fie pusa piesa
Dim dX As Double, dY As Double 'distanta dintre locatia tf si unde tb sa fie/ajunga
Dim PozX As Double, PozY As Double 'coloana 1 de pe tabla pentru px si py
Dim Pcs As Byte
Dim DkEsteSgnClkPs As Boolean 'daca a fost dat click pe o pozitie de pe tabla

If UseDblClkFnc Then
For i = 1 To 4
    For j = 1 To 4
        If TblDblClk(i, j) = True Then DkEsteSgnClkPs = True: Exit For
    Next
Next
End If

PozX = -247.5
PozY = 0
'TranspPiece = 1

glTranslatef PozX / ZoomRate, PozY / ZoomRate, 0
For i = 1 To 4
    px = PozX
    py = PozY
    
    dX = 82.5
    dY = 82.5
    For j = 1 To 4
        'deseneaza piesa
        If Tabla(i, j).ESett = 1 Then
            If Tabla(i, j).ACuloare = 0 Then
                glBindTexture glTexture2D, Texture(0)
            Else
                glBindTexture glTexture2D, Texture(1)
            End If
            Pcs = Tabla(i, j).DSoliditate * 2 ^ 0 + Tabla(i, j).CForma * 2 ^ 1 + _
                    Tabla(i, j).BMarime * 2 ^ 2 + 1
            'TranspPiece = TranspPiece + 1
            'If Not Transp(TranspPiece) Then
            '      Transp(TranspPiece - 1) = True
            '      glEnable glcBlend
            '      glDisable glcDepthTest
            'Else
            '      glDisable glcBlend
            '      glEnable glcDepthTest
            'End If
            glCallList Pcs
        ElseIf (Not TbSaDeaPiesa) And (Not CineMutaAcum) And (((HitsList((i - 1) * 4 + j) And Not DkEsteSgnClkPs) Or TblDblClk(i, j)) Or Not UseDblClkFnc And HitsList((i - 1) * 4 + j)) Then
            If Piese(PiesaForU).ACuloare = 0 Then
                glBindTexture glTexture2D, Texture(0)
            Else
                glBindTexture glTexture2D, Texture(1)
            End If
            glTranslatef 0, 0, 0.01!
            glCallList PiesaForU - Piese(PiesaForU).ACuloare * 2 ^ 3
            glTranslatef 0, 0, -0.01!
        End If
        px = px + 82.5
        py = py + 82.5
        glTranslatef dX / ZoomRate, dY / ZoomRate, 0
    Next
    PozX = PozX + 82.5
    PozY = PozY - 82.5
    dX = 165 + 82.5
    dY = 330 + 82.5
    glTranslatef -dX / ZoomRate, -dY / ZoomRate, 0
Next
glTranslatef -82.5 / ZoomRate, 330 / ZoomRate, 0
End Sub

Public Sub GLCreateShape(ByVal w As GLdouble, ByVal h As GLdouble, ListName As GLuint, Optional RoundedShape As Boolean = True, Optional Fillet As GLdouble = 5, Optional Density As Integer = 10, Optional LineLoop As Boolean = False)

Dim xp As GLdouble, yp As GLdouble '=  x' si y' (prim)=pct care se rot
Dim xpp As GLdouble, ypp As GLdouble
Dim x0 As GLdouble, y0 As GLdouble '=  x0 si y0 (zero)
Dim xf As GLdouble, yf As GLdouble '=  x si y (final)'pentru punct
Dim u As Double 'unghi
Dim pi As Double
Dim i As Long, j As Long
pi = Atn(1) * 4

glNewList ListName, lstCompile
If Not LineLoop Then
    glBegin bmPolygon
Else
    glBegin bmLineLoop
End If
If Not RoundedShape Then
    glTexCoord2d 0#, 0#: glVertex2d 0#, h
    glTexCoord2d 1#, 0#: glVertex2d w, h
    glTexCoord2d 1#, 1#: glVertex2d w, 0#
    glTexCoord2d 0#, 1#: glVertex2d 0#, 0#
Else
    xp = 0: yp = Fillet
    xf = 0: yf = Fillet
    x0 = Fillet: y0 = Fillet
    i = 1
    For j = 1 To (Density + 1) * 4
        If (j - 1) Mod (Density + 1) = 0 Then
            Select Case j
                Case 1
                    xpp = Fillet: ypp = Fillet:
                Case (Density + 2)
                    xpp = w - Fillet: ypp = Fillet:
                Case (2 * Density + 3)
                    xpp = w - Fillet: ypp = h - Fillet:
                Case 3 * Density + 4
                    xpp = Fillet: ypp = h - Fillet:
            End Select
        Else
            u = ((2 * pi) / (Density * 4) * i)
            i = i + 1
            xf = (xp - x0) * Cos(u) - (yp - y0) * Sin(u) + x0
            yf = (xp - y0) * Sin(u) + (yp - y0) * Cos(u) + y0
        End If
        glTexCoord2d (xf + xpp - Fillet) / w, 1 - ((yf + ypp - Fillet) / h):
        glVertex2d (xf + xpp - Fillet), (yf + ypp - Fillet)
        Next
End If
glEnd
glEndList
End Sub

Public Sub GLCreateTriangle(w As GLdouble, h As GLdouble, ListName As Integer, Direction As Boolean, Optional LineLoop As Boolean = False)
glNewList ListName, lstCompile
If LineLoop Then
glBegin bmLineLoop
Else
glBegin bmTriangles
End If
If Direction = True Then
glTexCoord2d 0, 1: glVertex2d 0, 0
glTexCoord2d w, h / 2: glVertex2d w, h / 2
glTexCoord2d 0, 0: glVertex2d 0, h
Else
glTexCoord2d 1, 1: glVertex2d w, 0
glTexCoord2d 0, h / 2: glVertex2d 0, h / 2
glTexCoord2d 0, 0: glVertex2d w, h
End If
glEnd
glEndList
End Sub

Public Sub GLCreateMenu(ByVal Idx As Byte, Optional ByRef Tx As GLdouble, Optional ByRef Ty As GLdouble, Optional Rolled As Boolean = False)
Dim a() As String, i As Integer
Dim mXYx As GLdouble, mXYy As GLdouble
Dim V As Integer
If Menus(Idx).Items = "" Then Exit Sub
mXYx = MenuXY(Idx).X
mXYy = MenuXY(Idx).Y
a = Split(Menus(Idx).Items, ";")

glBlendFunc sfSrcAlpha, dfSrcAlpha
If (MouseY > mXYy) And (MouseY < (mXYy + HeightMenuTitle)) And _
   (MouseX > mXYx) And (MouseX < (mXYx + WidthMenuTitle)) And Not ShowOptions Then
glColor4f 1, 1, 1, 0.1
Else
glColor4f 1, 1, 1, 0.2
End If

glCallList 35

If Rolled = True Then
glTranslatef 0, -7, 0
Ty = Ty - 7
V = 0
    For i = 1 To UBound(a)
        If a(i) = "|" Then
            glTranslatef 0, 2, 0
            Ty = Ty + 2
        Else
        V = V + 1
            glTranslatef 0, HeightRolledMenu + 1, 0
            Ty = Ty + HeightRolledMenu + 1
            If MenuHits(KtMenu + V) Then
                glColor4f 1, 1, 1, 0.1
            Else
                glColor4f 1, 1, 1, 0.2
            End If
            glCallList 36
        End If
    Next
End If
End Sub

Public Sub GLCreateMenuSel(ByVal Idx As Byte, Optional ByRef Tx As GLdouble, Optional ByRef Ty As GLdouble, Optional Rolled As Boolean = False)

Dim a() As String, i As Integer
Dim V As Integer
If Menus(Idx).Items = "" Then Exit Sub
a = Split(Menus(Idx).Items, ";")

glLoadName Idx: glPushMatrix
glCallList 35
glPopMatrix

If Rolled = True Then
glTranslatef 0, -7, 0
Ty = Ty - 7
V = 0
    For i = 1 To UBound(a)
        If a(i) = "|" Then
            glTranslatef 0, 2, 0
            Ty = Ty + 2
        Else
            V = V + 1
            glTranslatef 0, HeightRolledMenu + 1, 0
            Ty = Ty + HeightRolledMenu + 1
            glLoadName KtMenu + V: glPushMatrix
            glCallList 36
            glPopMatrix
        End If
    Next
End If
End Sub

Public Sub GLMsgBox()
'1:ok
'2:cancel
'3:yes
'4:no
'5:default button1
'6:defaul button2
'7: Default button3
glLoadIdentity
Dim a() As String, i As Integer, SpBet As GLfloat 'spacebetwwen
a = Split(glSett, ";")
SpBet = 30

'carou mare alb
glTranslatef (MainW - WidthMessage) \ 2, (MainH - HeightMessage) \ 2, 0
glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha
glColor4f 1, 1, 1, 0.5
glCallList 46

glTranslatef SpBet, HeightMessage - HeightButton - 15, 0
glTranslatef ((WidthMessage - 2 * SpBet) - ((WidthButton + SpBet) * (UBound(a) + 1))) \ 2 + SpBet \ 2, 0, 0

'buttons
glBlendFunc sfSrcAlpha, dfOneMinusDstAlpha
glColor4f 1, 1, 1, 0

For i = 1 To UBound(a) + 1
    glCallList 45
    glTranslatef SpBet + WidthButton, 0, 0
Next


'glBlendFunc sfOneMinusDstAlpha, dfOneMinusSrcAlpha
'glColor4f 1, 1, 1, 1
'glTranslatef 3, 4, 0
'glCallList 47
End Sub

Private Function GetCuvFromNumb(Categ As Byte, ZeroSauUnu As Byte) As String
Select Case Categ
Case 1 'cul
    If ZeroSauUnu = 0 Then GetCuvFromNumb = "light": Exit Function
    If ZeroSauUnu = 1 Then GetCuvFromNumb = "dark": Exit Function
Case 2 'marime
    If ZeroSauUnu = 0 Then GetCuvFromNumb = "short": Exit Function
    If ZeroSauUnu = 1 Then GetCuvFromNumb = "tall": Exit Function
Case 3 'forma
    If ZeroSauUnu = 0 Then GetCuvFromNumb = "round": Exit Function
    If ZeroSauUnu = 1 Then GetCuvFromNumb = "square": Exit Function
Case 4 'solid
    If ZeroSauUnu = 0 Then GetCuvFromNumb = "hollow": Exit Function
    If ZeroSauUnu = 1 Then GetCuvFromNumb = "solid": Exit Function
End Select
End Function
