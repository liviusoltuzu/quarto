VERSION 5.00
Begin VB.Form Stage2 
   BackColor       =   &H00CCCCCC&
   Caption         =   "Quarto!"
   ClientHeight    =   5430
   ClientLeft      =   2745
   ClientTop       =   555
   ClientWidth     =   9345
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   623
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bDebug 
      Default         =   -1  'True
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox tDebug 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   3360
   End
   Begin VB.VScrollBar scrZoom 
      Height          =   1215
      Left            =   0
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   3840
      Value           =   1
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   240
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Stage 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   120
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   470
      TabIndex        =   0
      Top             =   0
      Width           =   7080
   End
End
Attribute VB_Name = "Stage2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sX As Integer, sY As Integer
Private ResX As Single, ResY As Single  'resize x/y
Dim TwipsPerPixelX As Single
Dim TwipsPerPixelY As Single

Private mX As Long, mY As Long 'move window
Private AllLoaded As Boolean


Private Sub cmdEnd_Click()
glDeleteLists 1, FontBase - 1 '1 to 26 tabla,etc: 27 to 35 ortho si celelalte
glDeleteLists FontBase, 96
glDeleteLists FontBaseTip1, 96
glDeleteTextures 7, Texture(0)
Erase Texture()
'WriteSett
UnHook
If LogF <> 0 Then Print #LogF, GetTime & "Exit"
Close #LogF
End
End Sub



Private Sub bDebug_Click()
'Piese(13).ESett = 0
'KTpiese = CreateAvPieces

Select Case LCase(Mid(tDebug, 1, 1))
Case "d"
    Piese(Mid(tDebug, 2)).ESett = 1
    KTpiese = CreateAvPieces
    DebugP.ACuloare = Mid(ConvToBase(Mid(tDebug, 2) - 1, 2, 4), 1, 1)
    DebugP.BMarime = Mid(ConvToBase(Mid(tDebug, 2) - 1, 2, 4), 2, 1)
    DebugP.CForma = Mid(ConvToBase(Mid(tDebug, 2) - 1, 2, 4), 3, 1)
    DebugP.DSoliditate = Mid(ConvToBase(Mid(tDebug, 2) - 1, 2, 4), 4, 1)
    DebugP.ESett = 1
Case "p"
    Tabla(Mid(tDebug, 2, 1), Mid(tDebug, 3, 1)) = DebugP
    KTpiese = CreateAvPieces
End Select
bDebug.Caption = tDebug
tDebug = ""
tDebug.SetFocus
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim PosTablaT As Byte 'o copie/temp
'Select Case KeyControl
'    Case 2
'        If CineMutaAcum Or Remiza Or Ilose Or Youlose Or ShowOptions Or KeepMenu Then Exit Sub
'        PosTablaT = PosTabla
'        Select Case KeyCode
'        Case vbKeyLeft
'            While Tabla(PosTablaT \ 10, PosTablaT Mod 10).ESett = 1 And (PosTablaT \ 10) > 1
'                PosTablaT = PosTablaT - 10
'            Wend
'        Case vbKeyRight
'            While Tabla(PosTablaT \ 10, PosTablaT Mod 10).ESett = 1 And (PosTablaT \ 10) < 3
'                PosTablaT = PosTablaT + 10
'            Wend
'            'postablat=
'        Case vbKeyUp
'            While Tabla(PosTablaT \ 10, PosTablaT Mod 10).ESett = 1 And (PosTablaT Mod 10) > 1
'                PosTablaT = PosTablaT - 1
'            Wend
'        Case vbKeyDown
'            While Tabla(PosTablaT \ 10, PosTablaT Mod 10).ESett = 1 And (PosTablaT Mod 10) < 3
'                PosTablaT = PosTablaT + 1
'            Wend
'        End Select
'        PosTabla = PosTablaT
'End Select
'
'End Sub

'Private Function ConvSirToMat(Nr As Byte, Tip As Byte) As Byte
''un element din matricea tabla se poate nota fie cu coordonatele xy fie cu pozitia
''lui incepand de la stanga sus(de la 1 la 4*4=16)
''functia converteste
'Select Case Tip
'Case 1
'
'
'End Function

Private Sub Form_Load()
If DebugSes Then tDebug.Visible = True: bDebug.Visible = True
TwipsPerPixelX = Me.ScaleX(1, vbPixels, vbTwips)
TwipsPerPixelY = Me.ScaleY(1, vbPixels, vbTwips)
'Stage2.Width = 10500
'Stage2.Height = 8340
ZoomRate = 1000
'scrZoom.Value = 60
oneRendered = False
InitVar
AllLoaded = True
'fps2 = 20

Hook Stage2.hWnd
End Sub


Private Sub InitVar()
If Not Stage2.WindowState = vbMinimized Then

Stage.Left = 0
'Stage.Top = HeightMenu
Stage.Width = Stage2.ScaleWidth
Stage.Height = Stage2.ScaleHeight '- 20 '- 88
MainH = Stage2.Stage.ScaleHeight
If MainH <= 0 Then MainH = 1
MainW = Stage2.Stage.ScaleWidth

MHMinHOp2 = HeightMenu + HeightPickList + ((MainH - HeightMenu - HeightPickList) - HeightOptions) \ 2
'MHMinHOp2 = (MainH - HeightOptions) \ 2
MWMinWOp2 = (MainW - WidthOptions) \ 2

URW = HWGivenPiece '+ 25
URH = HWGivenPiece '+ 25
WidthPickList = MainW ' - 10
WidthMenu = MainW

glDeleteLists 27, 1
GLCreateShape WidthPickList, HeightPickList, 27, False 'picklist
glDeleteLists 31, 1
GLCreateShape WidthMenu, HeightMenu, 31, False 'menu

'If oneRendered Then
'DoEvents
'SwapBuffers Stage.hDC
'DoEvents
'DrawGLScene
'DoEvents
'End If

'glDeleteLists 39, 1
'GLCreateShape WidthMenu, HeightMenu - 1, 39, False

End If

End Sub

Private Sub Form_Resize()
InitVar
'Shape1.Width = Stage2.Width / TwipsPerPixelX
'Stage.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdEnd_Click
End Sub

Private Sub scrZoom_Scroll()
DoEvents
End Sub

Private Sub Stage_Click()
On Error Resume Next
'If ShowAbout Then
'ShowAbout = False
'Exit Sub
'End If
Dim j As Byte, Rez As Byte, j1 As Byte, j2 As Byte 'rez=variabila oarecare ce retine pozitia selectata/cu mouseul din hitslist
'If ShowMessage Then
'    For j = 1 To MsgBoxHitsConst
'        If MsgBoxHits(j) Then
'            Select Case TipMesaj
'                Case 1
'                    TipMesaj = 0
'            End Select
'            ShowMessage = False
'        Exit For
'        End If
'    Next
'Exit Sub
'End If


If ShowOptions Then
    For j = 1 To OptionsHitsConst
        If OptionsHits(j) Then
            Select Case j
                Case 1
                    WhoStarts = 0
                Case 2
                    WhoStarts = 1
                Case 3
                    WhoStarts = 2
                Case 4
                    StilDeJoc = 1
                Case 5
                    StilDeJoc = 2
                    For j1 = 1 To 4
                        SingleOption(j1) = True
                    Next
                Case 6
                    StilDeJoc = 3
                    For j1 = 1 To 4
                        SingleOption(j1) = True
                    Next
                Case 7
                    If StilDeJoc = 1 Then SingleOption(1) = Not SingleOption(1)
                Case 8
                    If StilDeJoc = 1 Then SingleOption(2) = Not SingleOption(2)
                Case 9
                    If StilDeJoc = 1 Then SingleOption(3) = Not SingleOption(3)
                Case 10
                    If StilDeJoc = 1 Then SingleOption(4) = Not SingleOption(4)
                Case 11 'default
                    WhoStarts = 0
                    StilDeJoc = 2
                    For j1 = 1 To 4
                        SingleOption(j1) = True
                    Next
                    UseDblClkFnc = False
                    ShowHtext = True
                'Case 12 'save
                '    WriteSett
                '    ShowMessage = True
                '    TipMesaj = 1
                Case 13 'close
                    If (StilDeJocTemp <> StilDeJoc) Or (WhoStartsTemp <> WhoStarts) Or _
                        (SingleOptionTemp(1) <> SingleOption(1)) Or _
                        (SingleOptionTemp(2) <> SingleOption(2)) Or _
                        (SingleOptionTemp(3) <> SingleOption(3)) Or _
                        (SingleOptionTemp(4) <> SingleOption(4)) Then
                    If LogF <> 0 Then Print #LogF, GetTime & "Options changed: First move(" & WhoStarts & "); Level(" & _
                        StilDeJoc & ", LD:" & SingleOption(1) & " ST:" & SingleOption(2) & " RS:" & _
                        SingleOption(3) & " HS:" & SingleOption(4) & "); Use 2clicks(" & UseDblClkFnc & _
                        "); Show instructions(" & ShowHtext & ")"

                        NewGame
                    End If
                    ShowOptions = False
                    
                 Case 14 'use dblclkfunc
                    UseDblClkFnc = Not UseDblClkFnc
                    For j1 = 1 To 4
                        For j2 = 1 To 4
                            TblDblClk(j1, j2) = False
                        Next
                    Next
                    For j1 = 1 To 16
                        LstDblClk(j1) = False
                    Next
                 Case 15 '[your turn] etc
                    If ShowHtext Then
                    ShowHtext = False
                    Else
                    ShowHtext = True
                    End If
                    'ShowHtext = Not ShowHtext
            End Select
         
            Exit For
        End If
    Next
Exit Sub
End If
            

'click pe rolled menu
If KeepMenu And Not ShowOptions Then
For j = 1 To KtMenu
    If MenuHits(j) And Menus(j).Rolled Then
        Select Case j
        Case 1 'File
            For j1 = KtMenu + 1 To KtMenu + KtRolledMenu
                If MenuHits(j1) Then
                    Select Case j1
                    Case KtMenu + 1 'New Game
                        NewGame
                    Case KtMenu + 2 'Exit
                         Unload Stage2
                    End Select
                    Exit For
                End If
            Next
        Case 3 'options
            For j1 = KtMenu + 1 To KtMenu + KtRolledMenu
                If MenuHits(j1) Then
                    Select Case j1
                        Case KtMenu + 1 'Gameplay
                            WhoStartsTemp = WhoStarts
                            StilDeJocTemp = StilDeJoc
                            SingleOptionTemp(1) = SingleOption(1)
                            SingleOptionTemp(2) = SingleOption(2)
                            SingleOptionTemp(3) = SingleOption(4)
                            SingleOptionTemp(4) = SingleOption(4)
                            ShowOptions = True
                            'ShowToolbars = False
                        'Case KtMenu + 2 'toolbars
                            'ShowOptions = True
                            'ShowToolbars = True
                        Case KtMenu + 2
                            WriteSett
                            'ShowMessage = True
                            'TipMesaj = 1
                    End Select
                    Exit For
                End If
            Next
        Case 2 'pieces
            For j1 = KtMenu + 1 To KtMenu + KtRolledMenu
                If MenuHits(j1) Then
                    If (j1 = KtMenu + 1) Or (j1 = KtMenu + 4) Or (j1 = KtMenu + 7) Or (j1 = KtMenu + 10) Then
                        MenuPieces(j1 + 1 - KtMenu) = True
                        MenuPieces(j1 + 2 - KtMenu) = True
                        MenuPieces(j1 - KtMenu) = True
                    Else
                        Select Case j1
                        Case KtMenu + 2
                            MenuPieces(1) = False
                            MenuPieces(3) = False
                        Case KtMenu + 3
                            MenuPieces(1) = False
                            MenuPieces(2) = False
                        Case KtMenu + 5
                            MenuPieces(4) = False
                            MenuPieces(6) = False
                        Case KtMenu + 6
                            MenuPieces(4) = False
                            MenuPieces(5) = False
                        Case KtMenu + 8
                            MenuPieces(7) = False
                            MenuPieces(9) = False
                        Case KtMenu + 9
                            MenuPieces(7) = False
                            MenuPieces(8) = False
                        Case KtMenu + 11
                            MenuPieces(10) = False
                            MenuPieces(12) = False
                        Case KtMenu + 12
                            MenuPieces(10) = False
                            MenuPieces(11) = False
                        Case KtMenu + 13
                            For j2 = 1 To 12
                                MenuPieces(j2) = True
                            Next
                        End Select
                        MenuPieces(j1 - KtMenu) = True
                    End If
                    
                    KTpiese = CreateAvPieces
                    If LogF <> 0 Then Print #LogF, GetTime & "Pieces set to: " & j1 - KtMenu
                    Exit For
                End If
            Next
        Case 4 'Help
            For j1 = KtMenu + 1 To KtMenu + KtRolledMenu
                If MenuHits(j1) Then
                    Select Case j1
                        Case KtMenu + 1 'help
                            Dim shex
                            If LogF <> 0 Then Print #LogF, GetTime & "Help: Show contents"
                            shex = ShellExecute(Stage2.hWnd, "open", App.Path & "\quarto_help.chm", vbNullString, vbNullString, 0)
                            If shex = 2 Then
                                MsgBox "File 'Quarto_help.chm' not found!", vbOKOnly + vbExclamation
                            ElseIf shex <= 32 Then
                                MsgBox "Error: " & shex & vbCrLf & vbCrLf & "Unable to load the help file.", vbExclamation + vbOKOnly
                            End If
                            
                        'Case KtMenu + 2 'Credits
                        '    'ShowCredits = True
                        'Case KtMenu + 3 'about
                        '    ShowAbout = True
                        '    AboutDelay = Timer + 1
                    End Select
                End If
            Next
        End Select
        Exit For
    End If
Next
End If


If KeepMenu Then
    KeepMenu = False
    For j = 1 To KtMenu
        MenuHits(j) = False
        Menus(j).Rolled = False
    Next
    Exit Sub
End If
'On Error Resume Next
If MouseY > HeightMenu Then
    Rez = 0
    For j = 1 To HitsListConst
        If HitsList(j) Then
            Rez = j
            Exit For
        End If
    Next
    If Rez <= 16 And Rez >= 1 And (Not Ilose And Not Youlose And Not Remiza) Then
        Rez = CByte(Mid(Lists(Rez + 10).ObjName, 2)) 'tabla cu pozitii
        If (TblDblClk(Rez \ 10, Rez Mod 10) And UseDblClkFnc) Or Not UseDblClkFnc Then
            Piese(PiesaForU).ESett = 1
            If Tabla(Rez \ 10, Rez Mod 10).ESett = 1 Then Exit Sub
            Tabla(Rez \ 10, Rez Mod 10) = Piese(PiesaForU)
            If LogF <> 0 Then Print #LogF, GetTime & "player on " & Rez
            If CalculeazaComb(Rez \ 10, Rez Mod 10).Victory <> 0 Then
                CalculeazaComb Rez \ 10, Rez Mod 10, True
                Ilose = True
                If LogF <> 0 Then Print #LogF, GetTime & "You win!"
                KTpiese = CreateAvPieces: Exit Sub
            End If
            TbSaDeaPiesa = True
            For j = 1 To 12
                MenuPieces(j) = True
            Next
            KTpiese = CreateAvPieces
            If KTpiese = 0 And Not Ilose And Not Youlose Then
                Remiza = True
                If LogF <> 0 Then Print #LogF, GetTime & "Draw!"
            End If
            If UseDblClkFnc Then
                For j = 1 To 4
                    For j1 = 1 To 4
                        TblDblClk(j, j1) = False
                    Next
                Next
            End If
        ElseIf UseDblClkFnc Then
            For j = 1 To 4
                For j1 = 1 To 4
                    TblDblClk(j, j1) = False
                Next
            Next
            TblDblClk(Rez \ 10, Rez Mod 10) = True
        End If
    ElseIf Rez = 21 Then 'right button scroll picklist
        If StartView + NrSqr <= KTpiese Then
            StartView = StartView + 1
        End If
    ElseIf Rez = 20 Then 'left button scroll picklist
        If StartView > 1 Then
            StartView = StartView - 1
        End If
    ElseIf Rez >= 22 And (Rez <= (NrSqr + 22 - 1)) And TbSaDeaPiesa And (Not Ilose And Not Youlose And Not Remiza) And _
    MouseY < (HeightMenu + HeightPickList) Then  'lista
        If (LstDblClk(AvPieces(Rez - 22 + StartView)) And UseDblClkFnc) Or Not UseDblClkFnc Then
            DelayPutGivenP = Timer + 0.6
            PiesaForC = AvPieces(Rez - 22 + StartView)
            If LogF <> 0 Then Print #LogF, GetTime & "player gives " & PiesaForC
            Piese(PiesaForC).ESett = 1
            CineMutaAcum = True
            KTpiese = CreateAvPieces
            TbSaDeaPiesa = False
            If UseDblClkFnc Then
                For j = 1 To 16
                    LstDblClk(j) = False
                Next
            End If
        ElseIf UseDblClkFnc Then
            For j = 1 To 16
                LstDblClk(j) = False
            Next
            LstDblClk(AvPieces(Rez - 22 + StartView)) = True
        End If
    End If
Else
    For j = 1 To KtMenu
        If MenuHits(j) Then
            Menus(j).Rolled = True
            KeepMenu = True
            Exit For
        End If
    Next
End If
End Sub

Private Sub Stage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'If Button = vbRightButton Then
'    'frmOptions.Visible = True
'End If
If Not KeepMenu Then
sX = X
sY = Y
End If
'If (X > ((Me.Width / TwipsPerPixelX) - 15)) And (Y > ((Stage.Height / TwipsPerPixelY) - 35)) And Button = 1 Then
'    ResX = X
'    ResY = Y
'End If
'If Button = vbLeftButton And Y < HeightMenu And X > 400 Then
'    mX = X
'    mY = Y
'End If

End Sub


Public Sub Stage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

'If Button = vbLeftButton And Y < HeightMenu And X > 400 Then
'    Me.Left = Me.Left - (mX - X)
'    Me.Top = Me.Top - (mY - Y)
'End If

'If (Format$(ResX) <> 0 Or Format$(ResY) <> 0) And (Button = 1) And _
'(Me.Width / TwipsPerPixelX >= 400) And (Me.Height / TwipsPerPixelY >= 500) Then
''(X > ((Me.Width / TwipsPerPixelX) - 15)) And (Y > ((Stage.Height / TwipsPerPixelY) - 35)) Then
'    '(X > (Me.Width / TwipsPerPixelX) - 15) And (Y > (Me.Height / TwipsPerPixelY) - 35) Then
'    Me.Width = Me.Width + (X - ResX) * TwipsPerPixelX
'    Me.Height = Me.Height + (Y - ResY) * TwipsPerPixelY
'    ResX = X
'    ResY = Y
'    If (X < ((Me.Width / TwipsPerPixelX) - 15)) Or (Y < ((Stage.Height / TwipsPerPixelY) - 35)) Then
'        ResX = 0
'        ResY = 0
'    End If
'End If
If Me.Width / TwipsPerPixelX <= 700 Then
    Me.Width = 700 * TwipsPerPixelX
End If
If Me.Height / TwipsPerPixelY <= 556 Then
    Me.Height = 556 * TwipsPerPixelY
End If

MouseX = CSng(X)
MouseY = CSng(Y)

If MouseY < HeightMenu + HeightPickList Then Exit Sub

If ResX <> 0 Or ResY <> 0 Then sX = 0: sY = 0

If Button = vbLeftButton And sY <> 0 And sX <> 0 And Not ShowOptions Then
zRot = zRot + (X - sX) / 2
xRot = xRot + (Y - sY) / 2
sY = Y
sX = X
ElseIf Button = vbRightButton And sX <> 0 And ResX <> 0 And ResY <> 0 And Not ShowOptions Then
yRot = yRot + (X - sX) / 2
sX = X
End If
xRot = Sgn(xRot) * Modulo("+", 360, Abs(CLng(xRot)), 0) '& Format(xRot, ".00")
yRot = Sgn(yRot) * Modulo("+", 360, Abs(CLng(yRot)), 0)  '& Format(yRot, "0.00")
zRot = Sgn(zRot) * Modulo("+", 360, Abs(CLng(zRot)), 0) '& Format(zRot, "0.00")

'If (MouseX > MainW - 20) And (MainH > MainH - 20) Then
'    Stage2.MousePointer = 8
'End If
End Sub

Private Sub Stage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
sX = 0
sY = 0
'ResX = 0
'ResY = 0

'mX = 0
'mY = 0
End Sub


Private Function Modulo(Sign As String, ModuloX As Long, Operator1 As Long, Operator2 As Long)
Dim ModuloF As Long 'modulo final
Dim O1 As Long, O2 As Long

O1 = Operator1
O2 = Operator2

If O1 >= ModuloX Then
     O1 = O1 - (O1 \ ModuloX) * ModuloX
End If

If O2 >= ModuloX Then
    O2 = O2 - (O1 \ ModuloX) * ModuloX
End If

Select Case Sign
Case "+"
    ModuloF = O1 + O2
Case "-"
    If O1 < O2 Then O1 = O1 + ModuloX
    ModuloF = O1 - O2
Case "*"
    ModuloF = O1 * O2 - (O1 * O2 \ ModuloX) * ModuloX
End Select

Modulo = ModuloF
End Function

Private Sub scrZoom_Change()
Zoom = -ZoomRate / 16 / scrZoom.Value
DoEvents
End Sub

Private Sub Text4_Change()
'If Text4 = "" Then Text4 = 0
End Sub

Private Sub Text5_Change()
'If Text5 = "" Then Text5 = 0

End Sub

Public Sub Timer1_Timer()
If AllLoaded Then
    Timer1.Enabled = False
    Render Stage
    'AllLoaded = False
End If
If ErrPrg Or (ErrNumb <> 0) Then
    MsgBox LoadResString(ErrNumb), vbOKOnly + vbCritical, "Quarto!"
    End
End If

DoEvents
End Sub

Public Sub MouseWheel(ByVal zDelta As Long)
If zDelta <> 0 Then
    If zDelta < 0 And (scrZoom.Value + (Abs(zDelta) / 120) * 5 <= 100) Then
        scrZoom.Value = scrZoom.Value + (Abs(zDelta) / 120) * 5
    ElseIf zDelta > 0 And (scrZoom.Value - (Abs(zDelta) / 120) * 5 >= 1) Then
        scrZoom.Value = scrZoom.Value - (Abs(zDelta) / 120) * 5
    End If
End If
End Sub


'Private Sub optStilDeJoc_Click(Index As Integer)
'Dim i As Byte
'If Index = 1 Then
'    For i = 1 To 4
'        chkCaract(i).Enabled = True
'        chkCaract(i).Value = 1
'    Next
'Else
'    For i = 1 To 4
'        chkCaract(i).Enabled = False
'        chkCaract(i).Value = 1
'    Next
'End If
'If optStilDeJoc(Index).Value Then
'    StilDeJoc = Index
'End If
'End Sub
'
'Private Sub optWhoStarts_Click(Index As Integer)
'If optWhoStarts(Index).Value Then
'    WhoStarts = Index
'End If
'End Sub
