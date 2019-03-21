VERSION 5.00
Begin VB.Form ZStart 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Quarto!"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   411
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2385
      ScaleWidth      =   6105
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.Label lblDoing 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   1920
         Width           =   6135
      End
   End
End
Attribute VB_Name = "ZStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
p1.Picture = LoadResPicture(108, 0)
If LogF <> 0 Then Print #LogF, GetTime & "init picture loaded"
End Sub

