Attribute VB_Name = "MainAndOthers"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'Ortho variables/constants
Public Const WidthPickListButton = 30
Public Const HeightPickListButton = 70

Public Const HeightPLBTriangle = HeightPickListButton - 30
Public Const WidthPLBTriangle = WidthPickListButton - 20 'width picklist button triangle

Public WidthPickList As GLdouble
Public Const HeightPickList = 100

Public Const HWPiece = HeightPickList - 10

Public Const HWGivenPiece = 100 'height width /viewport dr-sus,dr-jos

Public Const HeightQuartoButton = 40
Public Const WidthQuartoButton = 70

Public WidthMenu As GLdouble
Public Const HeightMenu = 20

Public Const WidthMenuTitle = 50
Public Const HeightMenuTitle = HeightMenu - 2

Public Const WidthRolledMenu = 114
Public Const HeightRolledMenu = 25

                'Public Const HeightMsgBox = 50
                'Public Const WidthMsgBox = 150

Public Const WidthMessage = 350
Public Const HeightMessage = 100

Public Const WidthOptions = 400
Public Const HeightOptions = 300

Public Const HeightButton = 25
Public Const WidthButton = 80

Public LogF As Integer
Public DebugSes As Boolean
Public DebugP As Piesa 'piesa data

'altele
'Public dU As Double 'pt timer


Public Function ConvToBase(Number As Long, Base As Integer, AddZero As Integer) As String
Dim s As String, i As Integer, S2 As String
While Number <> 0
    s = s & Number Mod Base
    Number = Number \ Base
Wend
If Len(s) < AddZero Then
    s = s & String(AddZero - Len(s), "0")
End If
For i = Len(s) To 1 Step -1
    S2 = S2 & Mid(s, i, 1)
Next
ConvToBase = S2
End Function

Public Function ReConvToBase(NumberB As String, Base As Integer) As Long
Dim l As Integer, r As Long, i As Long
l = Len(NumberB)
For i = 1 To l
    r = r + Mid(NumberB, i, 1) * Base ^ (l - i)
Next
ReConvToBase = r
End Function

Public Function Modulo(Sign As String, ModuloX As Long, Operator1 As Long, Operator2 As Long)
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
Public Function GetTime() As String
GetTime = "[" & Now & "]"
End Function
Sub Main()

Dim fso As New FileSystemObject, z
Set z = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName(App.Path)))
LogF = 0
If z.DriveType = 2 Or z.DriveType = 3 Or z.DriveType = 1 Then
    LogF = FreeFile
    Open App.Path & "\log.txt" For Output As #LogF
End If
If LogF <> 0 Then Print #LogF, GetTime & "Quarto! log"


If App.PrevInstance Then If LogF <> 0 Then Print #LogF, "Quarto! is running -> Exit": End
If VerifResF Then
    If LogF <> 0 Then Print #LogF, GetTime & "resource modified. exit"
    MsgBox "Resource modified." & vbCrLf & vbCrLf & "The program will close.", vbOKOnly + vbCritical
    End
End If
'dU = Timer

ZStart.Show
DoEvents
ReadRes
    CreateGL Stage2.Stage
    VerifErr True
    If LogF <> 0 Then Print #LogF, GetTime & "opengl ok"
    ZStart.lblDoing = "rendering..."
    ZStart.lblDoing.Refresh
    'CreateGLLists
    Erase Objects.objs
    Erase Objects.Verticies
    Erase Objects.Normals
    Erase Objects.TCoords
    'oneRendered = True

'read preferences

Unload ZStart
'WriteSett
If LogF <> 0 Then Print #LogF, GetTime & "reading quarto.sett..."
ReadSett

'ZStart.lblDoing.Refresh
CineMuta = True

NewGame


VerifErr
'frmQuarto.Show
'frmOptions.Show
Stage2.Show
Stage2.Timer1_Timer

End Sub

Public Sub VerifErr(Optional OffAll As Boolean = False)
If ErrPrg Or ErrNumb <> 0 Then
    MsgBox LoadResString(ErrNumb), vbOKOnly + vbCritical, "Quarto!"
    ErrPrg = False
    If LogF <> 0 Then Print #LogF, GetTime & "Error " & ErrNumb & ": " & LoadResString(ErrNumb)
    ErrNumb = 0
    If OffAll Then If LogF <> 0 Then Print #LogF, GetTime & "Exit": End
End If
End Sub

Private Function VerifResF() As Boolean
On Error GoTo err:
Dim i As Long
For i = 101 To 108
    LoadResPicture i, 0
Next
For i = 101 To 106
    LoadResString i
Next
err:
If err.Number <> 0 Then
    VerifResF = True
End If
End Function

