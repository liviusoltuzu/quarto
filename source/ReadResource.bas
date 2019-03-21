Attribute VB_Name = "WRFiles"
Option Explicit
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Public FilesPath As String

'''''''''''''''''
'               Settings
'StilDeJoc 1 or 2 or 3
'SingleOption 1 and/or 2 and/or 3 and/or 4
'Window size: Height Width
'Who Starts Alternative=0 player =1 computer=2
'Zoom
''''''''''''
Type TSettings
    sStilDeJoc As Byte
    sSingleOption(1 To 4) As Boolean
    sWindowHeight As Long
    sWindowWidth As Long
    sWindowState As Byte
    sWhoStarts As Byte
    sUseDblClkFnc As Boolean
    sShowHtext As Boolean
    sZoom As Integer
End Type
Dim AllSett As TSettings
'variabile pt intreruperea executiei la eroare
Public ErrPrg As Boolean
Public ErrNumb As Long


'Private Function GetTempDir() As String
'Dim S As String
'S = String(200, Chr$(0))
'GetTempPath 200, S
'S = Left$(S, InStr(S, Chr$(0)) - 1)
'GetTempDir = S
'End Function

Public Sub ReadRes()
ReadObj App.Path & "\objs.3df"
'Stage2.Stage.MouseIcon = LoadResPicture(101, 1)
End Sub

Public Function ReadSett() As Boolean
ReadSett = False
On Error GoTo err:

'in caz ca da eroare
AllSett.sSingleOption(1) = True
AllSett.sSingleOption(2) = True
AllSett.sSingleOption(3) = True
AllSett.sSingleOption(4) = True
AllSett.sStilDeJoc = 2
AllSett.sWhoStarts = 0
AllSett.sUseDblClkFnc = False
AllSett.sShowHtext = True
AllSett.sWindowState = 0
AllSett.sWindowHeight = 8340
AllSett.sWindowWidth = 10500
AllSett.sZoom = 60


Dim FileNumber As Integer
FileNumber = FreeFile
Dim fso As New FileSystemObject
If fso.FileExists(App.Path & "\quarto.sett") Then
    Open App.Path & "\quarto.sett" For Binary Access Read Lock Read Write As #FileNumber
    Get #FileNumber, , AllSett
    Close #FileNumber
Else
    ErrNumb = 106
    MsgBox LoadResString(106), vbOKOnly + vbCritical, "Quarto!"
    If LogF <> 0 Then Print #LogF, GetTime & "Error " & ErrNumb & ": " & LoadResString(ErrNumb)
    ErrNumb = 0
End If


    SingleOption(1) = AllSett.sSingleOption(1)
    SingleOption(2) = AllSett.sSingleOption(2)
    SingleOption(3) = AllSett.sSingleOption(3)
    SingleOption(4) = AllSett.sSingleOption(4)
    StilDeJoc = AllSett.sStilDeJoc
    WhoStarts = AllSett.sWhoStarts
    Stage2.Height = AllSett.sWindowHeight
    Stage2.Width = AllSett.sWindowWidth
    Stage2.scrZoom.Value = AllSett.sZoom
    Stage2.WindowState = AllSett.sWindowState
    UseDblClkFnc = AllSett.sUseDblClkFnc
    ShowHtext = AllSett.sShowHtext


'Dim C As Byte
'For C = 1 To 4
'frmOptions.chkCaract(C).Value = Abs(CLng(SingleOption(C)))
'Next
'StilDeJoc = 3
'frmOptions.optStilDeJoc(StilDeJoc).Value = True
'frmOptions.optWhoStarts(WhoStarts).Value = True
ReadSett = True
If LogF <> 0 Then Print #LogF, GetTime & "done."

err:
If err.Number <> 0 Then
    'ErrPrg = True
    ErrNumb = 106
    MsgBox LoadResString(ErrNumb), vbOKOnly + vbCritical, "Quarto!"
    If LogF <> 0 Then Print #LogF, GetTime & "Error " & ErrNumb & ": " & LoadResString(ErrNumb)
    ErrNumb = 0
    
    SingleOption(1) = True
    SingleOption(2) = True
    SingleOption(3) = True
    SingleOption(4) = True
    StilDeJoc = 2
    WhoStarts = 0
    Stage2.Height = 8340
    Stage2.Width = 10500
    Stage2.scrZoom.Value = 60
    Stage2.WindowState = 0
    UseDblClkFnc = False
    ShowHtext = True
    'End
End If
End Function
Public Function WriteSett() As Boolean
If LogF <> 0 Then Print #LogF, GetTime & "save options"
WriteSett = False
On Error GoTo err:

AllSett.sSingleOption(1) = SingleOption(1)
AllSett.sSingleOption(2) = SingleOption(2)
AllSett.sSingleOption(3) = SingleOption(3)
AllSett.sSingleOption(4) = SingleOption(4)
AllSett.sStilDeJoc = StilDeJoc
AllSett.sWhoStarts = WhoStarts
AllSett.sWindowHeight = Stage2.Height
AllSett.sWindowWidth = Stage2.Width
AllSett.sWindowState = Stage2.WindowState
AllSett.sUseDblClkFnc = UseDblClkFnc
AllSett.sZoom = Stage2.scrZoom.Value
AllSett.sShowHtext = ShowHtext


Dim FileNumber As Integer
FileNumber = FreeFile
'Kill App.Path & "\quarto.sett"

Open App.Path & "\quarto.sett" For Binary Access Write Lock Read Write As #FileNumber
Put #FileNumber, , AllSett
Close #FileNumber

WriteSett = True
If LogF <> 0 Then Print #LogF, GetTime & "done."
Exit Function

err:
If err.Number <> 0 Then
    ErrPrg = True
    ErrNumb = 105
    MsgBox LoadResString(ErrNumb), vbOKOnly + vbCritical, "Quarto!"
    If LogF <> 0 Then Print #LogF, GetTime & "Error " & ErrNumb & ": " & LoadResString(ErrNumb)
    ErrNumb = 0
    ErrPrg = False
    'End
End If
End Function
