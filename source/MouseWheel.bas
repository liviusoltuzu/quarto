Attribute VB_Name = "MouseWheel"
Option Explicit

'==============inside a MODULE
'************************************************************
'API
'************************************************************

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

'************************************************************
'Constants
'************************************************************

Public Const MK_CONTROL = &H8
Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_MBUTTON = &H10
Public Const MK_SHIFT = &H4
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A


'************************************************************
'Variables
'************************************************************

Private hControl As Long
Private lPrevWndProc As Long

'*************************************************************
'WindowProc
'*************************************************************

'zDelta: The value of the high-order word of wParam.
'Indicates the distance that the wheel is rotated, expressed in multiples or
'divisions of WHEEL_DELTA, which is 120. A positive value indicates that the
'wheel was rotated forward, away from the user; a negative value indicates
'that the wheel was rotated backward, toward the user.
Private Function WindowProc(ByVal lWnd As Long, ByVal lMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

    'Dim fwKeys As Long
    Dim zDelta As Long
    'Dim xPos As Long
    'Dim yPos As Long

    'Test if the message is WM_MOUSEWHEEL
    If lMsg = WM_MOUSEWHEEL Then
        'fwKeys = wParam And 65535
        zDelta = CLng(wParam \ 65536)
        'xPos = lParam And 65535
        'yPos = lParam / 65536
        'Call the Form1's Procedure to handle the MouseWheel event
        Stage2.MouseWheel zDelta
    End If
    'Sends message to previous procedure
    'This is VERY IMPORTANT!!!
    WindowProc = CallWindowProc(lPrevWndProc, lWnd, lMsg, wParam, lParam)
End Function

'*************************************************************
'Hook
'*************************************************************
Public Sub Hook(ByVal hControl_ As Long)
    hControl = hControl_
lPrevWndProc = SetWindowLong(hControl, GWL_WNDPROC, AddressOf WindowProc)
End Sub

'*************************************************************
'UnHook
'*************************************************************
Public Sub UnHook()
    Call SetWindowLong(hControl, GWL_WNDPROC, lPrevWndProc)
End Sub





