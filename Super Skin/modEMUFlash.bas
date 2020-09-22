Attribute VB_Name = "modEMUFlash"
Rem // --- I Didnt Write this MOD, do a Search on PSC for Ravindra Deuskar if you want a Full E.G.

'@@@@@@@@@ Developed by Ravindra Deuskar @@@@@@@@@@@@@@@@@@@@
Option Explicit
'###############
    'Subclass flash activex control
    'trap all messages pass to original window
    'procedure except right mouse messages
'###############

Public glPrevWndProc As Long
Public glPrevWndProc2 As Long
Public FHW As Long
Public FHW2 As Long
Public Const GWL_WNDPROC = (-4)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const WM_KEYDOWN = &H100
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_LBUTTONDBLCLK = &H201
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Sub UnSubClass()
Call SetWindowLong(FHW, GWL_WNDPROC, glPrevWndProc)
FHW = 0
End Sub

Public Sub UnSubClass2()
Call SetWindowLong(FHW, GWL_WNDPROC, glPrevWndProc)
FHW2 = 0
End Sub


Public Function MyWindowProc(ByVal HW As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_RBUTTONDOWN Then
             If frmAbout.Visible = True Then
              Unload frmAbout
              Debug.Print "unSubClassed About Form."
             End If
        Exit Function
    ElseIf uMsg = WM_RBUTTONUP Then
             If frmAbout.Visible = True Then
              Unload frmAbout
              Debug.Print "unSubClassed About Form."
             End If
        Exit Function
    ElseIf uMsg = WM_LBUTTONDBLCLK Then
             If frmAbout.Visible = True Then
              Unload frmAbout
              Debug.Print "unSubClassed About Form."
             End If
        Exit Function
    ElseIf uMsg = WM_KEYDOWN Then
    End If
    MyWindowProc = CallWindowProc(glPrevWndProc, HW, uMsg, wParam, lParam)
End Function

Public Function SubClass() As Long
    SubClass = SetWindowLong(FHW, GWL_WNDPROC, AddressOf MyWindowProc)
End Function

