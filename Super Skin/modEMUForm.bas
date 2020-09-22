Attribute VB_Name = "modEMUForm"
Rem // --- I Didnt Write this MOD, do a Search on PSC for Ravindra Deuskar if you want a Full E.G.

'@@@@@@@@@ Developed by Ravindra Deuskar @@@@@@@@@@@@@@@@@@@@
Option Explicit
'###############
    'Enum all child windows of Form1
    'steal hwnd of flash activex control for subsaclass
'###############
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long


Function EnumChildProc(ByVal lhWnd As Long, ByVal lParam As Long) As Long
Dim RetVal As Long
Dim WinClassBuf As String * 255
Dim WinClass As String

    RetVal = GetClassName(lhWnd, WinClassBuf, 255)
    
    If (InStr(WinClassBuf, Chr(0)) > 0) Then
        WinClass = Left(WinClassBuf, InStr(WinClassBuf, Chr(0)) - 1)
    End If
        
    If WinClass = "MacromediaFlashPlayerActiveX" Then
       FHW = lhWnd
       EnumChildProc = False
    ElseIf Left(WinClass, 4) = "ATL:" Then
        FHW = lhWnd
        EnumChildProc = False
    Else
        EnumChildProc = True
    End If
End Function
