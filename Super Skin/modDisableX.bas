Attribute VB_Name = "modDisableX"
Rem // --- This mod disables the X on any Form
Rem // --- I have tryed to keep each 'Tool' in a Seperate MOD, so you dont have to Wade through Heaps Of Code.
Option Explicit

Rem // --- This is just calling a Function in a DLL file that VB cant Read.
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Const SC_CLOSE = &HF060&
    Public Const MF_BYCOMMAND = &H0&


