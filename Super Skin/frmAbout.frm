VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmAbout.frx":67E2
   ScaleHeight     =   6630
   ScaleMode       =   0  'User
   ScaleWidth      =   9500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrFadeIn 
      Interval        =   1
      Left            =   720
      Top             =   6240
   End
   Begin VB.Timer tmrFadeOut 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   6240
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   4485
      Left            =   4621
      TabIndex        =   1
      Top             =   1725
      Width           =   4650
      _cx             =   8202
      _cy             =   7911
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   "LR"
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin VB.PictureBox picSplash 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      Picture         =   "frmAbout.frx":C28E4
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem // --- Declare what we need to make the Form 'Always On Top'
Option Explicit
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 Rem // --- Dim what we need to Fade the Form In & Out.
 Dim Fade%
 Rem // --- I also use a Unload Ststus So If the Form isn't Faded out it
 Rem // --- Will Cancel the Unload, Then Fade Out, Then Unload.
 Dim UnloadStatus As Boolean


Private Sub Form_Activate()
Rem // --- Set this Form on top of all others.
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Form_Click()
Rem // --- When the Form is clicked unload the 'About Screen'
Unload Me
End Sub

Private Sub Form_Load()
Rem // --- Dim what we need to Subclass The Flash
Dim lhWnd As Long
Dim lRet As Long, lParam As Long

Rem // --- Skin the Form, to the Splashes Shape
Call createSkinnedForm(Me, picSplash)

Rem // --- Load the 'About Section' in the Flash
On Error Resume Next
Me.ShockwaveFlash1.Movie = App.Path & "\Config\About.DSTMP"

Rem // --- Fade in the Form.
Fade% = 0
UnloadStatus = False

MakeTransparent Me.hWnd, Fade%

Rem // --- Sub Class the Flash Object, so u cant Right Click on it.
lhWnd = Me.hWnd
lRet = EnumChildWindows(lhWnd, AddressOf EnumChildProc, lParam)

If FHW <> 0 Then
   glPrevWndProc = SubClass()
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rem // --- Enable the Main Form Again
Rem // --- Check if the Form is right to Be Unloaded, if it isn't then
Rem // --- Fade out the Form, then Unload It.
If UnloadStatus = False Then
 Cancel = 1
 Me.tmrFadeOut.Enabled = True
End If
frmMain.Enabled = True
End Sub

Private Sub tmrFadeIn_Timer()
Rem // --- Fade in the Form
On Error Resume Next
Fade% = Fade% + 10
MakeTransparent Me.hWnd, Fade%
If Fade% > 255 Then
 Me.tmrFadeIn.Enabled = False
End If
End Sub

Private Sub tmrFadeOut_Timer()
Rem // --- Fade Out the Form
On Error Resume Next
Fade% = Fade% - 10
MakeTransparent Me.hWnd, Fade%
If Fade% < 0 Then
 Me.tmrFadeOut.Enabled = False
 UnloadStatus = True
 Unload Me
End If
End Sub
