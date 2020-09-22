VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmSplash.frx":67E2
   ScaleHeight     =   6630
   ScaleMode       =   0  'User
   ScaleWidth      =   9500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrLoader 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1680
      Top             =   5520
   End
   Begin VB.Timer tmrFadeOut 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   5520
   End
   Begin VB.Timer tmrFadeIn 
      Interval        =   1
      Left            =   2160
      Top             =   5520
   End
   Begin VB.PictureBox picSplash 
      BackColor       =   &H00FF00FF&
      Height          =   1215
      Left            =   1680
      Picture         =   "frmSplash.frx":C28E4
      ScaleHeight     =   1155
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Super Skin, Please Wait..."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   2040
      Width           =   3855
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem // --- Declare what we need to make the Form 'Always On Top'
Option Explicit
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 Rem // --- Dim what we need to Fade the Form In & Out.
 Dim Fade%
 Dim Fade2%
 Dim UnloadStatus As Boolean

Private Sub Form_Click()
Rem // --- In case the user dosnt want to wait 4 secconds, Click the
Rem // --- Form, and it will Advance to the Next Stage
Rem // --- This will Load the code in the 'tmrLoader_Timer' Sub
tmrLoader_Timer
End Sub

Private Sub Form_Load()
Rem // --- Skin the Form, to the Splashes Shape
Call createSkinnedForm(Me, picSplash)
Rem // --- Set this Form on top of all others.
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
Rem // --- Set the fade point
Fade% = 0
UnloadStatus = False
Rem // --- Fade to 0%
MakeTransparent Me.hWnd, Fade%
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rem // --- Check if the Form is right to Be Unloaded, if it isn't then
Rem // --- Fade out the Form, then Unload It.
If UnloadStatus = False Then
 Cancel = 1
 Me.tmrFadeOut.Enabled = True
End If
End Sub

Private Sub lblInfo_Click()
Rem // --- If the user gets impatient then When the Click the Form they can advance on.
Rem // --- This is Done by Calling the Timer SUB tht fades out The Form.
tmrLoader_Timer
End Sub

Private Sub tmrFadeIn_Timer()
Rem // --- Fade in the Form
On Error Resume Next
Fade% = Fade% + 5
MakeTransparent Me.hWnd, Fade%
If Fade% > 255 Then
 Me.tmrFadeIn.Enabled = False
 Me.tmrLoader.Enabled = True
End If
End Sub

Private Sub tmrFadeOut_Timer()
Rem // --- Fade Out the Form
On Error Resume Next
Fade% = Fade% - 10
Fade2% = Fade2% + 20
MakeTransparent Me.hWnd, Fade%
MakeTransparent frmMain.hWnd, Fade2%
If Fade% < 0 Then
 Me.tmrFadeOut.Enabled = False
 MakeTransparent frmMain.hWnd, 255
 MakeOpaque frmMain.hWnd
 UnloadStatus = True
 Unload Me
End If
End Sub

Private Sub tmrLoader_Timer()
Rem // --- The timer, is realy just to Show off the Splash,
Rem // --- it only takes about 3ms to Load the Main Form(s)...
On Error Resume Next
Me.tmrFadeIn.Enabled = False
DoEvents
Load frmMain
MakeTransparent frmMain.hWnd, 1
frmMain.Show
DoEvents
Unload Me
End Sub
