VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Processing..."
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrProgress 
      Interval        =   1
      Left            =   720
      Top             =   480
   End
   Begin ComctlLib.ProgressBar pbrProcess 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Please wait while Super Skin Processes the Commands You Specified."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem // --- Declare what we need to Make this form 'On Top' of all others
Option Explicit
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
Rem // --- Disable the "X" Button
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
Rem // --- Set this Form on top of all others
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub tmrProgress_Timer()
Rem // Make the Progress Bar Go Up.
Rem // It does this while the work is been done in / on another Form.
On Error Resume Next
DoEvents
Me.pbrProcess.Value = Me.pbrProcess.Value + 1
If Me.pbrProcess.Value > 99 Then Me.tmrProgress.Enabled = False: Unload frmProgress
End Sub
