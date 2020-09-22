VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUnSkinAPP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UnSkin Application"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "frmUnSkinAPP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   3240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Output Options"
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   870
      Width           =   4815
      Begin VB.CheckBox chkBackup 
         Appearance      =   0  'Flat
         Caption         =   "Create Backup of Application"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   300
      Left            =   4440
      TabIndex        =   4
      Top             =   360
      Width           =   300
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Application To UnSkin"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtPathName 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4120
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   $"frmUnSkinAPP.frx":67E2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   4575
   End
End
Attribute VB_Name = "frmUnSkinAPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem // --- Dim what we will use in this Form
Option Explicit
 Dim EXE As Boolean
 Dim sTempName As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Rem // --- logval is just used for writing the log
Dim logVal As Variant

Rem // ---------------------------------------------------------------------------------------------------------------------------------------------------

If Me.txtPathName.Text = "" Then MsgBox "Please select an Application To UnSkin.", vbExclamation, "Super Skin": Exit Sub

Rem // ---------------------------------------------------------------------------------------------------------------------------------------------------

On Error GoTo ErrorCode

Me.Hide
frmProgress.Show

Rem // --- If the user has requested a Backup then Backup the File First.

DoEvents
If Me.chkBackup.Value = 1 Then
 DoEvents
 FileCopy Me.txtPathName.Text, Me.txtPathName.Text & " (Backed Up).exe"
 DoEvents
 frmProgress.pbrProcess.Value = 0
End If

DoEvents
frmProgress.pbrProcess.Value = 0

Rem // --- Remove (KILL) the Manifest, this isnt a problem beacuse you can use ANY
Rem // --- Manifest by this Program to Skin ANY Application, even a C++, n Delphi APP!!!!
DoEvents
If EXE = True Then
 Kill Me.txtPathName.Text & ".manifest"
Else
 Kill Me.txtPathName.Text
End If
 
Rem // --- Add the UnSkinned App To The Log.
 Set logVal = frmMain.lvSkined.ListItems.Add(, , sTempName)
  logVal.SubItems(1) = "UnSkin - Direct IDE Manifest"
  If Me.chkBackup.Value = 1 Then
   logVal.SubItems(2) = "Yes"
  Else
   logVal.SubItems(2) = "No"
  End If
  logVal.SubItems(3) = Me.txtPathName.Text
  logVal.SubItems(4) = Date & " - " & Time

Unload Me

Exit Sub
Rem // ---------------------------------------------------------------------------------------------------------------------------------------------------
Rem // --- This sub is if there is an Error, e.g. The File isnt There, is in Use ect...
ErrorCode:
 MsgBox Err.Description, vbExclamation + vbSystemModal, "Super Skin"
 Unload Me
 Unload frmProgress
End Sub

Private Sub cmdPath_Click()
Rem // --- Shows a Dialoug Box to Select the File, refer to the eg in 'frmSkinnApp'
On Error Resume Next
Dim sFile
Dim sIndex
With Me.dlgMain
 .DialogTitle = "Skined Application's Output."
 .Filter = "Application (*.exe)|*.exe|Application Manifest|*.manifest"
 .ShowOpen
 sFile = .FileName
 sIndex = .FilterIndex
 sTempName = .FileTitle
End With
 Me.txtPathName.Text = sFile
 Rem // --- If its an EXE then Allow Backup, if its just a Manifest then Dont Allow it.
 If sIndex = 1 Then EXE = True: Me.chkBackup.Enabled = True
 If sIndex = 2 Then EXE = False: Me.chkBackup.Enabled = False: Me.chkBackup.Value = 0
End Sub

Private Sub Form_Load()
Rem // --- Remove the "X" on the Form but keep the ICON.
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
End Sub
