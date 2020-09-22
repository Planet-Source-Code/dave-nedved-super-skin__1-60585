VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSkinApp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skin Application"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   Icon            =   "frmSkinApp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdSkin 
      Caption         =   "Skin!"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdOutputPath 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   300
      Left            =   4440
      TabIndex        =   12
      Top             =   3020
      Width           =   300
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Manuel Skin Output"
      Enabled         =   0   'False
      Height          =   645
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   2790
      Width           =   4815
      Begin VB.TextBox txtOutputPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   4120
      End
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   3960
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Output Options"
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1935
      Width           =   4815
      Begin VB.CheckBox chkBackup 
         Appearance      =   0  'Flat
         Caption         =   "Create Backup of Application"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Skin Type"
      Height          =   1080
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   860
      Width           =   4815
      Begin VB.PictureBox picBG 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   4575
         TabIndex        =   4
         Top             =   240
         Width           =   4575
         Begin VB.OptionButton optManifest2 
            Appearance      =   0  'Flat
            Caption         =   "AUTO - Full Skin Manifest#2 (Use if #1 Doesn’t Work)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4335
         End
         Begin VB.OptionButton optManifestManuel 
            Appearance      =   0  'Flat
            Caption         =   "MANUEL - Just Create the IDE Manifest"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   3255
         End
         Begin VB.OptionButton optManifest 
            Appearance      =   0  'Flat
            Caption         =   "AUTO - Full Skin Manifest#1 (Almost Always Works)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   0
            Value           =   -1  'True
            Width           =   4455
         End
      End
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   300
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   300
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Application To Skin"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtPathName 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4120
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   $"frmSkinApp.frx":67E2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   4575
   End
End
Attribute VB_Name = "frmSkinApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim sTempName

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOutputPath_Click()
Rem // --- Show a Dialoug Box asking where to save the 'manifest' File
On Error Resume Next
Dim sFile
With Me.dlgMain
 .DialogTitle = "Select Application To Skin."
 .Filter = "Application Manifest|*.manifest"
 .ShowSave
 sFile = .FileName
 sTempName = .FileTitle
End With
 Rem // --- Stick the Path into the Text Box
 Me.txtOutputPath.Text = sFile
End Sub

Private Sub cmdPath_Click()
Rem // --- Show a Dialoug Box asking where the 'EXE' to Skin is.
On Error Resume Next
Dim sFile
Dim sIndex
With Me.dlgMain
 Rem // --- This Is The Dialoug Title
 .DialogTitle = "Skined Application's Output."
 Rem // --- This is the Filter, thats the thing where you select the Format.
 .Filter = "Applications (*.exe)|*.exe|All Files (*.*)"
 Rem // --- Show the Box Now.
 .ShowOpen
 sFile = .FileName
 sIndex = .FilterIndex
 sTempName = .FileTitle
End With
 Me.txtPathName.Text = sFile
 Rem // --- If the user has selected an 'All File' the warn them that it might not be skinnable.
 If sIndex = 2 Then MsgBox "Super Skin might not be able to Skin the File you have selected," & vbNewLine & "As The file you have Selected may not be a Valid Application.", vbExclamation, "Super Skin"
End Sub

Private Sub cmdSkin_Click()
Rem // --- logval is just used for writing the log
Dim logVal As Variant

Rem // ---------------------------------------------------------------------------------------------------------------------------------------------------

Rem // --- Check that everything is right, this should stop errors.
If Me.optManifest.Value = True Then
 If Me.txtPathName.Text = "" Then MsgBox "Please Select an Application To Skin.", vbExclamation, "Super Skin": Exit Sub
End If
If Me.optManifest2.Value = True Then
 If Me.txtPathName.Text = "" Then MsgBox "Please Select an Application To Skin.", vbExclamation, "Super Skin": Exit Sub
End If
If Me.optManifestManuel.Value = True Then
 If Me.txtOutputPath.Text = "" Then MsgBox "Please Specify an Manifest Output File.", vbExclamation, "Super Skin": Exit Sub
End If

Rem // ---------------------------------------------------------------------------------------------------------------------------------------------------

On Error Resume Next

Rem // --- Show the Progress Form.
Me.Hide
frmProgress.Show

Rem // --- Check if Backup Wanted
DoEvents
If Me.chkBackup.Value = 1 Then
 DoEvents
 FileCopy Me.txtPathName.Text, Me.txtPathName.Text & " (Backed Up).exe"
 DoEvents
 frmProgress.pbrProcess.Value = 0
End If

DoEvents
frmProgress.pbrProcess.Value = 0

Rem // --- Create Manifest (#1 Type)
DoEvents
If Me.optManifest.Value = True Then
 DoEvents
 FileCopy App.Path & "\Resources\resource1.style.tempfiless", Me.txtPathName.Text & ".manifest"   ' Define source file name & Path.
 DoEvents
 frmProgress.pbrProcess.Value = 0
End If

DoEvents
frmProgress.pbrProcess.Value = 0

Rem // --- Create Manifest (#2 Type)
DoEvents
If Me.optManifest2.Value = True Then
 DoEvents
 FileCopy App.Path & "\Resources\resource2.style.tempfiless", Me.txtPathName.Text & ".manifest"   ' Define source file name & Path.
 DoEvents
 frmProgress.pbrProcess.Value = 0
End If

DoEvents
frmProgress.pbrProcess.Value = 0

Rem // --- Create Manifest (Single IDE Type[No EXE Needed, Just create a manifest on its own])
If Me.optManifestManuel.Value = True Then
 DoEvents
 FileCopy App.Path & "\Resources\resource2.style.tempfiless", Me.txtOutputPath.Text
 DoEvents
 frmProgress.pbrProcess.Value = 0
End If

Rem // --- Add the Skinned Application to the Log on 'frmMain'
If Me.optManifest.Value = True Then
 Set logVal = frmMain.lvSkined.ListItems.Add(, , sTempName)
  logVal.SubItems(1) = "Skin - Manifest #1"
  If Me.chkBackup.Value = 1 Then
   logVal.SubItems(2) = "Yes"
  Else
   logVal.SubItems(2) = "No"
  End If
  logVal.SubItems(3) = Me.txtPathName.Text
  logVal.SubItems(4) = Date & " - " & Time
End If
If Me.optManifest2.Value = True Then
 Set logVal = frmMain.lvSkined.ListItems.Add(, , sTempName)
  logVal.SubItems(1) = "Skin - Manifest #2"
  If Me.chkBackup.Value = 1 Then
   logVal.SubItems(2) = "Yes"
  Else
   logVal.SubItems(2) = "No"
  End If
  logVal.SubItems(3) = Me.txtPathName.Text
  logVal.SubItems(4) = Date & " - " & Time
End If
If Me.optManifestManuel.Value = True Then
 Set logVal = frmMain.lvSkined.ListItems.Add(, , sTempName)
  logVal.SubItems(1) = "Skin - Direct IDE Manifest"
  If Me.chkBackup.Value = 1 Then
   logVal.SubItems(2) = "Yes"
  Else
   logVal.SubItems(2) = "No"
  End If
  logVal.SubItems(3) = Me.txtOutputPath.Text
  logVal.SubItems(4) = Date & " - " & Time
End If
Unload Me
End Sub

Private Sub Form_Load()
Rem // --- Disable the "X" on the form, but keep the Icon.
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
End Sub

Private Sub optManifest_Click()
Rem // --- Enable / Disable the IDE Manifest Text Box
If Me.optManifestManuel.Value = True Then
 Me.fraInfo(3).Enabled = True
 Me.txtOutputPath.Enabled = True
 Me.cmdOutputPath.Enabled = True
 Me.chkBackup.Enabled = False
Else
 Me.fraInfo(3).Enabled = False
 Me.txtOutputPath.Enabled = False
 Me.cmdOutputPath.Enabled = False
 Me.chkBackup.Enabled = True
End If
End Sub

Private Sub optManifest_GotFocus()
Rem // --- Enable / Disable the IDE Manifest Text Box
If Me.optManifestManuel.Value = True Then
 Me.fraInfo(3).Enabled = True
 Me.txtOutputPath.Enabled = True
 Me.cmdOutputPath.Enabled = True
 Me.chkBackup.Enabled = False
Else
 Me.fraInfo(3).Enabled = False
 Me.txtOutputPath.Enabled = False
 Me.cmdOutputPath.Enabled = False
 Me.chkBackup.Enabled = True
End If
End Sub

Private Sub optManifest2_Click()
Rem // --- Enable / Disable the IDE Manifest Text Box
If Me.optManifestManuel.Value = True Then
 Me.fraInfo(3).Enabled = True
 Me.txtOutputPath.Enabled = True
 Me.cmdOutputPath.Enabled = True
 Me.chkBackup.Enabled = False
Else
 Me.fraInfo(3).Enabled = False
 Me.txtOutputPath.Enabled = False
 Me.cmdOutputPath.Enabled = False
 Me.chkBackup.Enabled = True
End If
End Sub

Private Sub optManifest2_GotFocus()
Rem // --- Enable / Disable the IDE Manifest Text Box
If Me.optManifestManuel.Value = True Then
 Me.fraInfo(3).Enabled = True
 Me.txtOutputPath.Enabled = True
 Me.cmdOutputPath.Enabled = True
 Me.chkBackup.Enabled = False
Else
 Me.fraInfo(3).Enabled = False
 Me.txtOutputPath.Enabled = False
 Me.cmdOutputPath.Enabled = False
 Me.chkBackup.Enabled = True
End If
End Sub

Private Sub optManifestManuel_Click()
Rem // --- Enable / Disable the IDE Manifest Text Box
If Me.optManifestManuel.Value = True Then
 Me.fraInfo(3).Enabled = True
 Me.txtOutputPath.Enabled = True
 Me.cmdOutputPath.Enabled = True
 Me.chkBackup.Enabled = False
Else
 Me.fraInfo(3).Enabled = False
 Me.txtOutputPath.Enabled = False
 Me.cmdOutputPath.Enabled = False
 Me.chkBackup.Enabled = True
End If
End Sub

Private Sub optManifestManuel_GotFocus()
Rem // --- Enable / Disable the IDE Manifest Text Box
If Me.optManifestManuel.Value = True Then
 Me.fraInfo(3).Enabled = True
 Me.txtOutputPath.Enabled = True
 Me.cmdOutputPath.Enabled = True
 Me.chkBackup.Enabled = False
Else
 Me.fraInfo(3).Enabled = False
 Me.txtOutputPath.Enabled = False
 Me.cmdOutputPath.Enabled = False
 Me.chkBackup.Enabled = True
End If
End Sub
