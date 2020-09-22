VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Skin Options"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3675
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox picOPTBDR 
      BorderStyle     =   0  'None
      Height          =   1870
      Left            =   180
      ScaleHeight     =   1875
      ScaleWidth      =   3345
      TabIndex        =   2
      Top             =   460
      Width           =   3350
      Begin VB.CheckBox chkClear 
         Appearance      =   0  'Flat
         Caption         =   "Clear Skinned Log."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   $"frmOptions.frx":67E2
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3015
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4048
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Options"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Rem // --- If the user has 'Ticked' the Cleck Checkbox then Clear the Log / List on 'frmMain'
If Me.chkClear.Value = 1 Then frmMain.lvSkined.ListItems.Clear
Unload Me
End Sub

Private Sub Form_Load()
Rem // Disable the "X" Button, but Keep the Icon
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
End Sub
