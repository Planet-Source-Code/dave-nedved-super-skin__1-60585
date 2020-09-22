VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Super Skin"
   ClientHeight    =   6360
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8535
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView lvSkined 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Skined Application"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Skin Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Created Backup?"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Applications Path"
         Object.Width           =   14111
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Skinned On"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Frame fraMain 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   5655
      Begin VB.PictureBox picHelp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2280
         Picture         =   "frmMain.frx":67E2
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   6
         ToolTipText     =   "About Super Skin - F1"
         Top             =   160
         Width           =   615
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         Picture         =   "frmMain.frx":7095
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   5
         ToolTipText     =   "Options - Ctrl+K"
         Top             =   160
         Width           =   615
      End
      Begin VB.PictureBox picUnSkinn 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   840
         Picture         =   "frmMain.frx":7962
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   4
         ToolTipText     =   "Un Skin Application - Ctrl+U"
         Top             =   160
         Width           =   615
      End
      Begin VB.PictureBox picSkin 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmMain.frx":7F84
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   3
         ToolTipText     =   "Skin Application - Ctrl + S"
         Top             =   160
         Width           =   615
      End
   End
   Begin VB.Label lblInfo 
      Caption         =   "Applications Skinned && UnSkinned by Super Skin"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   70
      Width           =   7215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSkinApplication 
         Caption         =   "&Skin Application"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileUnSkinApplication 
         Caption         =   "&Un Skin Application"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuFileSkinApplicationBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpDSOnline 
         Caption         =   "DaTo Software &Online"
         Begin VB.Menu mnuHelpDSOnlineDSWebsite 
            Caption         =   "&DaTo Software Website"
         End
         Begin VB.Menu mnuHelpDSOnlineSuperSkinWebsite 
            Caption         =   "&Super Skin Online"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem // ===========================================================================
Rem // = Super Skin is a Program used to Skin your Own Applications.             =
Rem // = Or to Skin applications where the developer couldnt be bothered         =
Rem // = To skin the Application.                                                =
Rem // =                                                                         =
Rem // = I Was Inspired to write this Application when i was running             =
Rem // = Window Blinds. I had my PC All skinned to a mac OS X desktop and it     =
Rem // = Looked verry nice... Untill i ran a program that wasnt 'Theme Aware'    =
Rem // =                                                                         =
Rem // = This program creates a XML for the Application, making the Application  =
Rem // = Run Native on Windows XP, 2006 ect...                                   =
Rem // =                                                                         =
Rem // = You may use this Program to Skin your own Comercial Applications...     =
Rem // = Just rember to go to PSCode and Vote + Leave your Comments.             =
Rem // = And check out the DaTo Software Website: www.datosoftware.com           =
Rem // =                                                                         =
Rem // = Enjoy, oh and if you have any Comments + Sugestions, or you want to     =
Rem // = Help in Future Versions of Super Skin, Please E-Mail me:                =
Rem // = dnedved@datosoftware.com, or dnedved@gmail.com                          =
Rem // =                                                                         =
Rem // = P.S. The code to Subclass the Flash, and Skin the Splash & About Screen =
Rem // =      I Didn't crete. I Couldnt find the rightfull owner for this code,  =
Rem // =      So if you are the Owner, please E-Mail Me and i will be MORE than  =
Rem // =      Happy to Put you in the Credits.                                   =
Rem // ===========================================================================

Option Explicit
 Rem // This is used to Shell to www.datosoftware.com
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 Dim LastSave As String

Rem // --- I Always put on error code in just incase a error somehow forms :)
Rem // --- I like to be Safe.

Private Sub Form_Load()
Rem // --- logval is just used for writing the log
Dim logVal As Variant
Dim LC, TempLCcount, TempLCName, TempLCPath, TempLCDate, TempLCSkinType, TempLCBackup As String
Dim sFile As String

Rem // --- Put in Error Code, in case the Reg Keys are not There.
On Error Resume Next
Rem // --- Resume last Settings
Me.Top = GetSetting(App.Title, "Settings\POS", "Main X")
Me.Left = GetSetting(App.Title, "Settings\POS", "Main Y")
Me.Height = GetSetting(App.Title, "Settings\POS", "Main H")
Me.Width = GetSetting(App.Title, "Settings\POS", "Main W")
Me.WindowState = GetSetting(App.Title, "Settings\POS", "Main WS")

Rem // --- Set what we need to Load the Log to: sFile
sFile = App.Path & "\Config\Log.CFG"

Rem // --- Temp Count is a count of the Total Skins
TempLCcount = ReadINI("List", "Count", sFile)
    
Rem // --- Load the Log File into the List View.
For LC = 1 To TempLCcount
 TempLCName = ReadINI("Name", "C:" & LC, sFile)
 TempLCPath = ReadINI("Path", "C:" & LC, sFile)
 TempLCDate = ReadINI("Date", "C:" & LC, sFile)
 TempLCSkinType = ReadINI("SkinType", "C:" & LC, sFile)
 TempLCBackup = ReadINI("Backup", "C:" & LC, sFile)
    
 Set logVal = Me.lvSkined.ListItems.Add(LC, , TempLCName)
  logVal.SubItems(1) = TempLCSkinType
  logVal.SubItems(2) = TempLCBackup
  logVal.SubItems(3) = TempLCPath
  logVal.SubItems(4) = TempLCDate
Next LC
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rem // --- Call the Mouse Move on the Frame, so the Picture Box / Buttons
Rem // --- Don't 'Have Focus'
fraMain_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Rem // --- Ask if program close wanted.
Dim msgres As Variant
msgres = MsgBox("Are you sure you want to Exit Super Skin?", vbExclamation + vbYesNo + vbDefaultButton2, "Exit Super Skin?")
If msgres = vbNo Then Cancel = 1: Exit Sub
End Sub

Private Sub Form_Resize()
Rem // --- Resize Controls On Form ---
Rem // Stick Error code in incase form is resized smaller than control
On Error Resume Next
Me.lvSkined.Width = Me.ScaleWidth - 120 - 120
Me.lvSkined.Height = Me.ScaleHeight - 120 - 120 - 120 - 120 - Me.fraMain.Height
Me.fraMain.Top = Me.lvSkined.Height + 120 + 120 + 120
Me.fraMain.Width = Me.ScaleWidth - 120 - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sFile As String
Dim ListC As Variant

Rem // --- Save the Pos of the Main Form.

On Error Resume Next
SaveSetting App.Title, "Settings\POS", "Main X", Me.Top
SaveSetting App.Title, "Settings\POS", "Main Y", Me.Left
SaveSetting App.Title, "Settings\POS", "Main H", Me.Height
SaveSetting App.Title, "Settings\POS", "Main W", Me.Width
SaveSetting App.Title, "Settings\POS", "Main WS", Me.WindowState

Rem // --- Save the Log To File (Config\Log.cfg)

sFile = App.Path & "\Config\Log.CFG"

WriteINI "Super Skin Log File", "", "", sFile
    
WriteINI "List", "Count", Me.lvSkined.ListItems.Count, sFile
    
For ListC = 0 To Me.lvSkined.ListItems.Count
  WriteINI "Name", "C:" & ListC, Me.lvSkined.ListItems.Item(ListC).Text, sFile
  WriteINI "SkinType", "C:" & ListC, Me.lvSkined.ListItems.Item(ListC).SubItems(1), sFile
  WriteINI "Backup", "C:" & ListC, Me.lvSkined.ListItems.Item(ListC).SubItems(2), sFile
  WriteINI "Path", "C:" & ListC, Me.lvSkined.ListItems.Item(ListC).SubItems(3), sFile
  WriteINI "Date", "C:" & ListC, Me.lvSkined.ListItems.Item(ListC).SubItems(4), sFile
Next ListC

Rem // --- Shut Down Super Skin
End
End Sub

Private Sub fraMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rem // --- I had to stick 'If Not' Code in beacuse the Pictures are
Rem // --- Transparent. Otherwise the Pictures will 'Flicker' every time
Rem // --- You put your Mouse Over It.
Rem // --- If Not code is just saying that if this is NOT that then Do This...
Rem // --- Now Remove focus on any Buttons That Have Focus.
On Error Resume Next
If Not Me.picSkin.BackColor = vbButtonFace Then
 Me.picSkin.BackColor = vbButtonFace
End If
If Not Me.picSkin.BorderStyle = 0 Then
 Me.picSkin.BorderStyle = 0
End If
If Not Me.picUnSkinn.BackColor = vbButtonFace Then
 Me.picUnSkinn.BackColor = vbButtonFace
End If
If Not Me.picUnSkinn.BorderStyle = 0 Then
 Me.picUnSkinn.BorderStyle = 0
End If
If Not Me.picOptions.BackColor = vbButtonFace Then
 Me.picOptions.BackColor = vbButtonFace
End If
If Not Me.picOptions.BorderStyle = 0 Then
 Me.picOptions.BorderStyle = 0
End If
If Not Me.picHelp.BackColor = vbButtonFace Then
 Me.picHelp.BackColor = vbButtonFace
End If
If Not Me.picHelp.BorderStyle = 0 Then
 Me.picHelp.BorderStyle = 0
End If
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rem // --- Call the Mouse Move on the Frame, so the Picture Box / Buttons
Rem // --- Don't 'Have Focus'
fraMain_MouseMove Button, Shift, X, Y
End Sub

Private Sub lvSkined_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rem // --- Call the Mouse Move on the Frame, so the Picture Box / Buttons
Rem // --- Don't 'Have Focus'
fraMain_MouseMove Button, Shift, X, Y
End Sub

Private Sub mnuFileExit_Click()
Rem // --- Close the Program ---
Unload Me
End Sub

Private Sub mnuFileSkinApplication_Click()
Rem // --- Show the Skin APP Form ---
On Error Resume Next
frmSkinApp.Show vbModal, Me
End Sub

Private Sub mnuFileUnSkinApplication_Click()
Rem // --- Show the UnSkin APP Form ---
On Error Resume Next
frmUnSkinAPP.Show vbModal, Me
End Sub

Private Sub mnuHelpAbout_Click()
On Error Resume Next
Rem // --- Show the About Form.
frmAbout.Show vbModal
Rem // --- Disable this form, so users cant Click it.
Me.Enabled = False
End Sub

Private Sub mnuHelpDSOnlineDSWebsite_Click()
Rem // --- Go To The DaTo Software Website
ShellExecute Me.hWnd, "open", "http://www.datosoftware.com", 0&, LastSave, vbNormalFocus
End Sub

Private Sub mnuHelpDSOnlineSuperSkinWebsite_Click()
Rem // --- Go To The Super Skin Website.
ShellExecute Me.hWnd, "open", "http://www.datosoftware.com/products/superskin", 0&, LastSave, vbNormalFocus
End Sub

Private Sub mnuToolsOptions_Click()
On Error Resume Next
Rem // --- Show the Options Dialoug.
frmOptions.Show vbModal
End Sub

Private Sub picHelp_Click()
Rem // --- Call the About Sub
mnuHelpAbout_Click
End Sub

Private Sub picHelp_LostFocus()
Rem // --- If this Picture has Lost Focus then Remove the Box
Rem // --- Arround the Button.
If Not Me.picHelp.BackColor = vbButtonFace Then
 Me.picHelp.BackColor = vbButtonFace
End If
If Not Me.picHelp.BorderStyle = 0 Then
 Me.picHelp.BorderStyle = 0
End If
End Sub

Private Sub picHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rem // --- Make the Buttons go '3D' when you put your Mouse Over Them.
Rem // --- Again Using If Not Statements (See the Frame Mouse Move For More Info)
If Not Me.picHelp.BorderStyle = 1 Then
 Me.picHelp.BorderStyle = 1
End If
If Not Me.picHelp.BackColor = vbScrollBars Then
 Me.picHelp.BackColor = vbScrollBars
End If
End Sub

Private Sub picOptions_Click()
Rem // --- Call the Options Sub
mnuToolsOptions_Click
End Sub

Private Sub picOptions_LostFocus()
Rem // --- If this Picture has Lost Focus then Remove the Box
Rem // --- Arround the Button.
If Not Me.picOptions.BackColor = vbButtonFace Then
 Me.picOptions.BackColor = vbButtonFace
End If
If Not Me.picOptions.BorderStyle = 0 Then
 Me.picOptions.BorderStyle = 0
End If
End Sub

Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rem // --- Make the Buttons go '3D' when you put your Mouse Over Them.
Rem // --- Again Using If Not Statements (See the Frame Mouse Move For More Info)
If Not Me.picOptions.BorderStyle = 1 Then
 Me.picOptions.BorderStyle = 1
End If
If Not Me.picOptions.BackColor = vbScrollBars Then
 Me.picOptions.BackColor = vbScrollBars
End If
End Sub

Private Sub picSkin_Click()
Rem // --- Call the Skinn APP Sub
mnuFileSkinApplication_Click
End Sub

Private Sub picSkin_LostFocus()
Rem // --- If this Picture has Lost Focus then Remove the Box
Rem // --- Arround the Button.
If Not Me.picSkin.BackColor = vbButtonFace Then
 Me.picSkin.BackColor = vbButtonFace
End If
If Not Me.picSkin.BorderStyle = 0 Then
 Me.picSkin.BorderStyle = 0
End If
End Sub

Private Sub picSkin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rem // --- Make the Buttons go '3D' when you put your Mouse Over Them.
Rem // --- Again Using If Not Statements (See the Frame Mouse Move For More Info)
If Not Me.picSkin.BorderStyle = 1 Then
 Me.picSkin.BorderStyle = 1
End If
If Not Me.picSkin.BackColor = vbScrollBars Then
 Me.picSkin.BackColor = vbScrollBars
End If
End Sub

Private Sub picUnSkinn_Click()
Rem // --- Call the UnSkinn App Sub
mnuFileUnSkinApplication_Click
End Sub

Private Sub picUnSkinn_LostFocus()
Rem // --- If this Picture has Lost Focus then Remove the Box
Rem // --- Arround the Button.
If Not Me.picUnSkinn.BackColor = vbButtonFace Then
 Me.picUnSkinn.BackColor = vbButtonFace
End If
If Not Me.picUnSkinn.BorderStyle = 0 Then
 Me.picUnSkinn.BorderStyle = 0
End If
End Sub

Private Sub picUnSkinn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rem // --- Make the Buttons go '3D' when you put your Mouse Over Them.
Rem // --- Again Using If Not Statements (See the Frame Mouse Move For More Info)
If Not Me.picUnSkinn.BorderStyle = 1 Then
 Me.picUnSkinn.BorderStyle = 1
End If
If Not Me.picUnSkinn.BackColor = vbScrollBars Then
 Me.picUnSkinn.BackColor = vbScrollBars
End If
End Sub
