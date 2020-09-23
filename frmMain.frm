VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Electronic Reader "
   ClientHeight    =   3525
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6135
   FillColor       =   &H00808080&
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3270
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3069
            Text            =   "Version 1.0.0"
            TextSave        =   "Version 1.0.0"
            Object.ToolTipText     =   "Current Version..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock Status... "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:25 PM"
            Object.ToolTipText     =   "System Time..."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "02/07/2004"
            Object.ToolTipText     =   "System Date..."
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMain.frx":27A2
      Left            =   2520
      List            =   "frmMain.frx":27C4
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Voice Types..."
      Top             =   50
      Width           =   3615
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons(2)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Book06"
            Object.ToolTipText     =   "Begin Reading"
            ImageKey        =   "Book06"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Trffc14"
            Object.ToolTipText     =   "Stop Reading"
            ImageKey        =   "Trffc14"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtdata 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":2862
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   960
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   2
      Left            =   960
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28E4
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29F6
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B08
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C1A
            Key             =   "Book06"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F34
            Key             =   "Trffc14"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":324E
            Key             =   "Mike"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3568
            Key             =   "Msgbox04"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3882
            Key             =   "W95mbx01"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B9C
            Key             =   "Book01a"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EB6
            Key             =   "Book04"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41D0
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42E2
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43F4
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   6480
      Top             =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As "
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu controlsmnu 
      Caption         =   "&Controls"
      Begin VB.Menu begin 
         Caption         =   "&Begin Reading"
         Shortcut        =   {F5}
      End
      Begin VB.Menu end1 
         Caption         =   "&End Reading"
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      WindowList      =   -1  'True
      Begin VB.Menu contents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu readme 
         Caption         =   "&Readme..."
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Electronic Reader..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declares costants and API...
Dim pathx, sfile, fname As String
Dim shellx As Long
Dim Genie As IAgentCtlCharacter
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub copy_Click()
'copies to clipboard
Clipboard.Clear
Clipboard.SetText txtdata.SelText
End Sub

Private Sub cut_Click()
'cut's specified data...
Clipboard.Clear
Clipboard.SetText txtdata.SelText
txtdata.SelText = ""
End Sub

Private Sub mnuFileSave_Click()
'checks whether file is already saved...
If sfile = "" Then
    mnuFileSaveAs_Click
Else
    'saves the current changes
    txtdata.SaveFile sfile
End If
End Sub

Private Sub mnuFileSaveAs_Click()
    'opens save dialog box
    With dlgCommonDialog
    .FileName = ""
    .CancelError = False
    .Filter = "Rich Text Format|*.rtf|Text Document|*.txt|Write Document|*.wri|All Files|*.*"
    .Flags = cdlCCPreventFullOpen
    .DialogTitle = "Save As"
    .ShowSave
    sfile = .FileName
    End With
    txtdata.SaveFile sfile
End Sub
Private Sub contents_Click()
'opens help file...
On Error GoTo errh2:
shellx = ShellExecute(0, "open", App.Path & "\help\help.htm", "", "", 4)
errh2:
If Err.Number <> 0 Then
    MsgBox "Help file does not exist", vbCritical
End If
End Sub

Private Sub paste_Click()
'pastes data
txtdata.SelText = Clipboard.GetText()
End Sub
Private Sub readme_Click()
'opening readme file...
On Error GoTo errh1:
shellx = ShellExecute(0, "open", App.Path & "\readme.txt", "", "", 4)
errh1:
If Err.Number <> 0 Then
    MsgBox "The readme file does not exist or may be damaged", vbCritical
End If
End Sub
Private Sub txtdata_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'initialising popup menu...
If Button = 2 Then
    PopupMenu edit
End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    'setting toolbar...
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Book06"
            begin_Click
        Case "Trffc14"
            end1_Click
        Case "Help"
            contents_Click
    End Select
End Sub
Private Sub Agent1_Hide(ByVal CharacterID As String, ByVal Cause As Integer)
'what to do when reading is stoped
Toolbar1.Buttons.Item(6).Enabled = True
begin.Enabled = True
Toolbar1.Buttons.Item(7).Enabled = False
end1.Enabled = False
Combo1.Enabled = True
txtdata.Locked = False
mnuFile.Enabled = True
mnuHelp.Enabled = True
edit.Enabled = True
Toolbar1.Buttons.Item(1).Enabled = True
Toolbar1.Buttons.Item(2).Enabled = True
Toolbar1.Buttons.Item(4).Enabled = True
Toolbar1.Buttons.Item(9).Enabled = True
End Sub
Private Sub begin_Click()
'starts reading text...
If txtdata.Text <> "" Then
end1.Enabled = True
begin.Enabled = False
Toolbar1.Buttons.Item(7).Enabled = True
Toolbar1.Buttons.Item(6).Enabled = False
Toolbar1.Buttons.Item(1).Enabled = False
Toolbar1.Buttons.Item(2).Enabled = False
Toolbar1.Buttons.Item(4).Enabled = False
Toolbar1.Buttons.Item(9).Enabled = False
Combo1.Enabled = False
txtdata.Locked = True
mnuFile.Enabled = False
mnuHelp.Enabled = False
edit.Enabled = False
If txtdata.SelText = "" Then
Genie.Show
Genie.Speak txtdata.Text
Genie.Hide
Else
Genie.Show
Genie.Speak txtdata.SelText
Genie.Hide
End If
Else
MsgBox "Enter text for reading.", vbInformation, "Electronic Reader"
End If
End Sub
Private Sub Combo1_Click()
'decides the filename for the the selected character...
On Error GoTo errh3
Agent1.Characters.Unload "Genie"
    If Combo1.Text = Combo1.List(0) Then
    fname = "\female1.acf"
    End If
    If Combo1.Text = Combo1.List(1) Then
    fname = "\female2.acf"
    End If
    If Combo1.Text = Combo1.List(2) Then
    fname = "\male1.acf"
    End If
    If Combo1.Text = Combo1.List(3) Then
    fname = "\male2.acf"
    End If
    If Combo1.Text = Combo1.List(4) Then
    fname = "\male3.acf"
    End If
    If Combo1.Text = Combo1.List(5) Then
    fname = "\male4.acf"
    End If
    If Combo1.Text = Combo1.List(6) Then
    fname = "\male5.acf"
    End If
    If Combo1.Text = Combo1.List(7) Then
    fname = "\male6.acf"
    End If
    If Combo1.Text = Combo1.List(8) Then
    fname = "\male7.acf"
    End If
    If Combo1.Text = Combo1.List(9) Then
    fname = "\male8.acf"
    End If
    'sets the selected character as current character
    Agent1.Characters.Load "Genie", pathx + fname
    Set Genie = Agent1.Characters("Genie")
errh3:
If Err.Number <> 0 Then
    MsgBox Err.Description, vbCritical, "Electronic Reader"
End If
End Sub
Private Sub end1_Click()
'stops reading...
Genie.Stop
end1.Enabled = False
End Sub
Private Sub Form_Load()
On Error GoTo errh4:
    'sets the path of characters
    pathx = App.Path & "\Characters"
    'sets the character
    Agent1.Characters.Load "Genie", pathx & "\female1.acf"
    Set Genie = Agent1.Characters("Genie")
    Combo1.Text = Combo1.List(0)
errh4:
If Err.Number <> 0 Then
    MsgBox Err.Description, vbCritical, "Electronic Reader"
    End
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'ends the program
End
End Sub
Private Sub mnuHelpAbout_Click()
'shows about dialog after deactivating frmMain...
frmMain.Enabled = False
frmAbout.Show
End Sub
Private Sub mnuFileExit_Click()
    'exits the program...
    Unload Me
End Sub
Private Sub mnuFileOpen_Click()
     'prompts the user to open a text file
     With dlgCommonDialog
        sfile = ""
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "All Supported Files (*.RTF,*.TXT,*.WRI)|*.rtf;*.txt;*.wri|All Files (*.*)|*.*"
        .InitDir = App.Path
        .Flags = cdlOFNHideReadOnly
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sfile = .FileName
    End With
    txtdata.FileName = sfile
End Sub
Private Sub mnuFileNew_Click()
 'clears the whole text box...
 sfile = ""
 txtdata.Text = ""
End Sub

