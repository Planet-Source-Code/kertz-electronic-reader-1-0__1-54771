VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Electronic Reader"
   ClientHeight    =   3525
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   10740
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2433.017
   ScaleMode       =   0  'User
   ScaleWidth      =   10085.42
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   5400
      ScaleHeight     =   3435
      ScaleWidth      =   5235
      TabIndex        =   12
      Top             =   0
      Width           =   5295
      Begin VB.Image Image4 
         Height          =   3240
         Left            =   0
         Picture         =   "frmAbout.frx":0442
         Top             =   120
         Width           =   5400
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   5239
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3015
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   5318
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         FileName        =   "G:\Backup\G\Amal\PSC codes\Electronic Reader\License.txt"
         TextRTF         =   $"frmAbout.frx":768F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   5239
      Begin VB.Image Image3 
         Height          =   105
         Left            =   0
         Picture         =   "frmAbout.frx":7B91
         Top             =   2880
         Width           =   5820
      End
      Begin VB.Image Image2 
         Height          =   105
         Left            =   0
         Picture         =   "frmAbout.frx":7F35
         Top             =   240
         Width           =   5820
      End
      Begin VB.Image Image1 
         Height          =   105
         Left            =   0
         Picture         =   "frmAbout.frx":82D9
         Top             =   1680
         Width           =   5820
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":867D
         Height          =   975
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "All Rights Reserved"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Version 1.0.0"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Copyright © 2003  AMAL R S"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Electronic Reader"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   11.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7080
      TabIndex        =   0
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Frame Frame3 
      Height          =   3135
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   5235
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   3015
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   5318
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         FileName        =   "G:\Backup\G\Amal\Electronic Reader\Contact.txt"
         TextRTF         =   $"frmAbout.frx":8742
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3555
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   6271
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Information"
            Key             =   "a"
            Object.ToolTipText     =   "Information"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Licence Agreement"
            Key             =   "b"
            Object.ToolTipText     =   "Licence Agreement"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "How To Contact"
            Key             =   "c"
            Object.ToolTipText     =   "How To Contact"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "System Info"
            Key             =   "d"
            Object.ToolTipText     =   "System Information..."
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdOK_Click()
'unloads frmAbout...
Unload frmAbout
End Sub

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Unload(Cancel As Integer)
'activates frmMain...
frmMain.Enabled = True
End Sub

Private Sub TabStrip1_Click()
'determines which frame to show...
If TabStrip1.SelectedItem.Key = "a" Then
    Frame2.Visible = False
    Frame3.Visible = False
    Frame1.Visible = True
End If
If TabStrip1.SelectedItem.Key = "b" Then
    Frame1.Visible = False
    Frame3.Visible = False
    Frame2.Visible = True
End If
If TabStrip1.SelectedItem.Key = "c" Then
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = True
End If
If TabStrip1.SelectedItem.Key = "d" Then
    cmdSysInfo_Click
End If
End Sub
