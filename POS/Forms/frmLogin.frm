VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4095
   ClientLeft      =   5505
   ClientTop       =   5175
   ClientWidth     =   4470
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin ctrlButton.ThemedButton cmdCancel 
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmLogin.frx":038A
      Picture         =   "frmLogin.frx":0564
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLogin 
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Log-in"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmLogin.frx":08B8
      Picture         =   "frmLogin.frx":0A92
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   53
   End
   Begin VB.TextBox txtUserCode 
      Enabled         =   0   'False
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CheckBox chkNoAccess 
      Height          =   195
      Left            =   4080
      TabIndex        =   2
      Top             =   2760
      Width           =   180
   End
   Begin VB.TextBox txtPassword 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtUsername 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin ctrlButton.ThemedButton cmdOptions 
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Options >>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmLogin.frx":0DE6
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdHelp 
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Help"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmLogin.frx":0FC0
      Picture         =   "frmLogin.frx":119A
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdRetrieve 
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Retrieve"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmLogin.frx":14EE
      Picture         =   "frmLogin.frx":16C8
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3840
      Picture         =   "frmLogin.frx":1A1C
      Top             =   1480
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3960
      Picture         =   "frmLogin.frx":22E6
      Top             =   1120
      Width           =   240
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gain access to the system"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   720
      TabIndex        =   9
      Top             =   480
      Width           =   1620
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOG-IN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   645
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Left            =   120
      MouseIcon       =   "frmLogin.frx":2870
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":313A
      Stretch         =   -1  'True
      ToolTipText     =   "View warnings"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblUnable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unable to Log in?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   2760
      MouseIcon       =   "frmLogin.frx":3D7E
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2760
      Width           =   1230
   End
   Begin VB.Label lblcode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Secret code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   3120
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   915
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   -240
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub chkNoAccess_Click()
'if forgot password?
If chkNoAccess.Value = 1 Then
    txtUserCode.Enabled = True
    cmdRetrieve.Enabled = True
    cmdRetrieve.Default = True
    txtUserCode.SetFocus
Else
    txtUserCode.Enabled = False
    cmdRetrieve.Enabled = False
    cmdLogin.Default = True
    txtUsername.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
PathToDoc = App.Path & "\help.chm"
ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 5
End Sub

Private Sub cmdLogin_Click()

If TxtEmp(txtUsername) = True Then Exit Sub
If TxtEmp(txtPassword) = True Then Exit Sub

RunSql "SELECT * FROM tblAccountSecurity WHERE username = '" & txtUsername & "' and password = '" & txtPassword & "'"
With Rs
    If Rs.EOF = True Then
        MsgBox "Invalid 'Username' or 'Password'", vbExclamation
        txtUsername.SetFocus
        SelAll txtUsername
        Exit Sub
    End If
    Me.MousePointer = 11
    'set user values
    UserNo = .Fields!record_no
    UserNme = .Fields!UserName
    UserLvl = .Fields!Level
    UserId = .Fields!id
    mdiMain.cmdLogout.Caption = "&Log-out"
    Me.MousePointer = 0
    Me.Hide
    'set status bar on main form
    With mdiMain.sbrStatus
        .Panels(1).Text = "User: " & UserNme & "  "
        .Panels(2).Text = "Level: " & UserLvl & "  "
        .Panels(3).Text = "User ID: " & UserId & "  "
        .Panels(4).Text = "PC Name: " & PcId & "  "
    End With
    FrmShow UserLvl
End With
End Sub

Private Sub cmdOptions_Click()
    'show the options below
    If cmdOptions.Caption = "&Options >>" Then
        Me.Height = 4575
        cmdOptions.Caption = "&Options <<"
    Else
        Me.Height = 3135
        cmdOptions.Caption = "&Options >>"
        txtUsername.SetFocus
    End If
End Sub

Private Sub cmdRetrieve_Click()
If TxtEmp(txtUserCode) = True Then Exit Sub

'scan security accounts that mathches the security code
RunSql "Select * from tblAccountSecurity Where code = '" & txtUserCode & "'"
With Rs
    If .EOF = True Then
        MsgBox "No user account found", vbExclamation
        txtUserCode.SetFocus
        txtUserCode.SelStart = 0
        txtUserCode.SelLength = Len(txtUserCode)
        Exit Sub
    End If
    'load the user security info if code mathches
    MsgBox "Your Username is '" & .Fields!UserName & "'; Password is '" & .Fields!Password & "'", vbInformation
    txtUsername.SetFocus
    cmdLogin.Default = True
    chkNoAccess.Value = 0
End With
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
PcId = Environ("ComputerName")
Me.Caption = App.Title & " [" & PcId & "]"
Me.Height = 3135
End Sub

Private Sub Form_Unload(Cancel As Integer)
'cancel the process. Show the main form
mdiMain.cmdLogout.Caption = "&Log-in"
With mdiMain.sbrStatus
    For i = 1 To 4
        .Panels(i).Text = Empty
    Next i
End With
End Sub

Private Sub lblUnable_Click()
If chkNoAccess.Value = 0 Then
    chkNoAccess.Value = 1
Else
    chkNoAccess.Value = 0
End If
Call chkNoAccess_Click
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub
Private Sub txtUserCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub txtUsername_GotFocus()
txtUsername.SelStart = 0
txtUsername.SelLength = Len(txtUsername)
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub
