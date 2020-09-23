VERSION 5.00
Begin VB.Form frmxSplash 
   BackColor       =   &H00000007&
   BorderStyle     =   0  'None
   ClientHeight    =   2175
   ClientLeft      =   5250
   ClientTop       =   3270
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "frmSplash.frx":48A8E
   ScaleHeight     =   2175
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoad 
      Interval        =   18
      Left            =   5640
      Top             =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   840
   End
   Begin VB.Shape pgbLoad 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      Height          =   90
      Index           =   2
      Left            =   360
      Top             =   1980
      Width           =   90
   End
   Begin VB.Shape pgbLoad 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      Height          =   90
      Index           =   1
      Left            =   240
      Top             =   1980
      Width           =   90
   End
   Begin VB.Shape pgbLoad 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      Height          =   90
      Index           =   0
      Left            =   120
      Top             =   1980
      Width           =   90
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "initializing system..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   4920
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmxSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cntr As Integer
Dim RegCtrl As Boolean

Private Sub Form_Load()
RegCtrl = False
End Sub

Private Sub tmrLoad_Timer()

'set the progress bar codes
For i = 0 To 2
If pgbLoad(i).Left = 1200 Then
    pgbLoad(i).Visible = False
    pgbLoad(i).Left = 120
    If pgbLoad(2).Left = 120 Then
        Cntr = Cntr + 1
    End If
End If
Next i
For i = 0 To 2
    If pgbLoad(i).Left = 120 Then
        pgbLoad(i).Visible = True
    End If
Next i

For i = 0 To 2
    pgbLoad(i).Left = pgbLoad(i).Left + 30
Next i
Load Cntr
End Sub

Private Sub Load(Cntr As Integer)
On Error GoTo ErrorMsg
'conditions on loading processes.
If Cntr = 1 Then
    lblStatus.Caption = "registering system components..."
    tmrLoad.Enabled = False
    If RegCtrl = False Then
        PathToDoc = App.Path & "\Run.bat"
        ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 0
        RegCtrl = True
    End If
    tmrLoad.Enabled = True
ElseIf Cntr = 2 Then
    tmrLoad.Enabled = False
    lblStatus.Caption = "setting up connection..."
    Call OpenCon
    tmrLoad.Enabled = True
ElseIf Cntr = 3 Then
    lblStatus.Caption = "scanning for security accounts..."
    'searching for recent accounts
    RunSql "SELECT * FROM tblAccountProfile"
    With Rs
        If .RecordCount = 0 Then
            MsgBox "No registered account found. Please add an account", vbInformation
            FrstUsr = True
            frmAccountProfile.Show
            frmAccountProfile.cmdSecurity.Enabled = False
            frmAccountProfile.txtFname.SetFocus
            tmrLoad.Enabled = False
            Exit Sub
        End If
    End With
   
    RunSql "Select * from tblAccountSecurity"
    With Rs
        If .RecordCount = 0 Then
            MsgBox "No Security Account found. Please set a security account", vbExclamation
            FrstUsr = True
            frmAcntManage.Show
            tmrLoad.Enabled = False
            Exit Sub
        End If
    End With
ElseIf Cntr = 4 Then
    tmrLoad.Enabled = False
    lblStatus.Caption = "scanning for system detections..."
    mdiMain.cmdWarnings.Caption = Warnings & " Warnings"
    mdiMain.cmdNotifications.Caption = Notifications & " Notifications"
    tmrLoad.Enabled = True
ElseIf Cntr = 5 Then
    Screen.MousePointer = 11
    mdiMain.Show
    Unload Me
    frmLogin.Show 1
End If
Exit Sub
'if unexpected error will occur
ErrorMsg:
    MsgBox "The system has failed to initialized.  Please contact your Administrator. " & vbNewLine & vbNewLine & _
            "Error Status: " & lblStatus & vbNewLine & "System Error: " & Err.Description, vbCritical
    End
End Sub

