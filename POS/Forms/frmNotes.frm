VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmNotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "frmNotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin ctrlButton.ThemedButton cmdViewLog 
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&View Log"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNotes.frx":038A
      Picture         =   "frmNotes.frx":0564
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdScheduler 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Sche&duler"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNotes.frx":08B8
      Picture         =   "frmNotes.frx":0A92
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   6735
      Begin VB.TextBox txtNotes 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   360
         Width           =   6255
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNotes.frx":0DE6
      Picture         =   "frmNotes.frx":0FC0
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNotes.frx":1314
      Picture         =   "frmNotes.frx":14EE
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View and manage system notes"
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
      TabIndex        =   3
      Top             =   480
      Width           =   2010
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTES"
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
      TabIndex        =   2
      Top             =   120
      Width           =   585
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   120
      Picture         =   "frmNotes.frx":1842
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Saved As Boolean
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Open App.Path & "\Notes.log" For Output As #1
Print #1, txtNotes.Text
Close #1
MsgBox "Notes has been successfully saved", vbInformation
Saved = True
End Sub

Private Sub cmdScheduler_Click()
frmScheduler.Show 1
End Sub

Private Sub cmdViewLog_Click()
PathToDoc = App.Path & "\Notes.log"
ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 8
End Sub

Private Sub Form_Activate()
Saved = True
End Sub

Private Sub Form_Load()

Dim Text As String
Dim Output As String
Open App.Path & "\Notes.log" For Input As #1
While Not EOF(1) = True
    Line Input #1, Text
    Output = Output & Text & vbCrLf
Wend
txtNotes.Text = Output
Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Saved = False Then
    x = MsgBox("Would you like to save changes?", vbExclamation + vbYesNoCancel)
    If x = vbYes Then
        cmdSave_Click
    ElseIf x = vbCancel Then
        Cancel = 1
    End If
End If
End Sub

Private Sub txtNotes_Change()
Saved = False
End Sub

Private Sub txtNotes_GotFocus()
txtNotes.SelStart = Len(txtNotes)
End Sub
